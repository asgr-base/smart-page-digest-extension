/**
 * @file Side panel logic for Smart Page Digest
 */
(function () {
  'use strict';

  var config = globalThis.OCS_CONFIG;

  // --- State ---
  var currentSettings = Object.assign({}, config.DEFAULT_SETTINGS);
  var extractedPageData = null;
  var isSummarizing = false;
  var currentTabId = null;
  var tabCache = {}; // { tabId: { pageData, summaryType, tldr, keyPoints } }

  // --- DOM References ---
  var el = {
    summarizeBtn: document.getElementById('summarizeBtn'),
    settingsBtn: document.getElementById('settingsBtn'),
    langSelect: document.getElementById('langSelect'),
    statusBanner: document.getElementById('statusBanner'),
    pageInfo: document.getElementById('pageInfo'),
    pageTitle: document.getElementById('pageTitle'),
    charCount: document.getElementById('charCount'),
    resultsArea: document.getElementById('resultsArea'),
    tldrSection: document.getElementById('tldrSection'),
    tldrContent: document.getElementById('tldrContent'),
    keyPointsSection: document.getElementById('keyPointsSection'),
    keyPointsContent: document.getElementById('keyPointsContent'),
    customPromptInput: document.getElementById('customPromptInput'),
    customPromptBtn: document.getElementById('customPromptBtn'),
    customPromptDetails: document.getElementById('customPromptDetails'),
    chatHistory: document.getElementById('chatHistory'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    loadingText: document.getElementById('loadingText'),
    toast: document.getElementById('toast'),
    shortcutHint: document.getElementById('shortcutHint'),
    quizSection: document.getElementById('quizSection'),
    generateQuizBtn: document.getElementById('generateQuizBtn'),
    quizContent: document.getElementById('quizContent'),
    copyPageInfoBtn: document.getElementById('copyPageInfoBtn'),
    speechSpeedBtn: document.getElementById('speechSpeedBtn'),
    voiceSelect: document.getElementById('voiceSelect')
  };

  var SPEECH_SPEEDS = [0.5, 0.75, 1, 1.25, 1.5, 1.75, 2, 2.5, 3];
  var speechSpeedIndex = 2; // default 1x
  var savedVoiceURI = ''; // persisted voice selection

  // --- Initialization ---
  async function init() {
    await loadSettings();
    applyI18n();
    await checkApiAvailability();
    bindEvents();
    showShortcutHint();

    // Track current tab and check accessibility
    var tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tabs[0]) {
      currentTabId = tabs[0].id;
      var accessible = await updateTabAccessibility(tabs[0].id);
      if (!accessible) return; // Don't auto-summarize inaccessible pages
    }

    // Always auto-summarize on panel first open (regardless of autoSummarize setting).
    // The autoSummarize setting controls tab-switch behavior only.
    if (!el.summarizeBtn.disabled) {
      handleSummarize();
    }
  }

  function showShortcutHint() {
    chrome.commands.getAll(function (commands) {
      var cmd = commands.find(function (c) { return c.name === '_execute_action'; });
      if (cmd && cmd.shortcut) {
        var label = chrome.i18n.getMessage('shortcutHint') || 'Shortcut:';
        clearElement(el.shortcutHint);
        el.shortcutHint.appendChild(document.createTextNode(label + ' '));
        var kbd = document.createElement('kbd');
        kbd.textContent = cmd.shortcut;
        el.shortcutHint.appendChild(kbd);
      }
    });
  }

  // --- Settings ---
  async function loadSettings() {
    try {
      var data = await chrome.storage.sync.get(config.STORAGE_KEY);
      if (data[config.STORAGE_KEY]) {
        currentSettings = Object.assign({}, config.DEFAULT_SETTINGS, data[config.STORAGE_KEY]);
      }
    } catch (err) {
      console.warn('Failed to load settings:', err);
    }
    el.langSelect.value = currentSettings.outputLanguage;

    // Restore speech speed
    if (currentSettings.speechSpeed) {
      var idx = SPEECH_SPEEDS.indexOf(currentSettings.speechSpeed);
      if (idx >= 0) speechSpeedIndex = idx;
    }
    updateSpeedButton();

    // Restore saved voice
    if (currentSettings.voiceURI) {
      savedVoiceURI = currentSettings.voiceURI;
    }
    populateVoiceList();
  }

  function updateSpeedButton() {
    var speed = SPEECH_SPEEDS[speechSpeedIndex];
    el.speechSpeedBtn.textContent = speed === 1 ? '1x' : speed + 'x';
  }

  function cycleSpeechSpeed() {
    speechSpeedIndex = (speechSpeedIndex + 1) % SPEECH_SPEEDS.length;
    updateSpeedButton();
    saveSetting('speechSpeed', SPEECH_SPEEDS[speechSpeedIndex]);
  }

  function populateVoiceList() {
    var voices = speechSynthesis.getVoices();
    if (voices.length === 0) return; // voices not loaded yet

    var lang = getOutputLanguage(extractedPageData ? extractedPageData.lang : null);
    var langMap = { ja: 'ja', en: 'en', es: 'es' };
    var filterLang = langMap[lang] || 'en';

    // Filter voices by current language
    var matchingVoices = voices.filter(function (v) {
      return v.lang.startsWith(filterLang);
    });

    // Clear existing options
    clearElement(el.voiceSelect);

    // Default option
    var defaultOpt = document.createElement('option');
    defaultOpt.value = '';
    defaultOpt.textContent = chrome.i18n.getMessage('voiceDefault') || 'Default';
    el.voiceSelect.appendChild(defaultOpt);

    // Add matching voices
    for (var i = 0; i < matchingVoices.length; i++) {
      var v = matchingVoices[i];
      var opt = document.createElement('option');
      opt.value = v.voiceURI;
      // Show friendly name (remove language suffix for brevity)
      var name = v.name.replace(/\s*\(.*\)$/, '');
      opt.textContent = name;
      if (v.voiceURI === savedVoiceURI) {
        opt.selected = true;
      }
      el.voiceSelect.appendChild(opt);
    }
  }

  async function saveSetting(key, value) {
    currentSettings[key] = value;
    var storageData = {};
    storageData[config.STORAGE_KEY] = currentSettings;
    try {
      await chrome.storage.sync.set(storageData);
    } catch (err) {
      console.warn('Failed to save setting:', err);
    }
  }

  // --- i18n ---
  function applyI18n() {
    document.querySelectorAll('[data-i18n]').forEach(function (elem) {
      var key = elem.getAttribute('data-i18n');
      var msg = chrome.i18n.getMessage(key);
      if (msg) elem.textContent = msg;
    });
    document.querySelectorAll('[data-i18n-placeholder]').forEach(function (elem) {
      var key = elem.getAttribute('data-i18n-placeholder');
      var msg = chrome.i18n.getMessage(key);
      if (msg) elem.placeholder = msg;
    });
    document.querySelectorAll('[data-i18n-title]').forEach(function (elem) {
      var key = elem.getAttribute('data-i18n-title');
      var msg = chrome.i18n.getMessage(key);
      if (msg) elem.title = msg;
    });
  }

  // --- API Availability ---
  async function checkApiAvailability() {
    if (!('Summarizer' in self)) {
      showStatus('error', chrome.i18n.getMessage('summarizerNotSupported') ||
        'Summarizer API is not supported. Please use Chrome 138 or later.');
      el.summarizeBtn.disabled = true;
      return;
    }

    try {
      var availability = await Summarizer.availability({
        outputLanguage: currentSettings.outputLanguage === 'auto' ? 'en' : currentSettings.outputLanguage
      });
      if (availability === 'unavailable') {
        showStatus('error', chrome.i18n.getMessage('summarizerUnavailable') ||
          'AI model is not available. Please check chrome://flags.');
        el.summarizeBtn.disabled = true;
        return;
      }
      if (availability === 'downloadable') {
        showStatus('info', chrome.i18n.getMessage('modelDownloading') ||
          'AI model will be downloaded on first use.');
      }
    } catch (err) {
      showStatus('warning', 'Could not check API availability: ' + err.message);
    }

    // Check Prompt API for custom prompts and quiz
    if (!('LanguageModel' in self)) {
      el.customPromptInput.disabled = true;
      el.customPromptBtn.disabled = true;
      el.generateQuizBtn.disabled = true;
    }

    // Diagnostic: log available AI APIs for debugging translation support
    console.log('[OCS] API Detection:', {
      Summarizer: 'Summarizer' in self,
      LanguageModel: 'LanguageModel' in self,
      Translator: 'Translator' in self,
      'translation.createTranslator': ('translation' in self && typeof self.translation?.createTranslator === 'function'),
      'ai.translator': ('ai' in self && 'translator' in self.ai),
      isTranslatorAvailable: isTranslatorAvailable()
    });
  }

  // --- Text Extraction ---
  function extractText() {
    return new Promise(function (resolve, reject) {
      chrome.runtime.sendMessage(
        { type: config.MESSAGES.EXTRACT_TEXT },
        function (response) {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          if (!response || !response.success) {
            var errMsg = response ? (response.error + (response.message ? ': ' + response.message : '')) : 'extractionFailed';
            reject(new Error(errMsg));
            return;
          }
          resolve(response.data);
        }
      );
    });
  }

  // Ensure page data is extracted (lazy extraction for quiz/custom prompt).
  // Returns extractedPageData or null if extraction fails.
  async function ensurePageData() {
    if (extractedPageData) return extractedPageData;
    try {
      extractedPageData = await extractText();
      if (extractedPageData && extractedPageData.text) {
        showPageInfo(extractedPageData);
        return extractedPageData;
      }
      extractedPageData = null;
    } catch (e) {
      console.warn('[OCS] ensurePageData extraction failed:', e.message);
      extractedPageData = null;
    }
    return null;
  }

  // --- Determine output language ---
  function getOutputLanguage(pageLang) {
    var selected = el.langSelect.value;
    if (selected !== 'auto') return selected;
    if (pageLang) {
      var base = pageLang.split('-')[0].toLowerCase();
      if (base === 'ja' || base === 'en' || base === 'es') return base;
    }
    return 'en';
  }

  // --- URL Accessibility ---
  function isAccessibleUrl(url) {
    if (!url) return false;
    return url.startsWith('http://') || url.startsWith('https://');
  }

  // Check active tab URL and update button state accordingly
  async function updateTabAccessibility(tabId) {
    try {
      var tab = await chrome.tabs.get(tabId);
      if (!isAccessibleUrl(tab.url)) {
        el.summarizeBtn.disabled = true;
        showStatus('info', chrome.i18n.getMessage('inaccessiblePage') ||
          'Cannot access this page. Summary is not available for browser internal pages.');
        return false;
      }
    } catch (err) {
      // Tab may not exist anymore
      el.summarizeBtn.disabled = true;
      return false;
    }
    el.summarizeBtn.disabled = false;
    hideStatus();
    return true;
  }

  // --- Tab Cache ---
  function saveToCache(tabId, pageData, summaryType, tldrText, keyPointsText) {
    tabCache[tabId] = {
      pageData: pageData,
      summaryType: summaryType,
      tldr: tldrText || null,
      keyPoints: keyPointsText || null
    };
  }

  function restoreFromCache(tabId) {
    var cached = tabCache[tabId];
    if (!cached) return false;

    extractedPageData = cached.pageData;
    showPageInfo(cached.pageData);
    showSections(cached.summaryType);

    if (cached.tldr) {
      renderMarkdownSafe(cached.tldr, el.tldrContent);
    }
    if (cached.keyPoints) {
      renderKeyPointsWithImportance(cached.keyPoints, el.keyPointsContent);
    }

    // Restore or clear quiz content
    if (cached.quizHtml) {
      el.quizContent.innerHTML = cached.quizHtml;
      rebindQuizCards();
    } else {
      clearElement(el.quizContent);
    }

    // Restore or clear chat history
    if (cached.chatHtml) {
      el.chatHistory.innerHTML = cached.chatHtml;
    } else {
      clearElement(el.chatHistory);
    }

    return true;
  }

  function showFreshState() {
    extractedPageData = null;
    stopReadAloud();
    clearResults();
    el.pageInfo.classList.add('ocs-hidden');
    clearElement(el.chatHistory);
    clearElement(el.quizContent);
  }

  // Save quiz and chat state for a tab before switching away.
  function saveCurrentTabQuizChat(tabId) {
    if (!tabId) return;
    var hasQuiz = el.quizContent.childNodes.length > 0;
    var hasChat = el.chatHistory.childNodes.length > 0;
    if (!hasQuiz && !hasChat) return;

    if (!tabCache[tabId]) {
      // Create a minimal cache entry for quiz/chat-only tabs (no prior summarization)
      tabCache[tabId] = {
        pageData: extractedPageData,
        summaryType: currentSettings.summaryType,
        tldr: el.tldrContent.getAttribute('data-raw-text') || null,
        keyPoints: el.keyPointsContent.getAttribute('data-raw-text') || null
      };
    }
    tabCache[tabId].quizHtml = hasQuiz ? el.quizContent.innerHTML : null;
    tabCache[tabId].chatHtml = hasChat ? el.chatHistory.innerHTML : null;
  }

  // Re-attach click-to-reveal handlers on quiz cards after innerHTML restore.
  function rebindQuizCards() {
    var cards = el.quizContent.querySelectorAll('.ocs-quiz-card');
    for (var i = 0; i < cards.length; i++) {
      (function (card) {
        var qDiv = card.querySelector('.ocs-quiz-question');
        var aDiv = card.querySelector('.ocs-quiz-answer');
        if (qDiv && aDiv) {
          qDiv.addEventListener('click', function () {
            aDiv.classList.toggle('ocs-revealed');
          });
        }
      })(cards[i]);
    }
  }

  async function switchToTab(tabId) {
    if (tabId === currentTabId) return;

    // Save quiz and chat state for the tab we're leaving
    saveCurrentTabQuizChat(currentTabId);

    currentTabId = tabId;

    if (isSummarizing) return; // Don't interrupt active summarization

    // Check if the new tab's URL is accessible before doing anything
    var accessible = await updateTabAccessibility(tabId);
    if (!accessible) {
      showFreshState();
      return;
    }

    if (!restoreFromCache(tabId)) {
      showFreshState();
      // Auto-summarize for new tab if setting is on
      if (currentSettings.autoSummarize && !el.summarizeBtn.disabled) {
        handleSummarize();
      }
    }
  }

  // --- Summarization ---
  var pendingTranslator = null; // Pre-created translator (within user gesture context)
  var untranslatedResults = null; // Stores English results when translation model download needed

  async function handleSummarize(hasUserGesture) {
    if (isSummarizing) return;
    isSummarizing = true;
    stopReadAloud();
    showLoading(true, chrome.i18n.getMessage('extractingText') || 'Extracting page content...');
    clearResults();
    untranslatedResults = null; // Reset translation retry state

    // Capture the tab ID at the start of summarization
    var tabId = currentTabId;

    // Pre-create Translator within user gesture context to avoid NotAllowedError.
    // Chrome requires a user gesture for model downloads ("downloadable"/"downloading").
    // Only attempt when called from a direct user click, not from auto-summarize.
    var userSelectedLang = el.langSelect.value;
    var preOutputLang = userSelectedLang !== 'auto' ? userSelectedLang : currentSettings.outputLanguage;
    if (hasUserGesture && preOutputLang !== 'en' && isTranslatorAvailable()) {
      pendingTranslator = createTranslator(preOutputLang);
    } else {
      pendingTranslator = null;
    }

    try {
      extractedPageData = await extractText();
      if (!extractedPageData || !extractedPageData.text) {
        throw new Error('extractionFailed');
      }
      showPageInfo(extractedPageData);

      var outputLang = getOutputLanguage(extractedPageData.lang);
      var pageLangBase = extractedPageData.lang
        ? extractedPageData.lang.split('-')[0].toLowerCase() : null;
      // Only use cross-language translation path when Translator API is available.
      // Without it, fall back to Summarizer API's native outputLanguage parameter.
      var needsCrossLang = userSelectedLang !== 'auto' &&
        pageLangBase !== outputLang && isTranslatorAvailable();

      showLoading(true, chrome.i18n.getMessage('summarizing') || 'Generating summary...');

      var sharedContext = 'Page title: ' + extractedPageData.title;

      var results;
      if (needsCrossLang && ('LanguageModel' in self)) {
        results = await summarizeWithPromptApi(extractedPageData.text, outputLang, sharedContext);
      } else {
        results = await summarizeWithSummarizerApi(extractedPageData.text, sharedContext, outputLang);
      }

      // Show translation download banner if translation model wasn't available
      if (untranslatedResults) {
        showTranslationDownloadBanner();
      }

      // Cache results for this tab
      if (tabId) {
        saveToCache(tabId, extractedPageData, currentSettings.summaryType, results.tldr, results.keyPoints);
      }

    } catch (err) {
      showStatus('error', getErrorMessage(err));
    } finally {
      showLoading(false);
      isSummarizing = false;
      // Clean up pre-created translator
      if (pendingTranslator) {
        pendingTranslator.then(function (t) {
          if (t && t.destroy) t.destroy();
        }).catch(function () {});
        pendingTranslator = null;
      }
    }
  }

  // --- Summarizer API (TL;DR) + Prompt API (Key Points with importance) ---
  async function summarizeWithSummarizerApi(text, sharedContext, outputLang) {
    var summaryType = currentSettings.summaryType;
    var results = { tldr: null, keyPoints: null };

    var createOptions = function (type) {
      return {
        type: type,
        format: 'markdown',
        length: currentSettings.summaryLength,
        expectedInputLanguages: ['en', 'ja', 'es'],
        outputLanguage: outputLang,
        sharedContext: sharedContext,
        monitor: function (m) {
          m.addEventListener('downloadprogress', function (e) {
            var pct = Math.round((e.loaded / e.total) * 100);
            el.loadingText.textContent = 'Downloading model: ' + pct + '%';
          });
        }
      };
    };

    showSections(summaryType);
    showLoading(false);
    el.summarizeBtn.disabled = true;

    // Run sequentially (Gemini Nano handles one request at a time)
    try {
      if (summaryType === 'both' || summaryType === 'tldr') {
        showSectionLoading(el.tldrContent);
        results.tldr = await runSummarizer(createOptions('tldr'), text, el.tldrContent);
      }
      if (summaryType === 'both' || summaryType === 'key-points') {
        showSectionLoading(el.keyPointsContent);
        // Use Prompt API for importance-tagged key points when possible.
        // Generate in English (Nano's strongest) then translate if needed.
        if ('LanguageModel' in self) {
          var promptText = text.length > config.TEXT_LIMITS.PROMPT_API_MAX_CHARS
            ? text.substring(0, config.TEXT_LIMITS.PROMPT_API_MAX_CHARS) + config.TEXT_LIMITS.TRUNCATION_SUFFIX
            : text;
          var kpPrompt = 'Summarize the following text as key points (3-7 bullet points).' +
            '\nEach bullet MUST start with an importance tag: [HIGH], [MEDIUM], or [LOW].' +
            '\nFormat: - [HIGH] Most important point here\n\n' +
            sharedContext +
            '\n\nText:\n' + promptText;
          try {
            results.keyPoints = await runPromptApi(
              'You are a summarization assistant.', kpPrompt, el.keyPointsContent, 'en'
            );
            // Translate key points if output is non-English, preserving importance tags.
            // Skip if Nano already generated non-English (ignoring outputLanguage:'en').
            if (outputLang !== 'en' && results.keyPoints && isTranslatorAvailable()) {
              if (looksLikeEnglish(results.keyPoints)) {
                showTranslatingIndicator(el.keyPointsContent);
                var englishKp = results.keyPoints;
                results.keyPoints = await translateKeyPointsIfNeeded(results.keyPoints, outputLang);
                if (results.keyPoints === englishKp) {
                  if (!untranslatedResults) untranslatedResults = { outputLang: outputLang };
                  untranslatedResults.keyPoints = englishKp;
                }
              } else {
                console.log('[OCS] Key points (Summarizer path) already non-English, skipping translation');
              }
            }
            results.keyPoints = repairImportanceTags(results.keyPoints);
            renderKeyPointsWithImportance(results.keyPoints, el.keyPointsContent);
          } catch (err) {
            console.warn('Prompt API key points failed, falling back to Summarizer:', err);
            results.keyPoints = await runSummarizer(createOptions('key-points'), text, el.keyPointsContent);
          }
        } else {
          // No Prompt API: Summarizer API without importance tags
          results.keyPoints = await runSummarizer(createOptions('key-points'), text, el.keyPointsContent);
        }
      }
    } finally {
      el.summarizeBtn.disabled = false;
    }

    return results;
  }

  async function runSummarizer(options, text, targetElement) {
    // Build attempts: full text, then progressively smaller on "too large" errors
    var attempts = [text];
    if (text.length > 6000) {
      attempts.push(text.substring(0, Math.floor(text.length / 2)) + config.TEXT_LIMITS.TRUNCATION_SUFFIX);
    }
    if (text.length > 3000) {
      attempts.push(text.substring(0, Math.floor(text.length / 4)) + config.TEXT_LIMITS.TRUNCATION_SUFFIX);
    }

    for (var i = 0; i < attempts.length; i++) {
      var summarizer = await Summarizer.create(options);
      try {
        var summary = await summarizer.summarize(attempts[i]);
        if (summary) {
          renderMarkdownSafe(summary, targetElement);
          return summary;
        } else {
          targetElement.textContent = '(No summary generated)';
          return null;
        }
      } catch (err) {
        // Retry with smaller text if "too large" and more attempts remain
        if (/too large/i.test(err.message || '') && i < attempts.length - 1) {
          continue;
        }
        targetElement.textContent = 'Error: ' + err.message;
        return null;
      } finally {
        summarizer.destroy();
      }
    }
    targetElement.textContent = '(No summary generated)';
    return null;
  }

  // --- Prompt API (cross-language: summarize in English → translate) ---
  async function summarizeWithPromptApi(text, outputLang, sharedContext) {
    if (!('LanguageModel' in self)) {
      showStatus('error', 'LanguageModel API is not available.');
      return { tldr: null, keyPoints: null };
    }

    var summaryType = currentSettings.summaryType;
    var results = { tldr: null, keyPoints: null };

    // Cross-language strategy: always summarize in English (Gemini Nano's
    // strongest language), then translate to target language.
    // Uses Translator API if available, falls back to Prompt API translation.
    var canTranslate = outputLang !== 'en';
    var systemPrompt = 'You are a summarization assistant.';

    showSections(summaryType);
    showLoading(false);
    el.summarizeBtn.disabled = true;

    var promptText = text.length > config.TEXT_LIMITS.PROMPT_API_MAX_CHARS
      ? text.substring(0, config.TEXT_LIMITS.PROMPT_API_MAX_CHARS) + config.TEXT_LIMITS.TRUNCATION_SUFFIX
      : text;

    try {
      if (summaryType === 'both' || summaryType === 'tldr') {
        showSectionLoading(el.tldrContent);
        var tldrPrompt = 'Write a brief TL;DR summary (1-5 sentences).\n\n' +
          sharedContext + '\n\nText:\n' + promptText;
        try {
          results.tldr = await runPromptApi(systemPrompt, tldrPrompt, el.tldrContent, 'en');
          if (canTranslate && results.tldr) {
            // Gemini Nano may ignore outputLanguage:'en' when input is non-English.
            // If the response is already non-English, skip translation.
            if (looksLikeEnglish(results.tldr)) {
              showTranslatingIndicator(el.tldrContent);
              var englishTldr = results.tldr;
              results.tldr = await translateIfNeeded(results.tldr, outputLang);
              if (results.tldr === englishTldr) {
                if (!untranslatedResults) untranslatedResults = { outputLang: outputLang };
                untranslatedResults.tldr = englishTldr;
              }
            } else {
              console.log('[OCS] TL;DR already non-English, skipping translation');
            }
            renderMarkdownSafe(results.tldr, el.tldrContent);
          }
        } catch (err) {
          el.tldrContent.textContent = 'Error: ' + (err.message || 'Failed to generate TL;DR');
        }
      }
      if (summaryType === 'both' || summaryType === 'key-points') {
        showSectionLoading(el.keyPointsContent);
        var kpPrompt = 'Summarize the following text as key points (3-7 bullet points).' +
          '\nEach bullet MUST start with an importance tag: [HIGH], [MEDIUM], or [LOW].' +
          '\nFormat: - [HIGH] Most important point here\n\n' +
          sharedContext + '\n\nText:\n' + promptText;
        try {
          results.keyPoints = await runPromptApi(systemPrompt, kpPrompt, el.keyPointsContent, 'en');
          if (canTranslate && results.keyPoints) {
            // Same check: skip translation if Nano already generated non-English
            if (looksLikeEnglish(results.keyPoints)) {
              showTranslatingIndicator(el.keyPointsContent);
              var englishKp = results.keyPoints;
              results.keyPoints = await translateKeyPointsIfNeeded(results.keyPoints, outputLang);
              if (results.keyPoints === englishKp) {
                if (!untranslatedResults) untranslatedResults = { outputLang: outputLang };
                untranslatedResults.keyPoints = englishKp;
              }
            } else {
              console.log('[OCS] Key points already non-English, skipping translation');
            }
          }
          results.keyPoints = repairImportanceTags(results.keyPoints);
          renderKeyPointsWithImportance(results.keyPoints, el.keyPointsContent);
        } catch (err) {
          el.keyPointsContent.textContent = 'Error: ' + (err.message || 'Failed to generate key points');
        }
      }
    } finally {
      el.summarizeBtn.disabled = false;
    }

    return results;
  }

  async function runPromptApi(systemPrompt, prompt, targetElement, outputLanguage) {
    var lang = outputLanguage || 'en';
    var session = await LanguageModel.create({
      systemPrompt: systemPrompt,
      expectedInputs: [{ type: 'text' }],
      expectedOutputs: [{ type: 'text', languages: [lang] }],
      outputLanguage: lang
    });
    var fullText = '';
    try {
      var stream = session.promptStreaming(prompt);
      // Chrome's promptStreaming returns cumulative chunks (pre-131) or delta chunks (131+).
      // Auto-detect on the second chunk: if it starts with fullText, it's cumulative.
      var isCumulative = null;
      for await (var chunk of stream) {
        if (isCumulative === null && fullText.length > 0) {
          isCumulative = chunk.startsWith(fullText);
        }
        if (isCumulative) {
          fullText = chunk;
        } else {
          fullText += chunk;
        }
        renderMarkdownSafe(fullText, targetElement);
      }
    } finally {
      session.destroy();
    }
    return fullText;
  }

  // --- Language Detection (heuristic) ---
  // Check if text appears to be primarily English.
  // Used to skip translation when Prompt API ignores outputLanguage:'en'
  // and generates in the input language instead.
  function looksLikeEnglish(text) {
    if (!text) return true;
    // Count non-ASCII characters (Japanese, CJK, accented Latin, etc.)
    var sample = text.substring(0, 500);
    var nonAscii = 0;
    for (var i = 0; i < sample.length; i++) {
      if (sample.charCodeAt(i) > 127) nonAscii++;
    }
    // If more than 15% non-ASCII, it's likely not English
    return (nonAscii / sample.length) < 0.15;
  }

  // --- Translation (Chrome Built-in Translator API) ---
  function isTranslatorAvailable() {
    return ('Translator' in self) ||
      ('translation' in self && typeof self.translation.createTranslator === 'function') ||
      ('ai' in self && 'translator' in self.ai);
  }

  // Create a Translator instance. Call within user gesture context to allow model downloads.
  function createTranslator(targetLang) {
    if ('Translator' in self) {
      return Translator.create({
        sourceLanguage: 'en',
        targetLanguage: targetLang,
        monitor: function (m) {
          m.addEventListener('downloadprogress', function (e) {
            var pct = Math.round((e.loaded / e.total) * 100);
            el.loadingText.textContent = 'Downloading translation model: ' + pct + '%';
          });
        }
      });
    } else if ('translation' in self && typeof self.translation.createTranslator === 'function') {
      return self.translation.createTranslator({
        sourceLanguage: 'en',
        targetLanguage: targetLang
      });
    } else if ('ai' in self && 'translator' in self.ai) {
      return self.ai.translator.create({
        sourceLanguage: 'en',
        targetLanguage: targetLang
      });
    }
    return Promise.reject(new Error('No Translator API available'));
  }

  async function translateIfNeeded(text, targetLang) {
    if (!text || targetLang === 'en') return text;

    // Use pre-created translator (created within user gesture context in handleSummarize)
    // to avoid NotAllowedError when model needs downloading.
    // If pendingTranslator is null (auto-summarize / tab switch), try creating on-demand.
    // This works when the model is already downloaded (no user gesture needed).
    try {
      var translator = pendingTranslator ? await pendingTranslator : null;
      if (!translator) {
        try {
          translator = await createTranslator(targetLang);
        } catch (e) {
          // NotAllowedError = model not downloaded yet, need user gesture
          console.log('[OCS] On-demand translator creation failed (model may need download):', e.message);
          return text;
        }
      }
      try {
        var translated = await translator.translate(text);
        if (!pendingTranslator && translator.destroy) translator.destroy();
        if (translated && translated.trim()) return translated;
      } catch (err) {
        if (!pendingTranslator && translator.destroy) translator.destroy();
        console.warn('Translator.translate() failed:', err);
      }
    } catch (err) {
      console.warn('Translator API failed:', err);
    }

    return text;
  }

  // Translate key points using item-by-item strategy via pre-created translator.
  // Falls back to on-demand translator creation when pendingTranslator is null
  // (e.g., auto-summarize, tab switch — model already downloaded).
  async function translateKeyPointsIfNeeded(keyPointsText, targetLang) {
    if (!keyPointsText || targetLang === 'en') return keyPointsText;
    try {
      var translator = pendingTranslator ? await pendingTranslator : null;
      if (!translator) {
        try {
          translator = await createTranslator(targetLang);
        } catch (e) {
          console.log('[OCS] On-demand translator creation failed for key points:', e.message);
          return keyPointsText;
        }
      }
      try {
        var translated = await translateKeyPointsPreservingFormat(translator, keyPointsText);
        if (!pendingTranslator && translator.destroy) translator.destroy();
        if (translated && translated.trim()) return translated;
      } catch (err) {
        if (!pendingTranslator && translator.destroy) translator.destroy();
        console.warn('Key points translation failed:', err);
      }
    } catch (err) {
      console.warn('Translator API failed for key points:', err);
    }
    return keyPointsText;
  }

  function repairImportanceTags(text) {
    return text
      .replace(/\[(?:高|高い|ハイ|重要|重要度高|High)\]/gi, '[HIGH]')
      .replace(/\[(?:中|中程度|ミディアム|標準|重要度中|Medium)\]/gi, '[MEDIUM]')
      .replace(/\[(?:低|低い|ロー|重要度低|Low)\]/gi, '[LOW]')
      .replace(/\[(?:ALTO|ALTA)\]/gi, '[HIGH]')
      .replace(/\[(?:MEDIO|MEDIA)\]/gi, '[MEDIUM]')
      .replace(/\[(?:BAJO|BAJA)\]/gi, '[LOW]');
  }

  // Translate key points item-by-item to preserve bullet format and importance tags.
  // The Translator API merges lines and translates tags when given the whole block.
  function translateKeyPointsPreservingFormat(translator, keyPointsText) {
    var lines = keyPointsText.split('\n');
    var items = [];
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;
      var match = line.match(/^[-*]\s*\[?(HIGH|MEDIUM|LOW)\]?\s*(.*)/i);
      if (match) {
        items.push({ tag: '[' + match[1].toUpperCase() + ']', text: match[2] });
      }
    }
    if (items.length === 0) {
      // Can't parse structure, translate whole block as fallback
      return translator.translate(keyPointsText);
    }
    // Translate each item's text individually, preserving tags
    var translatedItems = [];
    var chain = Promise.resolve();
    items.forEach(function (item, idx) {
      chain = chain.then(function () {
        return translator.translate(item.text);
      }).then(function (translated) {
        translatedItems[idx] = '- ' + item.tag + ' ' + (translated && translated.trim() ? translated.trim() : item.text);
      });
    });
    return chain.then(function () {
      return translatedItems.join('\n');
    });
  }

  function showTranslatingIndicator(container) {
    clearElement(container);
    var wrapper = document.createElement('div');
    wrapper.className = 'ocs-section-loading';
    var spinner = document.createElement('div');
    spinner.className = 'ocs-section-spinner';
    var label = document.createElement('span');
    label.textContent = chrome.i18n.getMessage('translating') || 'Translating...';
    wrapper.appendChild(spinner);
    wrapper.appendChild(label);
    container.appendChild(wrapper);
  }

  function showTranslationDownloadBanner() {
    // Remove existing banner if any
    var existing = document.getElementById('translationDownloadBanner');
    if (existing) existing.remove();

    var data = untranslatedResults;
    if (!data) return;

    var banner = document.createElement('div');
    banner.id = 'translationDownloadBanner';
    banner.className = 'ocs-translation-banner';

    var msg = document.createElement('span');
    msg.textContent = chrome.i18n.getMessage('translationModelNeeded') ||
      'Translation model needs to be downloaded to translate summaries.';

    var btn = document.createElement('button');
    btn.className = 'ocs-translate-download-btn';
    btn.textContent = chrome.i18n.getMessage('downloadAndTranslate') || 'Download & Translate';

    // Use non-async handler so Translator.create() is called
    // synchronously within user gesture context (before any microtask yields).
    btn.addEventListener('click', function () {
      btn.disabled = true;
      btn.textContent = chrome.i18n.getMessage('translating') || 'Translating...';

      // Save current content so we can restore on failure
      var savedTldrHtml = el.tldrContent.innerHTML;
      var savedKpHtml = el.keyPointsContent.innerHTML;

      // Start translator creation synchronously (user gesture context)
      var translatorPromise = createTranslator(data.outputLang);

      translatorPromise.then(function (translator) {
        var translateWork = Promise.resolve();

        if (data.tldr) {
          showTranslatingIndicator(el.tldrContent);
          translateWork = translateWork.then(function () {
            return translator.translate(data.tldr);
          }).then(function (translatedTldr) {
            console.log('[OCS] Translated TL;DR result:', translatedTldr);
            if (translatedTldr && translatedTldr.trim()) {
              renderMarkdownSafe(translatedTldr, el.tldrContent);
            } else {
              // Fallback: show original English if translation returned empty
              renderMarkdownSafe(data.tldr, el.tldrContent);
            }
          });
        }

        if (data.keyPoints) {
          translateWork = translateWork.then(function () {
            showTranslatingIndicator(el.keyPointsContent);
            return translateKeyPointsPreservingFormat(translator, data.keyPoints);
          }).then(function (translatedKp) {
            console.log('[OCS] Translated key points result:', translatedKp);
            if (translatedKp && translatedKp.trim()) {
              renderKeyPointsWithImportance(translatedKp, el.keyPointsContent);
            }
            // Fallback: if nothing was rendered, show as plain markdown
            if (!el.keyPointsContent.querySelector('ul') ||
                el.keyPointsContent.querySelector('ul').childNodes.length === 0) {
              console.log('[OCS] Key points render empty, falling back to markdown');
              renderMarkdownSafe(translatedKp || data.keyPoints, el.keyPointsContent);
            }
          });
        }

        return translateWork.then(function () {
          if (translator.destroy) translator.destroy();
          banner.remove();
          untranslatedResults = null;
          showToast(chrome.i18n.getMessage('translationComplete') || 'Translation complete');
        });
      }).catch(function (err) {
        console.error('Translation download/translate failed:', err);
        // Restore original content
        el.tldrContent.innerHTML = savedTldrHtml;
        el.keyPointsContent.innerHTML = savedKpHtml;
        btn.disabled = false;
        btn.textContent = chrome.i18n.getMessage('downloadAndTranslate') || 'Download & Translate';
        showStatus('error', 'Translation failed: ' + err.message);
      });
    });

    banner.appendChild(msg);
    banner.appendChild(btn);
    el.resultsArea.insertBefore(banner, el.resultsArea.firstChild);
  }

  // --- Render Key Points with Importance ---
  function renderKeyPointsWithImportance(text, container) {
    container.setAttribute('data-raw-text', text);
    clearElement(container);
    // Normalize: Nano may output bullets without line breaks (e.g., all on one line).
    // Also handle optional **bold** markers around tags (e.g., "- **[HIGH]**").
    var normalized = text.replace(/\s+-\s+\*{0,2}\[(HIGH|MEDIUM|LOW)\]\*{0,2}/gi, '\n- [$1]');
    var lines = normalized.split('\n');
    var ul = document.createElement('ul');

    for (var i = 0; i < lines.length; i++) {
      var trimmed = lines[i].trim();
      if (!trimmed) continue;

      // Match bullet point (- or *)
      var listMatch = trimmed.match(/^[-*]\s+(.+)/);
      var content;
      if (listMatch) {
        content = listMatch[1];
      } else if (/^\*{0,2}\s*\[(HIGH|MEDIUM|LOW)\]/i.test(trimmed)) {
        // Line with importance tag but no bullet prefix
        content = trimmed;
      } else {
        continue;
      }

      var importance = null;
      // Match importance tag, handling optional bold markers (**) around it
      var importanceMatch = content.match(/^(\*{0,2})\s*\[(HIGH|MEDIUM|LOW)\]\s*(\*{0,2})\s*/i);
      if (importanceMatch) {
        importance = importanceMatch[2].toLowerCase();
        content = content.substring(importanceMatch[0].length);
        // If opening ** was consumed with tag, re-wrap remaining bold section
        if (importanceMatch[1] === '**' && importanceMatch[3] === '') {
          var closeIdx = content.indexOf('**');
          if (closeIdx >= 0) {
            content = '**' + content.substring(0, closeIdx) + '**' + content.substring(closeIdx + 2);
          }
        }
      }

      var li = document.createElement('li');
      if (importance) {
        li.setAttribute('data-importance', importance);
        var badge = document.createElement('span');
        badge.className = 'ocs-importance-badge ocs-' + importance;
        var labels = { high: 'HIGH', medium: 'MED', low: 'LOW' };
        badge.textContent = labels[importance] || importance.toUpperCase();
        li.appendChild(badge);
      }
      appendInlineFormatted(li, content);
      ul.appendChild(li);
    }

    if (ul.childNodes.length > 0) {
      container.appendChild(ul);
    } else {
      // Fallback: importance tags not found (e.g., Nano generated non-English without tags).
      // Render as plain markdown instead of leaving container empty.
      renderMarkdownSafe(text, container);
    }
  }

  // --- Quiz Generation (Roediger & Karpicke 2006: retrieval practice) ---
  async function handleGenerateQuiz() {
    if (!('LanguageModel' in self)) return;

    el.generateQuizBtn.disabled = true;
    clearElement(el.quizContent);
    showSectionLoading(el.quizContent);

    // Extract page data if not already available (e.g., quiz before summarize)
    var pageData = await ensurePageData();
    if (!pageData) {
      clearElement(el.quizContent);
      el.quizContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Not enough text content.';
      el.generateQuizBtn.disabled = false;
      return;
    }

    var pageLang = extractedPageData.lang || null;
    var resolvedLang = getOutputLanguage(pageLang);
    var quizLangEN = (el.langSelect.value !== 'auto') ?
      (config.LANGUAGE_NAMES_FOR_PROMPT[resolvedLang] || 'English') : 'the same language as the text';

    try {
      var session = await LanguageModel.create({
        systemPrompt: 'You are a quiz generator that always writes in ' + quizLangEN + '.',
        expectedInputs: [{ type: 'text' }],
        expectedOutputs: [{ type: 'text', languages: [resolvedLang] }],
        outputLanguage: resolvedLang
      });
      try {
        var prompt = 'Generate exactly 3 comprehension questions with short answers in ' + quizLangEN + '.\n' +
          'Format each Q&A pair on its own line like this:\n' +
          'Q1: [question]\n' +
          'A1: [answer]\n' +
          'Q2: [question]\n' +
          'A2: [answer]\n' +
          'Q3: [question]\n' +
          'A3: [answer]\n\n' +
          'Text:\n' + extractedPageData.text.substring(0, 5000);

        var result = await session.prompt(prompt);
        if (result) {
          renderQuiz(result);
        }
      } finally {
        session.destroy();
      }
    } catch (err) {
      clearElement(el.quizContent);
      el.quizContent.textContent = 'Quiz generation failed: ' + err.message;
    } finally {
      el.generateQuizBtn.disabled = false;
    }
  }

  function renderQuiz(quizText) {
    clearElement(el.quizContent);
    var lines = quizText.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
    var questions = [];
    var currentQ = null;

    for (var i = 0; i < lines.length; i++) {
      var qMatch = lines[i].match(/^Q\d+[:：]\s*(.+)/i);
      var aMatch = lines[i].match(/^A\d+[:：]\s*(.+)/i);

      if (qMatch) {
        currentQ = { question: qMatch[1], answer: null };
        questions.push(currentQ);
      } else if (aMatch && currentQ) {
        currentQ.answer = aMatch[1];
        currentQ = null;
      }
    }

    if (questions.length === 0) {
      el.quizContent.textContent = 'Could not parse quiz format.';
      return;
    }

    var answerLabel = chrome.i18n.getMessage('quizAnswerLabel') || 'Answer';
    var tapHint = chrome.i18n.getMessage('quizTapToReveal') || 'Tap to reveal answer';

    for (var j = 0; j < questions.length; j++) {
      var card = document.createElement('div');
      card.className = 'ocs-quiz-card';

      var questionDiv = document.createElement('div');
      questionDiv.className = 'ocs-quiz-question';
      questionDiv.title = tapHint;

      var numberSpan = document.createElement('span');
      numberSpan.className = 'ocs-quiz-number';
      numberSpan.textContent = String(j + 1);
      questionDiv.appendChild(numberSpan);

      var qText = document.createElement('span');
      qText.textContent = questions[j].question;
      questionDiv.appendChild(qText);

      var answerDiv = document.createElement('div');
      answerDiv.className = 'ocs-quiz-answer';

      var aLabel = document.createElement('div');
      aLabel.className = 'ocs-quiz-answer-label';
      aLabel.textContent = answerLabel;
      answerDiv.appendChild(aLabel);

      var aText = document.createElement('div');
      aText.textContent = questions[j].answer || '—';
      answerDiv.appendChild(aText);

      // Toggle reveal
      (function (qDiv, aDiv) {
        qDiv.addEventListener('click', function () {
          aDiv.classList.toggle('ocs-revealed');
        });
      })(questionDiv, answerDiv);

      card.appendChild(questionDiv);
      card.appendChild(answerDiv);
      el.quizContent.appendChild(card);
    }
  }

  // --- Custom Prompt (Chat UI) ---
  async function handleCustomPrompt() {
    var promptText = el.customPromptInput.value.trim();
    if (!promptText) return;

    // Add user bubble
    var userBubble = document.createElement('div');
    userBubble.className = 'ocs-chat-bubble ocs-user';
    userBubble.textContent = promptText;
    el.chatHistory.appendChild(userBubble);

    // Clear input and disable button
    el.customPromptInput.value = '';
    el.customPromptBtn.disabled = true;

    // Add assistant bubble with loading
    var assistantBubble = document.createElement('div');
    assistantBubble.className = 'ocs-chat-bubble ocs-assistant';
    var responseContent = document.createElement('div');
    responseContent.className = 'ocs-content';
    assistantBubble.appendChild(responseContent);
    el.chatHistory.appendChild(assistantBubble);
    showSectionLoading(responseContent);

    // Scroll to bottom
    el.chatHistory.scrollTop = el.chatHistory.scrollHeight;

    try {
      if (!('LanguageModel' in self)) {
        responseContent.textContent = 'LanguageModel API is not available.';
        return;
      }

      // Extract page data if not already available (e.g., question before summarize)
      var pageData = await ensurePageData();
      if (!pageData) {
        responseContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Not enough text content.';
        return;
      }

      var pageLang = pageData.lang || null;
      var pageLangBase = pageLang ? pageLang.split('-')[0].toLowerCase() : null;
      var outputLang = getOutputLanguage(pageLang);
      var isPageEnglish = (!pageLangBase || pageLangBase === 'en');
      var systemPrompt = 'You are a helpful assistant. Answer based on the provided web page content.';

      // Build context for the prompt.
      // Gemini Nano enters repetition loops with long non-English input,
      // so for non-English pages we use summaries as primary context with
      // only a small page excerpt. For English pages, use full context.
      var contextParts = [];
      var tldrText = el.tldrContent.getAttribute('data-raw-text');
      var kpText = el.keyPointsContent.getAttribute('data-raw-text');

      if (isPageEnglish) {
        // English page: safe to use full page text
        if (tldrText) contextParts.push('Summary: ' + tldrText);
        if (kpText) contextParts.push('Key points: ' + kpText);
        var pageText = pageData.text;
        var maxPageChars = config.TEXT_LIMITS.PROMPT_API_MAX_CHARS;
        if (pageText.length > maxPageChars) {
          pageText = pageText.substring(0, maxPageChars) + config.TEXT_LIMITS.TRUNCATION_SUFFIX;
        }
        contextParts.push('Page content:\n' + pageText);
      } else {
        // Non-English page: use summaries as primary context to avoid Nano repetition loops.
        // When no summaries are available (question before summarize), use more page text.
        if (tldrText) contextParts.push('Summary: ' + tldrText);
        if (kpText) contextParts.push('Key points: ' + kpText);
        var excerptLen = (tldrText || kpText) ? 800 : 2000;
        var excerpt = pageData.text.substring(0, excerptLen);
        if (excerpt) contextParts.push('Page excerpt:\n' + excerpt);
      }

      var fullPrompt = contextParts.join('\n\n') + '\n\nQuestion: ' + promptText;

      // Strategy for non-English output:
      // Always generate in English (Nano's strongest) then translate via Translator API.
      // Gemini Nano produces garbled output with long non-English input regardless of
      // outputLanguage setting, so we keep prompts English-centric and translate after.
      var needsTranslation = outputLang !== 'en' && isTranslatorAvailable();

      if (needsTranslation) {
        // Generate in English, then translate to target language
        var response = await runPromptApi(systemPrompt, fullPrompt, responseContent, 'en');
        if (response) {
          // Gemini Nano may ignore outputLanguage:'en' and generate in the input language
          // (especially when context contains non-English text like Japanese summaries).
          // If the response is already non-English, skip translation to avoid garbling.
          if (!looksLikeEnglish(response)) {
            console.log('[OCS] Prompt API generated non-English despite outputLanguage:en, skipping translation');
            renderMarkdownSafe(response, responseContent);
          } else {
            var englishResponse = response;
            try {
              showTranslatingIndicator(responseContent);
              el.chatHistory.scrollTop = el.chatHistory.scrollHeight;
              var translated = await translateIfNeeded(response, outputLang);
              if (translated && translated !== englishResponse) {
                renderMarkdownSafe(translated, responseContent);
              } else {
                renderMarkdownSafe(englishResponse, responseContent);
              }
            } catch (err) {
              console.warn('Custom prompt translation failed:', err);
              renderMarkdownSafe(englishResponse, responseContent);
            }
          }
        }
      } else if (outputLang !== 'en' && !isTranslatorAvailable()) {
        // Non-English output but no Translator API: generate directly in target language
        // (best effort — may produce lower quality for non-English pages)
        await runPromptApi(systemPrompt, fullPrompt, responseContent, outputLang);
      } else {
        // English output: generate directly
        await runPromptApi(systemPrompt, fullPrompt, responseContent, 'en');
      }
    } catch (err) {
      responseContent.textContent = getErrorMessage(err);
    } finally {
      el.customPromptBtn.disabled = false;
      el.chatHistory.scrollTop = el.chatHistory.scrollHeight;
    }
  }

  // --- Safe Markdown Rendering (DOM-based, no innerHTML) ---
  function renderMarkdownSafe(text, container) {
    container.setAttribute('data-raw-text', text);
    clearElement(container);
    var lines = text.split('\n');
    var currentList = null;

    for (var i = 0; i < lines.length; i++) {
      var trimmed = lines[i].trim();
      if (!trimmed) {
        if (currentList) {
          container.appendChild(currentList);
          currentList = null;
        }
        continue;
      }

      var listMatch = trimmed.match(/^[-*]\s+(.+)/);
      if (listMatch) {
        if (!currentList) {
          currentList = document.createElement('ul');
        }
        var li = document.createElement('li');
        appendInlineFormatted(li, listMatch[1]);
        currentList.appendChild(li);
        continue;
      }

      if (currentList) {
        container.appendChild(currentList);
        currentList = null;
      }

      var p = document.createElement('p');
      appendInlineFormatted(p, trimmed);
      container.appendChild(p);
    }

    if (currentList) {
      container.appendChild(currentList);
    }
  }

  function appendInlineFormatted(parent, text) {
    var regex = /(\*\*(.+?)\*\*|\*(.+?)\*)/g;
    var lastIndex = 0;
    var match;

    while ((match = regex.exec(text)) !== null) {
      if (match.index > lastIndex) {
        parent.appendChild(document.createTextNode(text.substring(lastIndex, match.index)));
      }
      if (match[2]) {
        var strong = document.createElement('strong');
        strong.textContent = match[2];
        parent.appendChild(strong);
      } else if (match[3]) {
        var em = document.createElement('em');
        em.textContent = match[3];
        parent.appendChild(em);
      }
      lastIndex = regex.lastIndex;
    }

    if (lastIndex < text.length) {
      parent.appendChild(document.createTextNode(text.substring(lastIndex)));
    }
  }

  // --- UI Helpers ---
  function clearElement(elem) {
    while (elem.firstChild) {
      elem.removeChild(elem.firstChild);
    }
  }

  function showSections(summaryType) {
    el.resultsArea.classList.remove('ocs-hidden');
    if (summaryType === 'both' || summaryType === 'tldr') {
      el.tldrSection.classList.remove('ocs-hidden');
    } else {
      el.tldrSection.classList.add('ocs-hidden');
    }
    if (summaryType === 'both' || summaryType === 'key-points') {
      el.keyPointsSection.classList.remove('ocs-hidden');
    } else {
      el.keyPointsSection.classList.add('ocs-hidden');
    }
  }

  function showStatus(type, message) {
    el.statusBanner.className = 'ocs-status-banner ocs-' + type;
    el.statusBanner.textContent = message;
    el.statusBanner.classList.remove('ocs-hidden');
  }

  function hideStatus() {
    el.statusBanner.classList.add('ocs-hidden');
  }

  function showSectionLoading(container) {
    clearElement(container);
    var wrapper = document.createElement('div');
    wrapper.className = 'ocs-section-loading';
    var spinner = document.createElement('div');
    spinner.className = 'ocs-section-spinner';
    var label = document.createElement('span');
    label.textContent = chrome.i18n.getMessage('summarizing') || 'Generating...';
    wrapper.appendChild(spinner);
    wrapper.appendChild(label);
    container.appendChild(wrapper);
  }

  function showLoading(show, text) {
    if (show) {
      el.loadingOverlay.classList.remove('ocs-hidden');
      if (text) el.loadingText.textContent = text;
    } else {
      el.loadingOverlay.classList.add('ocs-hidden');
    }
  }

  function clearResults() {
    clearElement(el.tldrContent);
    clearElement(el.keyPointsContent);
    el.resultsArea.classList.add('ocs-hidden');
    hideStatus();
  }

  function showPageInfo(data) {
    if (!data) return;
    el.pageTitle.textContent = data.title || data.url;
    var countText = (data.charCount || 0).toLocaleString() + ' chars';
    if (data.truncated) countText += ' (truncated)';
    el.charCount.textContent = countText;
    el.pageInfo.classList.remove('ocs-hidden');
  }

  function showToast(message) {
    el.toast.textContent = message;
    el.toast.classList.remove('ocs-hidden');
    setTimeout(function () {
      el.toast.classList.add('ocs-hidden');
    }, 2000);
  }

  function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(function () {
      showToast(chrome.i18n.getMessage('copied') || 'Copied!');
    }).catch(function () {
      showToast('Failed to copy');
    });
  }

  // --- Read Aloud (Web Speech API - fully on-device) ---
  var currentReadAloudBtn = null;

  function stopReadAloud() {
    speechSynthesis.cancel();
    if (currentReadAloudBtn) {
      currentReadAloudBtn.classList.remove('ocs-speaking');
      currentReadAloudBtn.querySelector('.ocs-icon-play').classList.remove('ocs-hidden');
      currentReadAloudBtn.querySelector('.ocs-icon-stop').classList.add('ocs-hidden');
      currentReadAloudBtn.title = chrome.i18n.getMessage('readAloud') || 'Read aloud';
      currentReadAloudBtn = null;
    }
  }

  function handleReadAloud(btn) {
    var targetId = btn.getAttribute('data-target');
    var targetEl = document.getElementById(targetId);
    if (!targetEl) return;

    // If this button is already speaking, stop
    if (btn === currentReadAloudBtn) {
      stopReadAloud();
      return;
    }

    // Stop any current speech first
    stopReadAloud();

    // Get text (raw markdown cleaned up, or innerText)
    var rawText = targetEl.getAttribute('data-raw-text');
    var text = rawText || targetEl.innerText;
    if (!text || !text.trim()) return;

    // Clean markdown for speech: remove **, *, [HIGH], [MEDIUM], [LOW] tags
    var speechText = text
      .replace(/\*\*(.+?)\*\*/g, '$1')
      .replace(/\*(.+?)\*/g, '$1')
      .replace(/\[(HIGH|MEDIUM|LOW)\]\s*/gi, '')
      .replace(/^[-*]\s+/gm, '')
      .trim();

    // Select voice based on output language
    var lang = getOutputLanguage(extractedPageData ? extractedPageData.lang : null);
    var langMap = { ja: 'ja-JP', en: 'en-US', es: 'es-ES' };
    var speechLang = langMap[lang] || 'en-US';

    var utterance = new SpeechSynthesisUtterance(speechText);
    utterance.lang = speechLang;
    utterance.rate = SPEECH_SPEEDS[speechSpeedIndex];

    // Use selected voice or find a matching one
    var voices = speechSynthesis.getVoices();
    var selectedURI = el.voiceSelect.value;
    var matchVoice = null;
    if (selectedURI) {
      matchVoice = voices.find(function (v) { return v.voiceURI === selectedURI; });
    }
    if (!matchVoice) {
      matchVoice = voices.find(function (v) {
        return v.lang === speechLang || v.lang.startsWith(lang);
      });
    }
    if (matchVoice) utterance.voice = matchVoice;

    // Update button state
    currentReadAloudBtn = btn;
    btn.classList.add('ocs-speaking');
    btn.querySelector('.ocs-icon-play').classList.add('ocs-hidden');
    btn.querySelector('.ocs-icon-stop').classList.remove('ocs-hidden');
    btn.title = chrome.i18n.getMessage('stopReadAloud') || 'Stop reading';

    utterance.onend = function () { stopReadAloud(); };
    utterance.onerror = function () { stopReadAloud(); };

    speechSynthesis.speak(utterance);
  }

  function getErrorMessage(err) {
    var msg = err.message || String(err);
    var map = {
      notEnoughText: chrome.i18n.getMessage('notEnoughText') ||
        'This page does not have enough text to summarize.',
      inaccessiblePage: chrome.i18n.getMessage('inaccessiblePage') ||
        'Cannot access this page.',
      noActiveTab: chrome.i18n.getMessage('noActiveTab') ||
        'No active tab found.',
      extractionFailed: 'Failed to extract page text.'
    };
    return map[msg] || msg;
  }

  // --- Event Binding ---
  function bindEvents() {
    el.summarizeBtn.addEventListener('click', function () { handleSummarize(true); });
    el.customPromptBtn.addEventListener('click', handleCustomPrompt);
    el.generateQuizBtn.addEventListener('click', handleGenerateQuiz);

    // Cmd+Enter (Mac) / Ctrl+Enter (Win/Linux) to send
    // Avoids conflicts with IME composition (e.isComposing check)
    el.customPromptInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey) && !e.isComposing) {
        e.preventDefault();
        if (!el.customPromptBtn.disabled) {
          handleCustomPrompt();
        }
      }
    });

    // Listen for auto-summarize trigger from background (icon click / keyboard shortcut)
    chrome.runtime.onMessage.addListener(function (message) {
      if (message.type === config.MESSAGES.START_SUMMARIZE) {
        if (currentSettings.autoSummarize && !el.summarizeBtn.disabled) {
          // Re-check tab accessibility before auto-summarizing
          if (currentTabId) {
            updateTabAccessibility(currentTabId).then(function (accessible) {
              if (accessible) handleSummarize();
            });
          }
        }
      }
    });

    // Tab switching: restore cached results or show fresh state
    chrome.tabs.onActivated.addListener(function (activeInfo) {
      switchToTab(activeInfo.tabId);
    });

    // Tab navigation: invalidate cache when URL changes
    chrome.tabs.onUpdated.addListener(function (tabId, changeInfo) {
      if (changeInfo.url) {
        delete tabCache[tabId];
        if (tabId === currentTabId) {
          showFreshState();
          // Check if new URL is accessible before auto-summarizing
          if (!isAccessibleUrl(changeInfo.url)) {
            el.summarizeBtn.disabled = true;
            showStatus('info', chrome.i18n.getMessage('inaccessiblePage') ||
              'Cannot access this page. Summary is not available for browser internal pages.');
            return;
          }
          el.summarizeBtn.disabled = false;
          hideStatus();
          if (currentSettings.autoSummarize) {
            handleSummarize();
          }
        }
      }
    });

    // Tab closed: clean up cache
    chrome.tabs.onRemoved.addListener(function (tabId) {
      delete tabCache[tabId];
    });

    el.settingsBtn.addEventListener('click', function () {
      chrome.runtime.openOptionsPage();
    });

    // Copy page title + URL as Markdown link
    el.copyPageInfoBtn.addEventListener('click', function () {
      if (extractedPageData) {
        var mdLink = '[' + extractedPageData.title + '](' + extractedPageData.url + ')';
        copyToClipboard(mdLink);
        el.copyPageInfoBtn.classList.add('ocs-copied');
        setTimeout(function () {
          el.copyPageInfoBtn.classList.remove('ocs-copied');
        }, 2000);
      }
    });

    el.langSelect.addEventListener('change', function () {
      saveSetting('outputLanguage', el.langSelect.value);
      populateVoiceList(); // refresh voices for new language
    });

    // Voice selection
    el.voiceSelect.addEventListener('change', function () {
      savedVoiceURI = el.voiceSelect.value;
      saveSetting('voiceURI', savedVoiceURI);
    });

    // Voices may load asynchronously
    speechSynthesis.addEventListener('voiceschanged', function () {
      populateVoiceList();
    });

    // Speech speed toggle
    el.speechSpeedBtn.addEventListener('click', cycleSpeechSpeed);

    // Read aloud buttons
    document.querySelectorAll('.ocs-read-aloud-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        handleReadAloud(btn);
      });
    });

    document.querySelectorAll('.ocs-copy-btn').forEach(function (btn) {
      btn.addEventListener('click', function () {
        var targetId = btn.getAttribute('data-target');
        var targetEl = document.getElementById(targetId);
        if (targetEl) {
          // Use stored raw Markdown text if available, fall back to innerText
          var rawText = targetEl.getAttribute('data-raw-text');
          copyToClipboard(rawText || targetEl.innerText);
          btn.classList.add('ocs-copied');
          setTimeout(function () {
            btn.classList.remove('ocs-copied');
          }, 2000);
        }
      });
    });
  }

  // --- Start ---
  document.addEventListener('DOMContentLoaded', init);
})();
