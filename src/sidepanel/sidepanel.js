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
    settingsBtn: document.getElementById('settingsBtn'),
    langSelect: document.getElementById('langSelect'),
    statusBanner: document.getElementById('statusBanner'),
    pageInfo: document.getElementById('pageInfo'),
    pageTitle: document.getElementById('pageTitle'),
    charCount: document.getElementById('charCount'),
    resultsArea: document.getElementById('resultsArea'),
    tldrSection: document.getElementById('tldrSection'),
    generateTldrBtn: document.getElementById('generateTldrBtn'),
    tldrContent: document.getElementById('tldrContent'),
    keyPointsSection: document.getElementById('keyPointsSection'),
    generateKeyPointsBtn: document.getElementById('generateKeyPointsBtn'),
    keyPointsContent: document.getElementById('keyPointsContent'),
    customPromptInput: document.getElementById('customPromptInput'),
    customPromptBtn: document.getElementById('customPromptBtn'),
    customPromptDetails: document.getElementById('customPromptDetails'),
    chatHistory: document.getElementById('chatHistory'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    loadingText: document.getElementById('loadingText'),
    toast: document.getElementById('toast'),
    quizSection: document.getElementById('quizSection'),
    generateQuizBtn: document.getElementById('generateQuizBtn'),
    quizContent: document.getElementById('quizContent'),
    dialogueSection: document.getElementById('dialogueSection'),
    generateDialogueBtn: document.getElementById('generateDialogueBtn'),
    dialogueContent: document.getElementById('dialogueContent'),
    copyPageInfoBtn: document.getElementById('copyPageInfoBtn'),
    speechSpeedBtn: document.getElementById('speechSpeedBtn'),
    voiceSelect: document.getElementById('voiceSelect')
  };

  var SPEECH_SPEEDS = [0.5, 0.75, 1, 1.25, 1.5, 1.75, 2, 2.5, 3];
  var speechSpeedIndex = 2; // default 1x
  var savedVoiceURI = ''; // persisted voice selection

  // --- Zoom Sync ---
  function syncZoomWithTab(tabId) {
    if (!tabId) return;
    chrome.tabs.getZoom(tabId).then(function (zoomFactor) {
      document.documentElement.style.zoom = zoomFactor;
    }).catch(function () {});
  }

  // --- Initialization ---
  async function init() {
    await loadSettings();
    applyI18n();
    await checkApiAvailability();
    bindEvents();

    // Track current tab and check accessibility
    var tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tabs[0]) {
      currentTabId = tabs[0].id;
      syncZoomWithTab(tabs[0].id);
      var accessible = await updateTabAccessibility(tabs[0].id);
      if (!accessible) return;

      // Auto-extract page data on panel open (but don't generate summaries)
      try {
        extractedPageData = await extractText();
        if (extractedPageData && extractedPageData.text) {
          showPageInfo(extractedPageData);
        }
      } catch (err) {
        console.warn('[OCS] Auto page data extraction failed:', err.message);
      }
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
        el.generateTldrBtn.disabled = true;
        el.generateKeyPointsBtn.disabled = true;
        showStatus('info', chrome.i18n.getMessage('inaccessiblePage') ||
          'Cannot access this page. Summary is not available for browser internal pages.');
        return false;
      }
    } catch (err) {
      // Tab may not exist anymore
      el.generateTldrBtn.disabled = true;
      el.generateKeyPointsBtn.disabled = true;
      return false;
    }
    el.generateTldrBtn.disabled = false;
    el.generateKeyPointsBtn.disabled = false;
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
    console.log('[Cache] Restore for tab', tabId, 'cached:', !!cached);
    if (!cached) return false;

    console.log('[Cache] Has TL;DR:', !!cached.tldr, 'Has Key Points:', !!cached.keyPoints);
    extractedPageData = cached.pageData;
    showPageInfo(cached.pageData);

    // Show results area and sections (always keep them visible so users can generate content)
    el.resultsArea.classList.remove('ocs-hidden');
    el.tldrSection.classList.remove('ocs-hidden');
    el.keyPointsSection.classList.remove('ocs-hidden');

    // Restore content if available, otherwise clear
    if (cached.tldr) {
      console.log('[Cache] Restoring TL;DR, length:', cached.tldr.length);
      renderMarkdownSafe(cached.tldr, el.tldrContent);
    } else {
      console.log('[Cache] No TL;DR to restore');
      clearElement(el.tldrContent);
    }

    if (cached.keyPoints) {
      console.log('[Cache] Restoring Key Points, length:', cached.keyPoints.length);
      renderKeyPointsWithImportance(cached.keyPoints, el.keyPointsContent);
    } else {
      console.log('[Cache] No Key Points to restore');
      clearElement(el.keyPointsContent);
    }

    // Restore or clear quiz content
    if (cached.quizHtml) {
      el.quizContent.innerHTML = cached.quizHtml;
      rebindQuizCards();
    } else {
      clearElement(el.quizContent);
    }

    // Restore or clear dialogue content
    if (cached.dialogueHtml) {
      el.dialogueContent.innerHTML = cached.dialogueHtml;
    } else {
      clearElement(el.dialogueContent);
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
    clearElement(el.dialogueContent);
  }

  // Save quiz, dialogue, and chat state for a tab before switching away.
  function saveCurrentTabQuizChat(tabId) {
    if (!tabId) return;
    var hasQuiz = el.quizContent.childNodes.length > 0;
    var hasDialogue = el.dialogueContent.childNodes.length > 0;
    var hasChat = el.chatHistory.childNodes.length > 0;
    if (!hasQuiz && !hasDialogue && !hasChat) return;

    if (!tabCache[tabId]) {
      // Create a minimal cache entry for quiz/dialogue/chat-only tabs (no prior summarization)
      tabCache[tabId] = {
        pageData: extractedPageData,
        summaryType: currentSettings.summaryType,
        tldr: el.tldrContent.getAttribute('data-raw-text') || null,
        keyPoints: el.keyPointsContent.getAttribute('data-raw-text') || null
      };
    }
    tabCache[tabId].quizHtml = hasQuiz ? el.quizContent.innerHTML : null;
    tabCache[tabId].dialogueHtml = hasDialogue ? el.dialogueContent.innerHTML : null;
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
    syncZoomWithTab(tabId);

    if (isSummarizing) return; // Don't interrupt active summarization

    // Check if the new tab's URL is accessible before doing anything
    var accessible = await updateTabAccessibility(tabId);
    if (!accessible) {
      showFreshState();
      return;
    }

    if (!restoreFromCache(tabId)) {
      showFreshState();
      // Extract page data for new tab
      try {
        extractedPageData = await extractText();
        if (extractedPageData && extractedPageData.text) {
          showPageInfo(extractedPageData);
        }
      } catch (err) {
        console.warn('[OCS] Auto page data extraction failed:', err.message);
      }
    }
  }

  // --- Summarization ---
  var pendingTranslator = null; // Pre-created translator (within user gesture context)
  var untranslatedResults = null; // Stores English results when translation model download needed
  var isSummarizingTldr = false;
  var isSummarizingKeyPoints = false;
  var isGeneratingQuiz = false;
  var isGeneratingDialogue = false;
  var isAnsweringCustomPrompt = false;
  var tldrAbortController = null;
  var keyPointsAbortController = null;
  var quizAbortController = null;
  var dialogueAbortController = null;
  var customPromptAbortController = null;

  // Generate TL;DR summary only
  async function handleGenerateTldr() {
    // Check if already generating - if so, cancel
    if (isSummarizingTldr) {
      if (tldrAbortController) {
        tldrAbortController.abort();
        tldrAbortController = null;
      }
      clearElement(el.tldrContent);
      el.tldrContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Generation cancelled';
      el.generateTldrBtn.querySelector('span').textContent = chrome.i18n.getMessage('generateTldrBtn') || 'Generate TL;DR';
      el.generateTldrBtn.disabled = false;
      isSummarizingTldr = false;
      return;
    }

    isSummarizingTldr = true;
    tldrAbortController = new AbortController();
    el.generateTldrBtn.querySelector('span').textContent = chrome.i18n.getMessage('cancelBtn') || 'Cancel';
    el.generateTldrBtn.disabled = false; // Keep enabled for cancellation
    clearElement(el.tldrContent);
    showSectionLoading(el.tldrContent);

    // Extract page data if not already available
    var pageData = await ensurePageData();
    if (!pageData) {
      clearElement(el.tldrContent);
      el.tldrContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Not enough text content.';
      el.generateTldrBtn.querySelector('span').textContent = chrome.i18n.getMessage('generateTldrBtn') || 'Generate TL;DR';
      el.generateTldrBtn.disabled = false;
      isSummarizingTldr = false;
      tldrAbortController = null;
      return;
    }

    var outputLang = getOutputLanguage(extractedPageData.lang);
    var sharedContext = 'Page title: ' + extractedPageData.title;

    try {
      var createOptions = {
        type: 'tldr',
        format: 'markdown',
        length: currentSettings.summaryLength,
        expectedInputLanguages: ['en', 'ja', 'es'],
        outputLanguage: outputLang,
        sharedContext: sharedContext,
        signal: tldrAbortController.signal,
        monitor: function (m) {
          m.addEventListener('downloadprogress', function (e) {
            var pct = Math.round((e.loaded / e.total) * 100);
            el.tldrContent.textContent = 'Downloading model: ' + pct + '%';
          });
        }
      };

      var tldrResult = await runSummarizer(createOptions, extractedPageData.text, el.tldrContent);

      // Update cache
      if (currentTabId) {
        if (!tabCache[currentTabId]) {
          tabCache[currentTabId] = {
            pageData: extractedPageData,
            summaryType: currentSettings.summaryType,
            tldr: null,
            keyPoints: null
          };
        }
        tabCache[currentTabId].tldr = tldrResult;
        console.log('[Cache] TL;DR saved for tab', currentTabId, 'length:', tldrResult ? tldrResult.length : 0);
      }
    } catch (err) {
      // Check for abort-related errors (name or message contains 'abort')
      var isAborted = err.name === 'AbortError' ||
                      err.message === 'AbortError' ||
                      (err.message && err.message.toLowerCase().includes('abort'));
      if (isAborted) {
        clearElement(el.tldrContent);
        el.tldrContent.textContent = 'Generation cancelled';
      } else {
        clearElement(el.tldrContent);
        el.tldrContent.textContent = 'Error: ' + (err.message || 'Failed to generate TL;DR');
      }
    } finally {
      el.generateTldrBtn.querySelector('span').textContent = chrome.i18n.getMessage('generateTldrBtn') || 'Generate TL;DR';
      el.generateTldrBtn.disabled = false;
      isSummarizingTldr = false;
      tldrAbortController = null;
    }
  }

  // Generate Key Points only
  async function handleGenerateKeyPoints() {
    // Check if already generating - if so, cancel
    if (isSummarizingKeyPoints) {
      if (keyPointsAbortController) {
        keyPointsAbortController.abort();
        keyPointsAbortController = null;
      }
      clearElement(el.keyPointsContent);
      el.keyPointsContent.textContent = 'Generation cancelled';
      el.generateKeyPointsBtn.querySelector('span').textContent = chrome.i18n.getMessage('generateKeyPointsBtn') || 'Generate Key Points';
      el.generateKeyPointsBtn.disabled = false;
      isSummarizingKeyPoints = false;
      return;
    }

    isSummarizingKeyPoints = true;
    keyPointsAbortController = new AbortController();
    el.generateKeyPointsBtn.querySelector('span').textContent = chrome.i18n.getMessage('cancelBtn') || 'Cancel';
    el.generateKeyPointsBtn.disabled = false; // Keep enabled for cancellation
    clearElement(el.keyPointsContent);
    showSectionLoading(el.keyPointsContent);

    // Extract page data if not already available
    var pageData = await ensurePageData();
    if (!pageData) {
      clearElement(el.keyPointsContent);
      el.keyPointsContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Not enough text content.';
      el.generateKeyPointsBtn.querySelector('span').textContent = chrome.i18n.getMessage('generateKeyPointsBtn') || 'Generate Key Points';
      el.generateKeyPointsBtn.disabled = false;
      isSummarizingKeyPoints = false;
      keyPointsAbortController = null;
      return;
    }

    var outputLang = getOutputLanguage(extractedPageData.lang);
    var sharedContext = 'Page title: ' + extractedPageData.title;

    try {
      // Use Prompt API for importance-tagged key points when possible
      if ('LanguageModel' in self) {
        var promptText = extractedPageData.text.length > config.TEXT_LIMITS.PROMPT_API_MAX_CHARS
          ? extractedPageData.text.substring(0, config.TEXT_LIMITS.PROMPT_API_MAX_CHARS) + config.TEXT_LIMITS.TRUNCATION_SUFFIX
          : extractedPageData.text;
        var kpPrompt = 'Summarize the following text as key points (3-7 bullet points).' +
          '\nEach bullet MUST start with an importance tag: [HIGH], [MEDIUM], or [LOW].' +
          '\nFormat: - [HIGH] Most important point here\n\n' +
          sharedContext +
          '\n\nText:\n' + promptText;
        try {
          var keyPointsResult = await runPromptApi(
            'You are a summarization assistant.', kpPrompt, el.keyPointsContent, 'en', keyPointsAbortController ? keyPointsAbortController.signal : undefined
          );

          // Translate if needed
          if (outputLang !== 'en' && keyPointsResult && isTranslatorAvailable()) {
            if (looksLikeEnglish(keyPointsResult)) {
              showTranslatingIndicator(el.keyPointsContent);
              keyPointsResult = await translateKeyPointsIfNeeded(keyPointsResult, outputLang);
            }
          }

          keyPointsResult = repairImportanceTags(keyPointsResult);
          renderKeyPointsWithImportance(keyPointsResult, el.keyPointsContent);

          // Update cache
          if (currentTabId) {
            if (!tabCache[currentTabId]) {
              tabCache[currentTabId] = {
                pageData: extractedPageData,
                summaryType: currentSettings.summaryType,
                tldr: null,
                keyPoints: null
              };
            }
            tabCache[currentTabId].keyPoints = keyPointsResult;
            console.log('[Cache] Key Points saved for tab', currentTabId, 'length:', keyPointsResult ? keyPointsResult.length : 0);
          }
        } catch (err) {
          if (err.message === 'AbortError') {
            throw err;
          }
          console.warn('Prompt API key points failed, falling back to Summarizer:', err);
          // Fallback to Summarizer API
          var createOptions = {
            type: 'key-points',
            format: 'markdown',
            length: currentSettings.summaryLength,
            expectedInputLanguages: ['en', 'ja', 'es'],
            outputLanguage: outputLang,
            sharedContext: sharedContext,
            signal: keyPointsAbortController ? keyPointsAbortController.signal : undefined
          };
          var kpResult = await runSummarizer(createOptions, extractedPageData.text, el.keyPointsContent);

          // Update cache
          if (currentTabId) {
            if (!tabCache[currentTabId]) {
              tabCache[currentTabId] = {
                pageData: extractedPageData,
                summaryType: currentSettings.summaryType,
                tldr: null,
                keyPoints: null
              };
            }
            tabCache[currentTabId].keyPoints = kpResult;
          }
        }
      } else {
        // No Prompt API: use Summarizer API
        var createOptions = {
          type: 'key-points',
          format: 'markdown',
          length: currentSettings.summaryLength,
          expectedInputLanguages: ['en', 'ja', 'es'],
          outputLanguage: outputLang,
          sharedContext: sharedContext,
          signal: keyPointsAbortController ? keyPointsAbortController.signal : undefined
        };
        var kpResult = await runSummarizer(createOptions, extractedPageData.text, el.keyPointsContent);

        // Update cache
        if (currentTabId) {
          if (!tabCache[currentTabId]) {
            tabCache[currentTabId] = {
              pageData: extractedPageData,
              summaryType: currentSettings.summaryType,
              tldr: null,
              keyPoints: null
            };
          }
          tabCache[currentTabId].keyPoints = kpResult;
        }
      }
    } catch (err) {
      // Check for abort-related errors (name or message contains 'abort')
      var isAborted = err.name === 'AbortError' ||
                      err.message === 'AbortError' ||
                      (err.message && err.message.toLowerCase().includes('abort'));
      if (isAborted) {
        clearElement(el.keyPointsContent);
        el.keyPointsContent.textContent = 'Generation cancelled';
      } else {
        clearElement(el.keyPointsContent);
        el.keyPointsContent.textContent = 'Error: ' + (err.message || 'Failed to generate key points');
      }
    } finally {
      el.generateKeyPointsBtn.querySelector('span').textContent = chrome.i18n.getMessage('generateKeyPointsBtn') || 'Generate Key Points';
      el.generateKeyPointsBtn.disabled = false;
      isSummarizingKeyPoints = false;
      keyPointsAbortController = null;
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
        // Pass signal to summarize() method as well
        var summarizeOptions = options.signal ? { signal: options.signal } : undefined;
        var summary = await summarizer.summarize(attempts[i], summarizeOptions);
        if (summary) {
          renderMarkdownSafe(summary, targetElement);
          return summary;
        } else {
          targetElement.textContent = '(No summary generated)';
          return null;
        }
      } catch (err) {
        // Check for abort-related errors and rethrow
        var isAborted = err.name === 'AbortError' ||
                        err.message === 'AbortError' ||
                        (err.message && err.message.toLowerCase().includes('abort'));
        if (isAborted) {
          throw err; // Rethrow abort errors to outer catch
        }
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

  async function runPromptApi(systemPrompt, prompt, targetElement, outputLanguage, signal) {
    var lang = outputLanguage || 'en';
    var createOptions = {
      systemPrompt: systemPrompt,
      expectedInputs: [{ type: 'text' }],
      expectedOutputs: [{ type: 'text', languages: [lang] }],
      outputLanguage: lang
    };
    if (signal) {
      createOptions.signal = signal;
    }
    var session = await LanguageModel.create(createOptions);
    var fullText = '';
    try {
      var promptOptions = signal ? { signal: signal } : undefined;
      var stream = session.promptStreaming(prompt, promptOptions);
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

    // Check if already generating - if so, cancel
    if (isGeneratingQuiz) {
      if (quizAbortController) {
        quizAbortController.abort();
        quizAbortController = null;
      }
      clearElement(el.quizContent);
      el.quizContent.textContent = 'Generation cancelled';
      el.generateQuizBtn.querySelector('span').textContent = chrome.i18n.getMessage('quizGenerateBtn') || 'Generate Quiz';
      el.generateQuizBtn.disabled = false;
      isGeneratingQuiz = false;
      return;
    }

    isGeneratingQuiz = true;
    quizAbortController = new AbortController();
    el.generateQuizBtn.querySelector('span').textContent = chrome.i18n.getMessage('cancelBtn') || 'Cancel';
    el.generateQuizBtn.disabled = false; // Keep enabled for cancellation
    clearElement(el.quizContent);
    showSectionLoading(el.quizContent);

    // Extract page data if not already available (e.g., quiz before summarize)
    var pageData = await ensurePageData();
    if (!pageData) {
      clearElement(el.quizContent);
      el.quizContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Not enough text content.';
      el.generateQuizBtn.querySelector('span').textContent = chrome.i18n.getMessage('quizGenerateBtn') || 'Generate Quiz';
      el.generateQuizBtn.disabled = false;
      isGeneratingQuiz = false;
      quizAbortController = null;
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
        outputLanguage: resolvedLang,
        signal: quizAbortController.signal
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

        var result = await session.prompt(prompt, { signal: quizAbortController.signal });
        if (result) {
          renderQuiz(result);
        }
      } finally {
        session.destroy();
      }
    } catch (err) {
      // Check for abort-related errors (name or message contains 'abort')
      var isAborted = err.name === 'AbortError' ||
                      (err.message && err.message.toLowerCase().includes('abort'));
      if (isAborted) {
        clearElement(el.quizContent);
        el.quizContent.textContent = 'Generation cancelled';
      } else {
        clearElement(el.quizContent);
        el.quizContent.textContent = 'Quiz generation failed: ' + err.message;
      }
    } finally {
      el.generateQuizBtn.querySelector('span').textContent = chrome.i18n.getMessage('quizGenerateBtn') || 'Generate Quiz';
      el.generateQuizBtn.disabled = false;
      isGeneratingQuiz = false;
      quizAbortController = null;
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

  // --- Dialogue Generation ---
  async function handleGenerateDialogue() {
    if (!('LanguageModel' in self)) return;

    // Check if already generating - if so, cancel
    if (isGeneratingDialogue) {
      if (dialogueAbortController) {
        dialogueAbortController.abort();
        dialogueAbortController = null;
      }
      clearElement(el.dialogueContent);
      el.dialogueContent.textContent = 'Generation cancelled';
      el.generateDialogueBtn.querySelector('span').textContent = chrome.i18n.getMessage('dialogueGenerateBtn') || 'Generate Dialogue';
      el.generateDialogueBtn.disabled = false;
      isGeneratingDialogue = false;
      return;
    }

    isGeneratingDialogue = true;
    dialogueAbortController = new AbortController();
    el.generateDialogueBtn.querySelector('span').textContent = chrome.i18n.getMessage('cancelBtn') || 'Cancel';
    el.generateDialogueBtn.disabled = false; // Keep enabled for cancellation
    clearElement(el.dialogueContent);
    showSectionLoading(el.dialogueContent);

    // Extract page data if not already available
    var pageData = await ensurePageData();
    if (!pageData) {
      clearElement(el.dialogueContent);
      el.dialogueContent.textContent = chrome.i18n.getMessage('notEnoughText') || 'Not enough text content.';
      el.generateDialogueBtn.querySelector('span').textContent = chrome.i18n.getMessage('dialogueGenerateBtn') || 'Generate Dialogue';
      el.generateDialogueBtn.disabled = false;
      isGeneratingDialogue = false;
      dialogueAbortController = null;
      return;
    }

    var pageLang = extractedPageData.lang || null;
    var resolvedLang = getOutputLanguage(pageLang);
    var dialogueLangEN = (el.langSelect.value !== 'auto') ?
      (config.LANGUAGE_NAMES_FOR_PROMPT[resolvedLang] || 'English') : 'the same language as the text';

    try {
      var session = await LanguageModel.create({
        systemPrompt: 'You are a dialogue content generator. You choose appropriate characters based on article content and create engaging conversations with impactful openings and humorous endings.',
        expectedInputs: [{ type: 'text' }],
        expectedOutputs: [{ type: 'text', languages: [resolvedLang] }],
        outputLanguage: resolvedLang,
        signal: dialogueAbortController.signal
      });
      try {
        // Build language instruction
        var langInstruction = '';
        if (resolvedLang === 'ja') {
          langInstruction = '重要: すべての対話を日本語で生成してください。登場人物名も日本語の役割名（例: 開発者、批評家、記者）を使用してください。\n\n';
        } else if (resolvedLang === 'es') {
          langInstruction = 'Importante: Genera todo el diálogo en español. Usa nombres de roles en español (ejemplo: Desarrollador, Crítico, Periodista).\n\n';
        } else {
          langInstruction = 'Important: Generate all dialogue in English. Use role names in English (example: Developer, Critic, Reporter).\n\n';
        }

        var prompt = langInstruction +
          'Read the following page content and create a dialogue between 2 characters.\n\n' +
          '【Character Selection - CRITICAL】\n' +
          '1. FIRST: Extract ACTUAL PARTICIPANTS from the article (people, entities, or agents directly involved in the story)\n' +
          '   - Personal experience/incident → "筆者" (Author) vs the other party mentioned\n' +
          '   - Interview → "インタビュアー" vs "被インタビュアー" (or their actual names/roles)\n' +
          '   - Debate/controversy → The two opposing parties by their actual roles\n' +
          '   - AI incident → "筆者" vs "AI" (or specific AI agent name if mentioned)\n' +
          '   - Company announcement → "企業" vs "ユーザー" (using actual company name if clear)\n' +
          '\n' +
          '2. ONLY if no clear participants exist (pure technical/analytical article):\n' +
          '   - Fall back to observer roles: "開発者" vs "批評家", "記者" vs "専門家", etc.\n' +
          '\n' +
          '3. Use SPECIFIC role names from the article context (NOT generic "Kenji" or "Aiko")\n' +
          '   - Example: If article mentions "Scott" and "AI agent", use "筆者" and "AI"\n' +
          '   - Example: If article is by a CEO, use "CEO" instead of generic "筆者"\n\n' +
          '【Dialogue Requirements】\n' +
          '- 【Opening】First line must be an impactful statement that hooks the reader\n' +
          '- Each line: short (1-2 sentences, max 50 chars)\n' +
          '- Tempo: fast, comic\n' +
          '- 10-15 exchanges\n' +
          '- Cover all key points from the article\n' +
          '- 【Ending】Close with humor or memorable conclusion\n' +
          '- IMPORTANT: Characters should speak FROM THEIR PERSPECTIVE (not as third-party observers)\n\n' +
          '【Output Format】\n' +
          '役割名1: [セリフ]\n' +
          '役割名2: [セリフ]\n' +
          '...\n\n' +
          'Page Title: ' + extractedPageData.title + '\n\n' +
          'Page Content:\n' + extractedPageData.text.substring(0, 5000);

        var result = await session.prompt(prompt, { signal: dialogueAbortController.signal });
        if (result) {
          renderDialogue(result);

          // Update cache
          if (currentTabId && tabCache[currentTabId]) {
            tabCache[currentTabId].dialogue = result;
          }
        }
      } finally {
        session.destroy();
      }
    } catch (err) {
      // Check for abort-related errors (name or message contains 'abort')
      var isAborted = err.name === 'AbortError' ||
                      (err.message && err.message.toLowerCase().includes('abort'));
      if (isAborted) {
        clearElement(el.dialogueContent);
        el.dialogueContent.textContent = 'Generation cancelled';
      } else {
        clearElement(el.dialogueContent);
        el.dialogueContent.textContent = 'Dialogue generation failed: ' + err.message;
      }
    } finally {
      el.generateDialogueBtn.querySelector('span').textContent = chrome.i18n.getMessage('dialogueGenerateBtn') || 'Generate Dialogue';
      el.generateDialogueBtn.disabled = false;
      isGeneratingDialogue = false;
      dialogueAbortController = null;
    }
  }

  function renderDialogue(dialogueText) {
    clearElement(el.dialogueContent);
    var lines = dialogueText.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      // Match "CharacterName: message" or "CharacterName：message"
      var match = line.match(/^([^:：]+)[:：]\s*(.+)/);
      if (!match) continue;

      var speaker = match[1].trim();
      var message = match[2].trim();

      var bubble = document.createElement('div');
      bubble.className = 'ocs-dialogue-bubble';

      var speakerDiv = document.createElement('div');
      speakerDiv.className = 'ocs-dialogue-speaker';
      speakerDiv.textContent = speaker;

      var messageDiv = document.createElement('div');
      messageDiv.className = 'ocs-dialogue-message';
      messageDiv.textContent = message;

      bubble.appendChild(speakerDiv);
      bubble.appendChild(messageDiv);
      el.dialogueContent.appendChild(bubble);
    }

    if (el.dialogueContent.childNodes.length === 0) {
      el.dialogueContent.textContent = 'Could not parse dialogue format.';
    }
  }

  // --- Custom Prompt (Chat UI) ---
  async function handleCustomPrompt() {
    // Check if already generating - if so, cancel
    if (isAnsweringCustomPrompt) {
      if (customPromptAbortController) {
        customPromptAbortController.abort();
        customPromptAbortController = null;
      }
      // Remove last assistant bubble (the one being generated)
      var lastBubble = el.chatHistory.lastElementChild;
      if (lastBubble && lastBubble.classList.contains('ocs-assistant')) {
        el.chatHistory.removeChild(lastBubble);
      }
      el.customPromptBtn.querySelector('span').textContent = chrome.i18n.getMessage('askButton') || 'Ask';
      el.customPromptBtn.disabled = false;
      isAnsweringCustomPrompt = false;
      return;
    }

    var promptText = el.customPromptInput.value.trim();
    if (!promptText) return;

    isAnsweringCustomPrompt = true;
    customPromptAbortController = new AbortController();

    // Add user bubble
    var userBubble = document.createElement('div');
    userBubble.className = 'ocs-chat-bubble ocs-user';
    userBubble.textContent = promptText;
    el.chatHistory.appendChild(userBubble);

    // Clear input and change button to cancel
    el.customPromptInput.value = '';
    el.customPromptBtn.querySelector('span').textContent = chrome.i18n.getMessage('cancelBtn') || 'Cancel';
    el.customPromptBtn.disabled = false; // Keep enabled for cancellation

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
        var response = await runPromptApi(systemPrompt, fullPrompt, responseContent, 'en', customPromptAbortController.signal);
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
        await runPromptApi(systemPrompt, fullPrompt, responseContent, outputLang, customPromptAbortController.signal);
      } else {
        // English output: generate directly
        await runPromptApi(systemPrompt, fullPrompt, responseContent, 'en', customPromptAbortController.signal);
      }
    } catch (err) {
      // Check for abort-related errors (name or message contains 'abort')
      var isAborted = err.name === 'AbortError' ||
                      (err.message && err.message.toLowerCase().includes('abort'));
      if (isAborted) {
        // Remove the assistant bubble on abort
        if (assistantBubble && assistantBubble.parentNode) {
          assistantBubble.parentNode.removeChild(assistantBubble);
        }
      } else {
        responseContent.textContent = getErrorMessage(err);
      }
    } finally {
      el.customPromptBtn.querySelector('span').textContent = chrome.i18n.getMessage('askButton') || 'Ask';
      el.customPromptBtn.disabled = false;
      isAnsweringCustomPrompt = false;
      customPromptAbortController = null;
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
    // Keep resultsArea visible so users can generate content on fresh tabs
    el.resultsArea.classList.remove('ocs-hidden');
    // Show TL;DR and Key Points sections so generate buttons are accessible
    el.tldrSection.classList.remove('ocs-hidden');
    el.keyPointsSection.classList.remove('ocs-hidden');
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
    el.generateTldrBtn.addEventListener('click', handleGenerateTldr);
    el.generateKeyPointsBtn.addEventListener('click', handleGenerateKeyPoints);
    el.customPromptBtn.addEventListener('click', handleCustomPrompt);
    el.generateQuizBtn.addEventListener('click', handleGenerateQuiz);
    el.generateDialogueBtn.addEventListener('click', handleGenerateDialogue);

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

    // Note: Auto-summarize removed - users now click individual generate buttons

    // Tab switching: restore cached results or show fresh state
    chrome.tabs.onActivated.addListener(function (activeInfo) {
      switchToTab(activeInfo.tabId);
    });

    chrome.tabs.onZoomChange.addListener(function (zoomChangeInfo) {
      if (zoomChangeInfo.tabId === currentTabId) {
        document.documentElement.style.zoom = zoomChangeInfo.newZoomFactor;
      }
    });

    // Tab navigation: invalidate cache when URL changes
    chrome.tabs.onUpdated.addListener(function (tabId, changeInfo) {
      if (changeInfo.url) {
        delete tabCache[tabId];
        if (tabId === currentTabId) {
          showFreshState();
          // Check if new URL is accessible
          if (!isAccessibleUrl(changeInfo.url)) {
            el.generateTldrBtn.disabled = true;
            el.generateKeyPointsBtn.disabled = true;
            showStatus('info', chrome.i18n.getMessage('inaccessiblePage') ||
              'Cannot access this page. Summary is not available for browser internal pages.');
            return;
          }
          el.generateTldrBtn.disabled = false;
          el.generateKeyPointsBtn.disabled = false;
          hideStatus();
          // Extract page data for new URL
          extractText().then(function (data) {
            if (data && data.text) {
              extractedPageData = data;
              showPageInfo(data);
            }
          }).catch(function (err) {
            console.warn('[OCS] Auto page data extraction failed:', err.message);
          });
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
