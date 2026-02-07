/**
 * @file Shared configuration for Smart Page Digest extension
 */
(function () {
  'use strict';

  const OCS_CONFIG = Object.freeze({
    SUMMARY_TYPES: Object.freeze({
      TLDR: 'tldr',
      KEY_POINTS: 'key-points',
      BOTH: 'both'
    }),

    SUMMARY_LENGTHS: Object.freeze({
      SHORT: 'short',
      MEDIUM: 'medium',
      LONG: 'long'
    }),

    OUTPUT_LANGUAGES: Object.freeze({
      AUTO: 'auto',
      JA: 'ja',
      EN: 'en',
      ES: 'es'
    }),

    LANGUAGE_LABELS: Object.freeze({
      auto: 'Auto',
      ja: '日本語',
      en: 'English',
      es: 'Español'
    }),

    // English language names for AI prompts (Gemini Nano follows English instructions more reliably)
    LANGUAGE_NAMES_FOR_PROMPT: Object.freeze({
      ja: 'Japanese',
      en: 'English',
      es: 'Spanish'
    }),

    TEXT_LIMITS: Object.freeze({
      MAX_CHARS: 15000,
      PROMPT_API_MAX_CHARS: 4000,
      MIN_CHARS: 50,
      TRUNCATION_SUFFIX: '\n\n[Text truncated for summarization]'
    }),

    DEFAULT_SETTINGS: Object.freeze({
      summaryType: 'both',
      summaryLength: 'medium',
      outputLanguage: 'ja',
      autoSummarize: false,
      customPromptTemplate: ''
    }),

    MESSAGES: Object.freeze({
      EXTRACT_TEXT: 'extractText',
      START_SUMMARIZE: 'startSummarize',
      GET_SETTINGS: 'getSettings',
      SAVE_SETTINGS: 'saveSettings'
    }),

    EXCLUDE_SELECTORS: [
      'nav', 'header', 'footer',
      '[role="navigation"]', '[role="banner"]', '[role="contentinfo"]',
      '.sidebar', '.advertisement', '.ad', '.ads',
      'script', 'style', 'noscript', 'iframe',
      '[aria-hidden="true"]',
      '.cookie-banner', '.popup', '.modal'
    ].join(', '),

    STORAGE_KEY: 'ocsSettings'
  });

  if (typeof globalThis !== 'undefined') {
    globalThis.OCS_CONFIG = OCS_CONFIG;
  }
})();
