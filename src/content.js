/**
 * @file Content script for extracting page text
 */
(function () {
  'use strict';

  function extractPageText() {
    const config = globalThis.OCS_CONFIG;
    const clone = document.body.cloneNode(true);

    // Remove noise elements
    const excludeElements = clone.querySelectorAll(config.EXCLUDE_SELECTORS);
    excludeElements.forEach(function (el) { el.remove(); });

    // Get visible text
    var text = clone.innerText || '';

    // Normalize whitespace
    text = text
      .replace(/\n{3,}/g, '\n\n')
      .replace(/[ \t]+/g, ' ')
      .trim();

    // Truncate if needed
    var truncated = false;
    if (text.length > config.TEXT_LIMITS.MAX_CHARS) {
      text = text.substring(0, config.TEXT_LIMITS.MAX_CHARS) +
        config.TEXT_LIMITS.TRUNCATION_SUFFIX;
      truncated = true;
    }

    return {
      text: text,
      title: document.title,
      url: window.location.href,
      lang: document.documentElement.lang || '',
      charCount: text.length,
      truncated: truncated
    };
  }

  chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
    if (message.type === globalThis.OCS_CONFIG.MESSAGES.EXTRACT_TEXT) {
      try {
        var result = extractPageText();

        if (result.text.length < globalThis.OCS_CONFIG.TEXT_LIMITS.MIN_CHARS) {
          sendResponse({
            success: false,
            error: 'notEnoughText',
            charCount: result.text.length
          });
        } else {
          sendResponse({
            success: true,
            data: result
          });
        }
      } catch (err) {
        sendResponse({
          success: false,
          error: 'extractionFailed',
          message: err.message
        });
      }
    }
    return true;
  });
})();
