/**
 * @file Background service worker
 */
(function () {
  'use strict';

  // Disable default panel-on-click; we handle it manually
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: false })
    .catch(function (err) {
      console.error('[OCS-BG] Failed to set panel behavior:', err);
    });

  // --- Context Menu ---
  function createContextMenu() {
    chrome.contextMenus.removeAll(function () {
      chrome.commands.getAll(function (commands) {
        var cmd = commands.find(function (c) { return c.name === '_execute_action'; });
        var shortcutLabel = cmd && cmd.shortcut ? '  (' + cmd.shortcut + ')' : '';
        var title = chrome.i18n.getMessage('contextMenuSummarize') ||
          'Summarize this page';
        chrome.contextMenus.create({
          id: 'ocs-summarize',
          title: title + shortcutLabel,
          contexts: ['page']
        });
      });
    });
  }

  // Create context menu on install/update
  chrome.runtime.onInstalled.addListener(createContextMenu);
  // Also create on startup (in case SW was terminated)
  chrome.runtime.onStartup.addListener(createContextMenu);

  // Handle context menu click
  chrome.contextMenus.onClicked.addListener(function (info, tab) {
    if (info.menuItemId === 'ocs-summarize' && tab) {
      openPanelAndSummarize(tab);
    }
  });

  // Open side panel + trigger auto-summarize
  function openPanelAndSummarize(tab) {
    chrome.sidePanel.open({ tabId: tab.id }).then(function () {
      // Small delay to ensure panel script is initialized on first open
      setTimeout(function () {
        chrome.runtime.sendMessage({ type: 'startSummarize' }).catch(function () {
          // Panel might not be ready on first open; init() handles that case
        });
      }, 300);
    });
  }

  // Icon click (or keyboard shortcut)
  chrome.action.onClicked.addListener(function (tab) {
    console.log('[OCS-BG] Action clicked, tab:', tab.id);
    openPanelAndSummarize(tab);
  });

  chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
    if (message.type === 'extractText') {
      handleExtractText(sendResponse);
      return true;
    }
  });

  async function handleExtractText(sendResponse) {
    try {
      var tabs = await chrome.tabs.query({
        active: true,
        currentWindow: true
      });
      var tab = tabs[0];
      console.log('[OCS-BG] Active tab:', tab ? { id: tab.id, url: tab.url } : 'none');

      if (!tab || !tab.id) {
        sendResponse({ success: false, error: 'noActiveTab' });
        return;
      }

      if (!isAccessibleUrl(tab.url)) {
        sendResponse({ success: false, error: 'inaccessiblePage', message: 'URL: ' + (tab.url || 'undefined') });
        return;
      }

      // Try sending message to existing content script
      try {
        console.log('[OCS-BG] Sending message to content script...');
        var response = await chrome.tabs.sendMessage(tab.id, {
          type: 'extractText'
        });
        console.log('[OCS-BG] Content script response:', response);
        sendResponse(response);
      } catch (err) {
        console.log('[OCS-BG] Content script not ready, injecting:', err.message);
        // Content script not loaded; inject programmatically
        try {
          await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            files: ['config.js', 'content.js']
          });
          console.log('[OCS-BG] Scripts injected, retrying...');

          // Small delay to ensure scripts are loaded
          await new Promise(function (r) { setTimeout(r, 100); });

          var retryResponse = await chrome.tabs.sendMessage(tab.id, {
            type: 'extractText'
          });
          console.log('[OCS-BG] Retry response:', retryResponse);
          sendResponse(retryResponse);
        } catch (injectErr) {
          console.error('[OCS-BG] Injection failed:', injectErr.message);
          sendResponse({
            success: false,
            error: 'injectionFailed',
            message: injectErr.message
          });
        }
      }
    } catch (err) {
      console.error('[OCS-BG] Background error:', err.message);
      sendResponse({
        success: false,
        error: 'backgroundError',
        message: err.message
      });
    }
  }

  function isAccessibleUrl(url) {
    if (!url) return false;
    return url.startsWith('http://') || url.startsWith('https://');
  }
})();
