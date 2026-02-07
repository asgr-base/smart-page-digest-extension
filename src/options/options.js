/**
 * @file Options page logic
 */
(function () {
  'use strict';

  var config = globalThis.OCS_CONFIG;

  var fields = {
    summaryType: document.getElementById('summaryType'),
    summaryLength: document.getElementById('summaryLength'),
    outputLanguage: document.getElementById('outputLanguage'),
    autoSummarize: document.getElementById('autoSummarize'),
    customPromptTemplate: document.getElementById('customPromptTemplate')
  };
  var resetBtn = document.getElementById('resetBtn');
  var statusMessage = document.getElementById('statusMessage');

  var shortcutKey = document.getElementById('shortcutKey');
  var changeShortcutBtn = document.getElementById('changeShortcutBtn');
  var shortcutNotSet = document.getElementById('shortcutNotSet');

  async function init() {
    applyI18n();
    await loadSettings();
    loadShortcutInfo();
    bindEvents();
  }

  function loadShortcutInfo() {
    chrome.commands.getAll(function (commands) {
      var cmd = commands.find(function (c) { return c.name === '_execute_action'; });
      if (cmd && cmd.shortcut) {
        shortcutKey.textContent = cmd.shortcut;
        shortcutKey.parentElement.style.display = '';
        shortcutNotSet.style.display = 'none';
      } else {
        shortcutKey.parentElement.style.display = 'none';
        shortcutNotSet.style.display = '';
      }
    });
  }

  function applyI18n() {
    document.querySelectorAll('[data-i18n]').forEach(function (elem) {
      var key = elem.getAttribute('data-i18n');
      var msg = chrome.i18n.getMessage(key);
      if (msg) elem.textContent = msg;
    });
  }

  async function loadSettings() {
    try {
      var data = await chrome.storage.sync.get(config.STORAGE_KEY);
      var settings = Object.assign({}, config.DEFAULT_SETTINGS, data[config.STORAGE_KEY] || {});
      fields.summaryType.value = settings.summaryType;
      fields.summaryLength.value = settings.summaryLength;
      fields.outputLanguage.value = settings.outputLanguage;
      fields.autoSummarize.checked = settings.autoSummarize !== false;
      fields.customPromptTemplate.value = settings.customPromptTemplate || '';
    } catch (err) {
      showStatus('error', 'Failed to load settings');
    }
  }

  async function saveSettings() {
    var settings = {
      summaryType: fields.summaryType.value,
      summaryLength: fields.summaryLength.value,
      outputLanguage: fields.outputLanguage.value,
      autoSummarize: fields.autoSummarize.checked,
      customPromptTemplate: fields.customPromptTemplate.value.trim()
    };
    var storageData = {};
    storageData[config.STORAGE_KEY] = settings;
    try {
      await chrome.storage.sync.set(storageData);
      showStatus('success', chrome.i18n.getMessage('optionsSaved') || 'Settings saved');
    } catch (err) {
      showStatus('error', 'Failed to save settings');
    }
  }

  async function resetSettings() {
    fields.summaryType.value = config.DEFAULT_SETTINGS.summaryType;
    fields.summaryLength.value = config.DEFAULT_SETTINGS.summaryLength;
    fields.outputLanguage.value = config.DEFAULT_SETTINGS.outputLanguage;
    fields.autoSummarize.checked = config.DEFAULT_SETTINGS.autoSummarize !== false;
    fields.customPromptTemplate.value = config.DEFAULT_SETTINGS.customPromptTemplate || '';
    await saveSettings();
  }

  function showStatus(type, message) {
    statusMessage.className = 'ocs-status-message ocs-' + type;
    statusMessage.textContent = message;
    statusMessage.classList.remove('ocs-hidden');
    setTimeout(function () {
      statusMessage.classList.add('ocs-hidden');
    }, 3000);
  }

  function bindEvents() {
    // Auto-save on any change
    Object.keys(fields).forEach(function (key) {
      fields[key].addEventListener('change', saveSettings);
    });
    fields.customPromptTemplate.addEventListener('input', debounce(saveSettings, 500));
    resetBtn.addEventListener('click', resetSettings);
    changeShortcutBtn.addEventListener('click', function () {
      chrome.tabs.create({ url: 'chrome://extensions/shortcuts' });
    });
  }

  function debounce(fn, delay) {
    var timer;
    return function () {
      clearTimeout(timer);
      timer = setTimeout(fn, delay);
    };
  }

  document.addEventListener('DOMContentLoaded', init);
})();
