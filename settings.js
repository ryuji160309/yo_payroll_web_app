function showToastWithNativeNotice(message, options) {
  if (!message) {
    return null;
  }
  if (typeof window === 'undefined') {
    return null;
  }
  if (typeof window.showToastWithFeedback === 'function') {
    return window.showToastWithFeedback(message, options);
  }
  let toastHandle = null;
  if (typeof window.showToast === 'function') {
    toastHandle = window.showToast(message, options);
  }
  if (typeof window.notifyPlatformFeedback === 'function') {
    window.notifyPlatformFeedback(message, options);
  }
  return toastHandle;
}

document.addEventListener('DOMContentLoaded', async () => {
  initializeHelp('help/settings.txt');
  await ensureSettingsLoaded();
  if (window.settingsError) {
    const err = document.getElementById('settings-error');
    if (err) {
      const lines = ['設定が読み込めませんでした。', 'デフォルトの値を表示しています。'];
      if (Array.isArray(window.settingsErrorDetails)) {
        lines.push(...window.settingsErrorDetails);
      }
      err.textContent = lines.join('\n');
    }
  }
  const select = document.getElementById('store-select');
  Object.keys(DEFAULT_STORES).forEach(key => {
    const opt = document.createElement('option');
    opt.value = key;
    opt.textContent = DEFAULT_STORES[key].name;
    select.appendChild(opt);
  });

  function load(key) {
    const store = DEFAULT_STORES[key] || {};
    document.getElementById('url').value = store.url || '';
    document.getElementById('baseWage').value = store.baseWage;
    document.getElementById('overtime').value = store.overtime;
    document.getElementById('excludeWords').value = (store.excludeWords || []).join(',');
  }

  select.addEventListener('change', () => load(select.value));
  select.value = Object.keys(DEFAULT_STORES)[0];
  load(select.value);

  const message = window.settingsError
    ? '設定を読み込めませんでした。デフォルトの値を表示しています。'
    : '店舗設定を読み込みました。';
  showToastWithNativeNotice(message, { duration: 3200 });

  const urlInput = document.getElementById('url');
  if (urlInput) {
    let isCopying = false;

    const copyUrlToClipboard = async () => {
      if (isCopying) {
        return;
      }
      const value = urlInput.value.trim();
      if (!value) {
        return;
      }

      isCopying = true;
      let copied = false;
      try {
        urlInput.focus();
        urlInput.select();
        urlInput.setSelectionRange(0, value.length);

        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
          try {
            await navigator.clipboard.writeText(value);
            copied = true;
          } catch (error) {
            console.warn('Failed to copy URL via navigator.clipboard', error);
          }
        }

        if (!copied) {
          const hiddenField = document.createElement('textarea');
          hiddenField.value = value;
          hiddenField.setAttribute('readonly', '');
          hiddenField.style.position = 'fixed';
          hiddenField.style.left = '-9999px';
          hiddenField.style.opacity = '0';
          document.body.appendChild(hiddenField);
          hiddenField.select();
          hiddenField.setSelectionRange(0, hiddenField.value.length);
          try {
            copied = document.execCommand('copy');
          } catch (error) {
            console.warn('Failed to copy URL via execCommand', error);
          }
          document.body.removeChild(hiddenField);
        }
      } finally {
        setTimeout(() => {
          try {
            urlInput.setSelectionRange(value.length, value.length);
          } catch (error) {
            // Ignore selection errors on some browsers.
          }
        }, 0);
        isCopying = false;
      }

      if (copied) {
        showToastWithNativeNotice('シートURLをコピーしました。');
      }
    };

    urlInput.addEventListener('click', () => {
      copyUrlToClipboard();
    });

    urlInput.addEventListener('keydown', event => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        copyUrlToClipboard();
      }
    });
  }
});
