const APP_VERSION = '1.9.2';
const SETTINGS_CACHE_KEY = 'remoteSettingsCache';
const VERSION_CHECK_URL = 'version.json';
const UPDATE_DISMISS_KEY = 'updateNoticeDismissedVersion';
const THEME_STORAGE_KEY = 'yoPayrollThemePreference';

(function setupThemePreference() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const STORAGE_KEY = THEME_STORAGE_KEY;
  const metaThemeColor = document.querySelector('meta[name="theme-color"]');
  const THEME_COLORS = {
    light: '#f5f7fa',
    dark: '#0b1218'
  };
  const systemPreference = window.matchMedia('(prefers-color-scheme: dark)');

  let storedPreference = null;
  let currentTheme = 'light';
  let toggleButton = null;

  const readStoredPreference = () => {
    try {
      const stored = window.localStorage.getItem(STORAGE_KEY);
      if (stored === 'light' || stored === 'dark') {
        return stored;
      }
    } catch (error) {
      // Ignore access issues (private mode, etc.)
    }
    return null;
  };

  const persistPreference = theme => {
    try {
      window.localStorage.setItem(STORAGE_KEY, theme);
      storedPreference = theme;
    } catch (error) {
      // Ignore persistence failures
    }
  };

  const updateMetaThemeColor = theme => {
    if (!metaThemeColor) {
      return;
    }
    const normalized = theme === 'dark' ? 'dark' : 'light';
    metaThemeColor.setAttribute('content', THEME_COLORS[normalized] || THEME_COLORS.light);
  };

  const updateToggleState = () => {
    if (!toggleButton) {
      return;
    }
    toggleButton.setAttribute('aria-pressed', currentTheme === 'dark' ? 'true' : 'false');
    toggleButton.title = currentTheme === 'dark'
      ? 'ライトモードに切り替え'
      : 'ダークモードに切り替え';
  };

  const applyTheme = theme => {
    currentTheme = theme === 'dark' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', currentTheme);
    document.documentElement.style.colorScheme = currentTheme;
    updateMetaThemeColor(currentTheme);
    updateToggleState();
  };

  const ensureToggleButton = () => {
    if (toggleButton && toggleButton.isConnected) {
      return toggleButton;
    }
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'theme-toggle';
    button.textContent = '☼';
    button.setAttribute('aria-label', 'テーマ切り替え');
    button.addEventListener('click', () => {
      const nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
      applyTheme(nextTheme);
      persistPreference(nextTheme);
    });

    toggleButton = button;
    const target = document.body || document.documentElement;
    if (target) {
      target.appendChild(button);
    }
    updateToggleState();
    return button;
  };

  const initTheme = () => {
    storedPreference = readStoredPreference();
    const preferredTheme = storedPreference || (systemPreference.matches ? 'dark' : 'light');
    applyTheme(preferredTheme);
    ensureToggleButton();
  };

  const handleSystemChange = event => {
    const persisted = readStoredPreference();
    if (persisted) {
      storedPreference = persisted;
      return;
    }
    storedPreference = null;
    applyTheme(event.matches ? 'dark' : 'light');
  };

  if (typeof systemPreference.addEventListener === 'function') {
    systemPreference.addEventListener('change', handleSystemChange);
  } else if (typeof systemPreference.addListener === 'function') {
    systemPreference.addListener(handleSystemChange);
  }

  window.addEventListener('storage', event => {
    if (event.key !== STORAGE_KEY) {
      return;
    }
    if (event.newValue === 'light' || event.newValue === 'dark') {
      storedPreference = event.newValue;
      applyTheme(event.newValue);
      return;
    }
    storedPreference = null;
    applyTheme(systemPreference.matches ? 'dark' : 'light');
  });

  initTheme();
})();

(function setupTodayButton() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  // 今日の出勤者ページは廃止しました。
})();

(function setupToastSystem() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const DEFAULT_DURATION = 2600;

  function ensureContainer() {
    let container = document.getElementById('toast-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'toast-container';
      container.setAttribute('role', 'status');
      container.setAttribute('aria-live', 'polite');
      container.setAttribute('aria-atomic', 'true');
    }
    if (!container.isConnected) {
      const parent = document.body || document.documentElement;
      if (parent) {
        parent.appendChild(container);
      }
    }
    return container;
  }

  function scheduleRemoval(toast, container) {
    const removeToastElement = () => {
      if (toast.parentElement) {
        toast.parentElement.removeChild(toast);
      }
      if (container && container.parentElement && container.childElementCount === 0) {
        container.parentElement.removeChild(container);
      }
    };

    let removalStarted = false;
    let fallbackRemoval = null;

    const cleanupAfterTransition = () => {
      removeToastElement();
    };

    const onTransitionEnd = event => {
      if (!removalStarted || event.propertyName !== 'opacity') {
        return;
      }
      toast.removeEventListener('transitionend', onTransitionEnd);
      if (fallbackRemoval !== null) {
        clearTimeout(fallbackRemoval);
        fallbackRemoval = null;
      }
      cleanupAfterTransition();
    };

    const startRemoval = () => {
      if (removalStarted) {
        return;
      }
      removalStarted = true;
      toast.addEventListener('transitionend', onTransitionEnd);
      fallbackRemoval = setTimeout(() => {
        toast.removeEventListener('transitionend', onTransitionEnd);
        cleanupAfterTransition();
      }, 700);
    };

    const forceRemove = () => {
      if (!removalStarted) {
        removalStarted = true;
      }
      toast.removeEventListener('transitionend', onTransitionEnd);
      if (fallbackRemoval !== null) {
        clearTimeout(fallbackRemoval);
        fallbackRemoval = null;
      }
      cleanupAfterTransition();
    };

    return { startRemoval, forceRemove };
  }

  window.showToast = function showToast(message, options = {}) {
    if (!message) {
      return null;
    }
    const container = ensureContainer();
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = message;
    container.appendChild(toast);

    const duration = typeof options.duration === 'number' && options.duration > 0
      ? options.duration
      : DEFAULT_DURATION;

    const { startRemoval, forceRemove } = scheduleRemoval(toast, container);

    requestAnimationFrame(() => {
      toast.classList.add('is-visible');
    });

    const hide = () => {
      startRemoval();
      toast.classList.remove('is-visible');
    };

    const hideTimer = setTimeout(hide, duration);

    toast.addEventListener('click', () => {
      clearTimeout(hideTimer);
      hide();
    });

    return {
      element: toast,
      hide: () => {
        clearTimeout(hideTimer);
        hide();
      },
      remove: () => {
        clearTimeout(hideTimer);
        forceRemove();
      }
    };
  };
})();

(function setupPlatformFeedback() {
  if (typeof window === 'undefined' || typeof navigator === 'undefined') {
    return;
  }

  const DEFAULT_SUCCESS_PATTERN = 200;
  const ERROR_PATTERN = [50, 50, 50, 50, 50];

  function resolvePattern(options) {
    if (options && Object.prototype.hasOwnProperty.call(options, 'vibrationPattern')) {
      return options.vibrationPattern;
    }

    const level = options && typeof options.feedbackLevel === 'string'
      ? options.feedbackLevel.toLowerCase()
      : 'success';

    switch (level) {
      case 'error':
        return ERROR_PATTERN;
      case 'success':
      default:
        return DEFAULT_SUCCESS_PATTERN;
    }
  }

  function triggerVibration(pattern) {
    if (typeof navigator.vibrate !== 'function') {
      return;
    }
    try {
      navigator.vibrate(pattern);
    } catch (error) {
      // Ignore vibration failures silently to avoid disrupting the user.
    }
  }

  window.notifyPlatformFeedback = function notifyPlatformFeedback(message, options = {}) {
    const pattern = resolvePattern(options);

    if (typeof navigator.vibrate === 'function') {
      triggerVibration(pattern);
    }
  };

  window.showToastWithFeedback = function showToastWithFeedback(message, options = {}) {
    if (!message) {
      return null;
    }
    const toastHandle = typeof window.showToast === 'function'
      ? window.showToast(message, options)
      : null;

    if (typeof window.notifyPlatformFeedback === 'function') {
      window.notifyPlatformFeedback(message, options);
    }

    return toastHandle;
  };
})();

(function setupButtonPressFeedback() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const PRESSABLE_SELECTOR = 'button, input[type="button"], input[type="submit"], input[type="reset"]';
  const PRESSED_ATTRIBUTE = 'data-pressed';
  const BUTTON_PRESS_VIBRATION = 10;
  const activePresses = new Map();

  function setPressedState(button, isPressed) {
    if (!button) {
      return;
    }
    if (isPressed) {
      button.setAttribute(PRESSED_ATTRIBUTE, 'true');
    } else if (button.hasAttribute(PRESSED_ATTRIBUTE)) {
      button.removeAttribute(PRESSED_ATTRIBUTE);
    }
  }

  function clearPress(pointerId) {
    const record = activePresses.get(pointerId);
    if (!record) {
      return;
    }

    const { button, onPointerLeave, onPointerEnter } = record;
    if (button) {
      button.removeEventListener('pointerleave', onPointerLeave);
      button.removeEventListener('pointerenter', onPointerEnter);
      setPressedState(button, false);
    }
    activePresses.delete(pointerId);
  }

  function clearAllPresses() {
    if (activePresses.size === 0) {
      return;
    }
    Array.from(activePresses.keys()).forEach(clearPress);
  }

  function handlePointerDown(event) {
    if (typeof event.button === 'number' && event.button !== 0) {
      return;
    }

    const button = event.target && event.target.closest
      ? event.target.closest(PRESSABLE_SELECTOR)
      : null;

    if (!button || button.disabled) {
      return;
    }

    const pointerId = event.pointerId;
    if (activePresses.has(pointerId)) {
      clearPress(pointerId);
    }

    const onPointerLeave = leaveEvent => {
      if (leaveEvent.pointerId !== pointerId) {
        return;
      }
      setPressedState(button, false);
    };

    const onPointerEnter = enterEvent => {
      if (enterEvent.pointerId !== pointerId || (enterEvent.buttons & 1) === 0) {
        return;
      }
      setPressedState(button, true);
    };

    setPressedState(button, true);
    button.addEventListener('pointerleave', onPointerLeave);
    button.addEventListener('pointerenter', onPointerEnter);

    activePresses.set(pointerId, {
      button,
      onPointerLeave,
      onPointerEnter
    });
  }

  function handlePointerEnd(event) {
    clearPress(event.pointerId);
  }

  function handleVisibilityChange() {
    if (typeof document === 'undefined') {
      return;
    }
    if (document.visibilityState !== 'visible') {
      clearAllPresses();
    }
  }

  function handleButtonActivation(event) {
    const button = event.target && event.target.closest
      ? event.target.closest(PRESSABLE_SELECTOR)
      : null;

    if (!button || button.disabled) {
      return;
    }

    if (typeof window.notifyPlatformFeedback === 'function') {
      window.notifyPlatformFeedback(null, { vibrationPattern: BUTTON_PRESS_VIBRATION });
    }
  }

  document.addEventListener('pointerdown', handlePointerDown, { passive: true });
  document.addEventListener('pointerup', handlePointerEnd, true);
  document.addEventListener('pointercancel', handlePointerEnd, true);
  window.addEventListener('pointerup', handlePointerEnd);
  window.addEventListener('pointercancel', handlePointerEnd);
  window.addEventListener('mouseup', clearAllPresses, true);
  window.addEventListener('touchend', clearAllPresses, true);
  window.addEventListener('touchcancel', clearAllPresses, true);
  window.addEventListener('blur', clearAllPresses);
  document.addEventListener('visibilitychange', handleVisibilityChange);
  document.addEventListener('click', handleButtonActivation, true);
})();

(function setupHeaderScrollToTop() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const INTERACTIVE_SELECTOR = 'a, button, input, select, textarea, label, [role="button"], [role="link"]';

  function scrollToTop() {
    if (typeof window.scrollTo === 'function') {
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }

    if (typeof document.documentElement !== 'undefined') {
      document.documentElement.scrollTop = 0;
    }
    if (typeof document.body !== 'undefined') {
      document.body.scrollTop = 0;
    }
  }

  function handleHeaderClick(event) {
    if (event.defaultPrevented) {
      return;
    }

    const header = event.currentTarget;
    if (!header || !header.contains(event.target)) {
      return;
    }

    if (event.target && event.target.closest(INTERACTIVE_SELECTOR)) {
      return;
    }

    if (window.scrollY <= 0) {
      return;
    }

    scrollToTop();
  }

  function init() {
    const header = document.querySelector('header');
    if (!header || header.dataset.scrollToTopBound === 'true') {
      return;
    }

    header.dataset.scrollToTopBound = 'true';
    header.addEventListener('click', handleHeaderClick);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();

(function setupUpdateChecker() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  async function clearCachesAndReload(button, statusLabel) {
    button.disabled = true;
    button.textContent = '更新中…';
    if (statusLabel) {
      statusLabel.textContent = '更新の準備をしています…';
    }

    let hadIssues = false;

    async function attempt(description, fn) {
      try {
        await fn();
      } catch (error) {
        hadIssues = true;
        console.error(description, error);
      }
    }

    await attempt('Failed to clear caches', async () => {
      if (window.caches && typeof caches.keys === 'function') {
        const keys = await caches.keys();
        await Promise.all(keys.map(key => caches.delete(key)));
      }
    });

    await attempt('Failed to unregister service workers', async () => {
      if ('serviceWorker' in navigator) {
        if (typeof navigator.serviceWorker.getRegistrations === 'function') {
          const registrations = await navigator.serviceWorker.getRegistrations();
          await Promise.all(
            registrations.map(reg =>
              reg
                .unregister()
                .catch(error => {
                  hadIssues = true;
                  console.error('Service worker unregister failed', error);
                })
            )
          );
        } else if (typeof navigator.serviceWorker.getRegistration === 'function') {
          const registration = await navigator.serviceWorker.getRegistration();
          if (registration) {
            try {
              await registration.unregister();
            } catch (error) {
              hadIssues = true;
              console.error('Service worker unregister failed', error);
            }
          }
        }
      }
    });

    await attempt('Failed to clear cached settings', async () => {
      if (typeof localStorage !== 'undefined') {
        localStorage.removeItem(SETTINGS_CACHE_KEY);
      }
    });

    if (statusLabel) {
      statusLabel.textContent = hadIssues
        ? '一部の古いデータを削除できませんでしたが、更新を試みています…'
        : '最新バージョンを読み込み直しています…';
    }

    const url = new URL(window.location.href);
    url.searchParams.set('forceReload', Date.now().toString());

    let reloadAttempted = false;
    const triggerReload = () => {
      if (reloadAttempted) return;
      reloadAttempted = true;
      try {
        window.location.replace(url.toString());
      } catch (error) {
        console.warn('Reload via replace failed, falling back to href', error);
        window.location.href = url.toString();
      }
    };

    if ('serviceWorker' in navigator && navigator.serviceWorker) {
      const onControllerChange = () => {
        if (!navigator.serviceWorker.controller) {
          navigator.serviceWorker.removeEventListener('controllerchange', onControllerChange);
          triggerReload();
        }
      };
      try {
        navigator.serviceWorker.addEventListener('controllerchange', onControllerChange);
      } catch (error) {
        console.warn('Failed to listen for service worker controller changes', error);
      }
      setTimeout(triggerReload, 500);
    } else {
      triggerReload();
    }

    setTimeout(() => {
      if (!reloadAttempted) {
        triggerReload();
      } else {
        try {
          window.location.reload();
        } catch (error) {
          console.warn('Fallback reload failed', error);
        }
      }
    }, 2000);
  }

  function showUpdateNotice(latestVersion) {
    if (document.getElementById('update-overlay')) {
      return;
    }

    const overlay = document.createElement('div');
    overlay.id = 'update-overlay';

    const popup = document.createElement('div');
    popup.id = 'update-popup';

    const heading = document.createElement('h2');
    heading.id = 'update-title';
    heading.textContent = '更新のお知らせ';
    popup.appendChild(heading);

    const message = document.createElement('p');
    message.id = 'update-message';
    message.textContent = '簡易給与計算ソフトの更新があります。\n下のボタンを押して更新してください。';
    popup.appendChild(message);

    const versionInfo = document.createElement('p');
    versionInfo.id = 'update-version';
    versionInfo.textContent = `現在: ver.${APP_VERSION}\n最新: ver.${latestVersion}`;
    popup.appendChild(versionInfo);

    const status = document.createElement('p');
    status.id = 'update-status';
    popup.appendChild(status);

    const button = document.createElement('button');
    button.id = 'update-confirm';
    button.textContent = '更新する';
    button.addEventListener('click', () => clearCachesAndReload(button, status));
    popup.appendChild(button);

    const continueButton = document.createElement('button');
    continueButton.id = 'update-continue';
    continueButton.textContent = 'このまま続ける';
    popup.appendChild(continueButton);

    const dismissOption = document.createElement('label');
    dismissOption.id = 'update-dismiss-option';
    dismissOption.style.display = 'none';

    const dismissCheckbox = document.createElement('input');
    dismissCheckbox.type = 'checkbox';
    dismissCheckbox.id = 'update-dismiss';
    dismissOption.appendChild(dismissCheckbox);

    const dismissText = document.createElement('span');
    dismissText.textContent = '再び表示しない';
    dismissOption.appendChild(dismissText);
    popup.appendChild(dismissOption);

    const closeButton = document.createElement('button');
    closeButton.id = 'update-close';
    closeButton.textContent = '閉じる';
    closeButton.style.display = 'none';
    popup.appendChild(closeButton);

    continueButton.addEventListener('click', () => {
      message.textContent = '更新しないと不具合の修正や機能の追加が反映されない場合があります。';
      versionInfo.style.display = 'none';
      status.style.display = 'none';
      button.style.display = 'none';
      continueButton.style.display = 'none';
      dismissOption.style.display = 'flex';
      closeButton.style.display = 'block';
    });

    closeButton.addEventListener('click', () => {
      if (dismissCheckbox.checked) {
        try {
          localStorage.setItem(UPDATE_DISMISS_KEY, latestVersion);
        } catch (error) {
          console.warn('Failed to remember update dismissal', error);
        }
      }
      overlay.remove();
    });

    overlay.appendChild(popup);
    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      overlay.style.display = 'flex';
    });
  }

  async function checkForUpdates() {
    try {
      const response = await fetch(`${VERSION_CHECK_URL}?t=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(`Unexpected status ${response.status}`);
      }
      const data = await response.json();
      const latestVersion = data && typeof data.version === 'string' ? data.version.trim() : '';
      if (!latestVersion || latestVersion === APP_VERSION) {
        return;
      }
      try {
        if (localStorage.getItem(UPDATE_DISMISS_KEY) === latestVersion) {
          return;
        }
      } catch (error) {
        console.warn('Failed to read update dismissal', error);
      }
      showUpdateNotice(latestVersion);
    } catch (error) {
      console.warn('Version check failed', error);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', checkForUpdates, { once: true });
  } else {
    checkForUpdates();
  }
})();

(function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) {
    return;
  }
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('service-worker.js').catch(err => {
      console.error('Service worker registration failed:', err);
    });
  });
})();

(function setupHeaderClock() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const MINUTE = 60 * 1000;
  let statusElement = null;
  let messageActive = false;
  let timerId = null;
  let latestClockText = '';

  const pad = value => String(value).padStart(2, '0');

  const formatClock = date => {
    const year = date.getFullYear();
    const month = pad(date.getMonth() + 1);
    const day = pad(date.getDate());
    const hours = pad(date.getHours());
    const minutes = pad(date.getMinutes());
    return `${year}年${month}月${day}日${hours}時${minutes}分`;
  };

  const clearTimer = () => {
    if (timerId !== null) {
      window.clearTimeout(timerId);
      timerId = null;
    }
  };

  const updateLatestClock = () => {
    latestClockText = formatClock(new Date());
    return latestClockText;
  };

  const renderClock = () => {
    if (!statusElement) {
      return;
    }
    const clockText = updateLatestClock();
    if (messageActive) {
      return;
    }
    statusElement.textContent = clockText;
    statusElement.classList.add('has-clock');
    statusElement.classList.remove('is-visible');
    statusElement.classList.remove('network-status--online', 'network-status--offline');
    statusElement.setAttribute('aria-live', 'polite');
  };

  const scheduleNextTick = () => {
    if (!statusElement) {
      return;
    }
    clearTimer();
    const now = Date.now();
    let delay = MINUTE - (now % MINUTE);
    if (!Number.isFinite(delay) || delay <= 0 || delay > MINUTE) {
      delay = MINUTE;
    }
    timerId = window.setTimeout(() => {
      timerId = null;
      renderClock();
      scheduleNextTick();
    }, delay);
  };

  const setMessageActive = active => {
    messageActive = !!active;
    if (!statusElement) {
      return;
    }
    if (messageActive) {
      statusElement.classList.remove('has-clock');
    } else {
      statusElement.classList.add('has-clock');
      statusElement.classList.remove('is-visible');
      statusElement.classList.remove('network-status--online', 'network-status--offline');
      statusElement.textContent = latestClockText || updateLatestClock();
    }
  };

  const refreshClock = () => {
    if (!statusElement) {
      return;
    }
    renderClock();
    scheduleNextTick();
  };

  const handleVisibilityChange = () => {
    if (document.visibilityState === 'visible') {
      renderClock();
      scheduleNextTick();
    }
  };

  const init = () => {
    if (statusElement) {
      return;
    }
    statusElement = document.getElementById('network-status');
    if (!statusElement) {
      return;
    }
    statusElement.classList.add('has-clock');
    renderClock();
    scheduleNextTick();
    document.addEventListener('visibilitychange', handleVisibilityChange);
  };

  window.refreshHeaderClock = () => {
    if (!statusElement) {
      init();
    }
    renderClock();
    scheduleNextTick();
  };

  window.setHeaderStatusMessageActive = active => {
    if (!statusElement) {
      init();
    }
    setMessageActive(active);
    if (!active) {
      renderClock();
      scheduleNextTick();
    }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();

(function setupHeaderDebugMenu() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const RAPID_TAP_INTERVAL = 350;
  const RAPID_TAP_WINDOW = 1200;
  const TAP_COUNT_TO_OPEN = 5;
  const INTERACTIVE_SELECTOR = 'a, button, input, select, textarea, label, [role="button"], [role="link"]';
  const MAX_STORAGE_DETAIL = 10;

  let header = null;
  let tapCount = 0;
  let lastTapTime = 0;
  let firstTapTime = 0;
  let popover = null;
  let outsideClickHandler = null;
  let keydownHandler = null;
  let lastRenderToken = 0;

  const removeDocumentListeners = () => {
    if (outsideClickHandler) {
      document.removeEventListener('pointerdown', outsideClickHandler, true);
      outsideClickHandler = null;
    }
    if (keydownHandler) {
      document.removeEventListener('keydown', keydownHandler, true);
      keydownHandler = null;
    }
  };

  const hidePopover = () => {
    lastRenderToken += 1;
    removeDocumentListeners();
    if (!popover || popover.hidden) {
      return;
    }
    popover.hidden = true;
    popover.setAttribute('aria-hidden', 'true');
    if (header) {
      header.classList.remove('debug-popover-open');
    }
  };

  const ensurePopover = () => {
    if (popover && document.contains(popover)) {
      return popover;
    }
    if (!header || !document.contains(header)) {
      header = document.querySelector('header');
    }
    if (!header) {
      return null;
    }
    popover = document.createElement('div');
    popover.className = 'debug-popover';
    popover.hidden = true;
    popover.setAttribute('aria-hidden', 'true');
    popover.setAttribute('role', 'dialog');
    popover.setAttribute('aria-label', 'デバッグメニュー');

    const inner = document.createElement('div');
    inner.className = 'debug-popover__inner';

    const arrow = document.createElement('div');
    arrow.className = 'debug-popover__arrow';
    inner.appendChild(arrow);

    const headerRow = document.createElement('div');
    headerRow.className = 'debug-popover__header';

    const title = document.createElement('span');
    title.className = 'debug-popover__title';
    title.textContent = 'デバッグメニュー';

    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'debug-popover__close';
    closeBtn.textContent = '閉じる';

    headerRow.appendChild(title);

    const codexButton = document.createElement('button');
    codexButton.type = 'button';
    codexButton.className = 'debug-popover__action-button';
    codexButton.textContent = 'Codex';
    codexButton.addEventListener('click', () => {
      window.open('https://chatgpt.com/codex', '_blank', 'noreferrer');
    });

    const headerActions = document.createElement('div');
    headerActions.className = 'debug-popover__header-actions';
    headerActions.appendChild(codexButton);
    headerActions.appendChild(closeBtn);

    const content = document.createElement('div');
    content.className = 'debug-popover__content';
    content.textContent = '読み込み中…';

    inner.appendChild(headerRow);
    headerRow.appendChild(headerActions);
    inner.appendChild(content);
    popover.appendChild(inner);
    header.appendChild(popover);

    closeBtn.addEventListener('click', () => {
      hidePopover();
    });

    return popover;
  };

  const renderSections = (container, sections) => {
    if (!container) {
      return;
    }
    container.textContent = '';
    if (!Array.isArray(sections) || sections.length === 0) {
      const emptyMessage = document.createElement('div');
      emptyMessage.className = 'debug-popover__value';
      emptyMessage.textContent = '表示できる情報がありません。';
      container.appendChild(emptyMessage);
      return;
    }
    sections.forEach(section => {
      if (!section || !Array.isArray(section.rows) || section.rows.length === 0) {
        return;
      }
      const sectionEl = document.createElement('section');
      sectionEl.className = 'debug-popover__section';
      if (section.title) {
        const titleEl = document.createElement('h3');
        titleEl.className = 'debug-popover__section-title';
        titleEl.textContent = section.title;
        sectionEl.appendChild(titleEl);
      }
      const rowsEl = document.createElement('div');
      rowsEl.className = 'debug-popover__rows';
      section.rows.forEach(row => {
        if (!row || row.value === undefined || row.value === null) {
          return;
        }
        const labelEl = document.createElement('div');
        labelEl.className = 'debug-popover__label';
        labelEl.textContent = row.label || '';
        const valueEl = document.createElement('div');
        valueEl.className = 'debug-popover__value';
        valueEl.textContent = String(row.value);
        rowsEl.appendChild(labelEl);
        rowsEl.appendChild(valueEl);
      });
      if (rowsEl.childNodes.length > 0) {
        sectionEl.appendChild(rowsEl);
        container.appendChild(sectionEl);
      }
    });
    if (!container.hasChildNodes()) {
      const emptyMessage = document.createElement('div');
      emptyMessage.className = 'debug-popover__value';
      emptyMessage.textContent = '表示できる情報がありません。';
      container.appendChild(emptyMessage);
    }
  };

  const formatBytes = bytes => {
    if (!Number.isFinite(bytes) || bytes < 0) {
      return '不明';
    }
    const units = ['B', 'KB', 'MB', 'GB', 'TB'];
    let value = bytes;
    let unitIndex = 0;
    while (value >= 1024 && unitIndex < units.length - 1) {
      value /= 1024;
      unitIndex += 1;
    }
    const precision = value >= 10 || unitIndex === 0 ? 0 : 1;
    return `${value.toFixed(precision)}${units[unitIndex]}`;
  };

  const formatDecimal = (value, fractionDigits = 2) => {
    if (!Number.isFinite(value)) {
      return null;
    }
    const maximumFractionDigits = value >= 10 ? 0 : fractionDigits;
    return value.toLocaleString('ja-JP', {
      minimumFractionDigits: Math.min(maximumFractionDigits, fractionDigits),
      maximumFractionDigits
    });
  };

  const summarizeStorage = getter => {
    if (typeof getter !== 'function') {
      return undefined;
    }
    let storage;
    try {
      storage = getter();
    } catch (error) {
      return undefined;
    }
    if (!storage) {
      return undefined;
    }
    try {
      const entries = [];
      let totalChars = 0;
      for (let i = 0; i < storage.length; i += 1) {
        const key = storage.key(i);
        if (typeof key !== 'string') {
          continue;
        }
        let value = '';
        try {
          value = storage.getItem(key) || '';
        } catch (error) {
          value = '';
        }
        const size = key.length + value.length;
        totalChars += size;
        entries.push({ key, size });
      }
      entries.sort((a, b) => b.size - a.size);
      const lines = entries.map(entry => `${entry.key}: 約${formatBytes(entry.size * 2)}`);
      return {
        count: entries.length,
        bytes: totalChars * 2,
        lines
      };
    } catch (error) {
      console.warn('Failed to summarize storage', error);
      return undefined;
    }
  };

  const gatherDebugSections = async () => {
    const sections = [];
    const now = new Date();
    const nav = typeof navigator !== 'undefined' ? navigator : {};
    const doc = typeof document !== 'undefined' ? document : {};
    const win = typeof window !== 'undefined' ? window : {};

    const evaluateMediaQuery = query => {
      if (!win.matchMedia || typeof win.matchMedia !== 'function') {
        return '未対応';
      }
      try {
        const media = win.matchMedia(query);
        if (!media) {
          return '未対応';
        }
        return media.matches ? '該当' : '非該当';
      } catch (error) {
        return '取得失敗';
      }
    };

    const nowRows = [];
    if (typeof Intl !== 'undefined' && typeof Intl.DateTimeFormat === 'function') {
      try {
        const formatter = new Intl.DateTimeFormat('ja-JP', { dateStyle: 'full', timeStyle: 'medium' });
        nowRows.push({ label: 'ローカル時刻', value: formatter.format(now) });
      } catch (error) {
        nowRows.push({ label: 'ローカル時刻', value: now.toString() });
      }
    } else {
      nowRows.push({ label: 'ローカル時刻', value: now.toString() });
    }
    nowRows.push({ label: 'ISO時刻', value: now.toISOString() });
    if (typeof Intl !== 'undefined' && typeof Intl.DateTimeFormat === 'function') {
      try {
        const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
        if (tz) {
          nowRows.push({ label: 'タイムゾーン', value: tz });
        }
      } catch (error) {
        nowRows.push({ label: 'タイムゾーン', value: '取得失敗' });
      }
    }
    if (win.location && typeof win.location.href === 'string') {
      nowRows.push({ label: 'ページURL', value: win.location.href });
    }
    if (doc.visibilityState) {
      nowRows.push({ label: '表示状態', value: doc.visibilityState });
    }
    if (typeof win.matchMedia === 'function') {
      try {
        const darkMedia = win.matchMedia('(prefers-color-scheme: dark)');
        if (darkMedia) {
          nowRows.push({ label: 'カラースキーム', value: darkMedia.matches ? 'ダーク' : 'ライト' });
        }
      } catch (error) {
        nowRows.push({ label: 'カラースキーム', value: '取得失敗' });
      }
    }
    sections.push({ title: '現在の状態', rows: nowRows });

    const documentRows = [];
    if (doc.title) {
      documentRows.push({ label: 'ページタイトル', value: doc.title });
    }
    if (doc.referrer) {
      documentRows.push({ label: 'リファラー', value: doc.referrer });
    }
    if (doc.documentElement && doc.documentElement.lang) {
      documentRows.push({ label: '言語属性', value: doc.documentElement.lang });
    }
    if (doc.characterSet) {
      documentRows.push({ label: '文字セット', value: doc.characterSet });
    }
    if (doc.contentType) {
      documentRows.push({ label: 'コンテンツタイプ', value: doc.contentType });
    }
    if (doc.lastModified) {
      documentRows.push({ label: '最終更新', value: doc.lastModified });
    }
    if (doc.readyState) {
      documentRows.push({ label: '読み込み状態', value: doc.readyState });
    }
    if (typeof doc.hidden === 'boolean') {
      documentRows.push({ label: 'hidden', value: doc.hidden ? 'はい' : 'いいえ' });
    }
    if (typeof doc.fullscreenElement !== 'undefined') {
      documentRows.push({ label: 'フルスクリーン要素', value: doc.fullscreenElement ? 'あり' : 'なし' });
    }
    sections.push({ title: 'ドキュメント', rows: documentRows });

    const systemRows = [];
    if (nav.userAgent) {
      systemRows.push({ label: 'ユーザーエージェント', value: nav.userAgent });
    }
    if (nav.userAgentData && Array.isArray(nav.userAgentData.brands)) {
      const brands = nav.userAgentData.brands
        .map(brand => `${brand.brand} ${brand.version}`)
        .join(', ');
      if (brands) {
        systemRows.push({ label: 'UA Brands', value: brands });
      }
    }
    if (nav.platform) {
      systemRows.push({ label: 'プラットフォーム', value: nav.platform });
    }
    if (nav.language || (Array.isArray(nav.languages) && nav.languages.length > 0)) {
      const languages = Array.isArray(nav.languages) && nav.languages.length > 0
        ? `${nav.language || nav.languages[0]} (${nav.languages.join(', ')})`
        : nav.language;
      systemRows.push({ label: '言語', value: languages || '不明' });
    }
    if (typeof nav.hardwareConcurrency === 'number') {
      systemRows.push({ label: '論理CPU数', value: `${nav.hardwareConcurrency}コア` });
    }
    if (typeof nav.deviceMemory === 'number') {
      systemRows.push({ label: '推定メモリ', value: `${nav.deviceMemory}GB` });
    }
    if (typeof nav.maxTouchPoints === 'number') {
      systemRows.push({ label: 'タッチポイント', value: String(nav.maxTouchPoints) });
    }
    systemRows.push({ label: 'CPU使用率', value: 'ブラウザからは取得できません' });
    if (typeof nav.cookieEnabled === 'boolean') {
      systemRows.push({ label: 'Cookie', value: nav.cookieEnabled ? '有効' : '無効' });
    }
    sections.push({ title: '端末情報', rows: systemRows });

    const screenRows = [];
    if (win.screen) {
      const scr = win.screen;
      if (scr.width && scr.height) {
        screenRows.push({ label: 'スクリーン解像度', value: `${scr.width}×${scr.height}` });
      }
      if (scr.availWidth && scr.availHeight) {
        screenRows.push({ label: '利用可能領域', value: `${scr.availWidth}×${scr.availHeight}` });
      }
      if (scr.orientation && scr.orientation.type) {
        screenRows.push({ label: '画面の向き', value: scr.orientation.type });
      }
    }
    if (typeof win.innerWidth === 'number' && typeof win.innerHeight === 'number') {
      screenRows.push({ label: 'ビューポート', value: `${Math.round(win.innerWidth)}×${Math.round(win.innerHeight)}` });
    }
    if (typeof win.devicePixelRatio === 'number') {
      screenRows.push({ label: 'デバイスピクセル比', value: win.devicePixelRatio.toFixed(2) });
    }
    if (typeof win.scrollX === 'number' && typeof win.scrollY === 'number') {
      screenRows.push({ label: 'スクロール位置', value: `${Math.round(win.scrollX)}, ${Math.round(win.scrollY)}` });
    }
    sections.push({ title: '画面情報', rows: screenRows });

    const navigationRows = [];
    if (win.history && typeof win.history.length === 'number') {
      navigationRows.push({ label: '履歴件数', value: `${win.history.length}件` });
    }
    if (win.history && typeof win.history.scrollRestoration === 'string') {
      navigationRows.push({ label: 'スクロール復元', value: win.history.scrollRestoration });
    }
    if (win.location && typeof win.location.hash === 'string') {
      navigationRows.push({ label: 'ハッシュ', value: win.location.hash || '(なし)' });
    }
    if (typeof performance !== 'undefined') {
      if (performance.navigation && typeof performance.navigation.type !== 'undefined') {
        const typeMap = {
          0: '通常ロード',
          1: 'リロード',
          2: '履歴移動',
          255: 'その他'
        };
        navigationRows.push({ label: 'NavigationType', value: typeMap[performance.navigation.type] || String(performance.navigation.type) });
      }
      if (typeof performance.getEntriesByType === 'function') {
        try {
          const navEntries = performance.getEntriesByType('navigation');
          if (navEntries && navEntries.length > 0) {
            const navEntry = navEntries[navEntries.length - 1];
            if (navEntry.type) {
              navigationRows.push({ label: 'Navigation API', value: navEntry.type });
            }
            if (Number.isFinite(navEntry.redirectCount)) {
              navigationRows.push({ label: 'リダイレクト回数', value: `${navEntry.redirectCount}回` });
            }
            if (Number.isFinite(navEntry.domContentLoadedEventEnd)) {
              navigationRows.push({ label: 'DOMContentLoaded', value: `${Math.round(navEntry.domContentLoadedEventEnd)}ms` });
            }
            if (Number.isFinite(navEntry.loadEventEnd)) {
              navigationRows.push({ label: 'Loadイベント', value: `${Math.round(navEntry.loadEventEnd)}ms` });
            }
          }
        } catch (error) {
          navigationRows.push({ label: 'Navigation API', value: '取得失敗' });
        }
      }
    }
    sections.push({ title: 'ナビゲーション', rows: navigationRows });

    const storageRows = [];
    if (nav.storage && typeof nav.storage.estimate === 'function') {
      try {
        const estimate = await nav.storage.estimate();
        if (estimate) {
          if (Number.isFinite(estimate.usage)) {
            storageRows.push({ label: 'ストレージ使用量', value: formatBytes(estimate.usage) });
          }
          if (Number.isFinite(estimate.quota)) {
            storageRows.push({ label: 'ストレージ上限', value: formatBytes(estimate.quota) });
          }
          if (Number.isFinite(estimate.usage) && Number.isFinite(estimate.quota) && estimate.quota > 0) {
            const percent = ((estimate.usage / estimate.quota) * 100).toFixed(1);
            storageRows.push({ label: '使用率', value: `${percent}%` });
          }
        }
      } catch (error) {
        storageRows.push({ label: 'ストレージ推定', value: '取得失敗' });
      }
    }
    const localSummary = summarizeStorage(() => win.localStorage);
    if (localSummary) {
      storageRows.push({ label: 'localStorage合計', value: `${formatBytes(localSummary.bytes)} / ${localSummary.count}件` });
      if (Array.isArray(localSummary.lines) && localSummary.lines.length > 0) {
        const limited = localSummary.lines.slice(0, MAX_STORAGE_DETAIL);
        if (localSummary.lines.length > MAX_STORAGE_DETAIL) {
          limited.push(`…他${localSummary.lines.length - MAX_STORAGE_DETAIL}件`);
        }
        storageRows.push({ label: 'localStorage詳細', value: limited.join('\n') });
      } else {
        storageRows.push({ label: 'localStorage詳細', value: '登録なし' });
      }
    } else {
      storageRows.push({ label: 'localStorage', value: '利用不可' });
    }
    const sessionSummary = summarizeStorage(() => win.sessionStorage);
    if (sessionSummary) {
      storageRows.push({ label: 'sessionStorage合計', value: `${formatBytes(sessionSummary.bytes)} / ${sessionSummary.count}件` });
      if (Array.isArray(sessionSummary.lines) && sessionSummary.lines.length > 0) {
        const limited = sessionSummary.lines.slice(0, MAX_STORAGE_DETAIL);
        if (sessionSummary.lines.length > MAX_STORAGE_DETAIL) {
          limited.push(`…他${sessionSummary.lines.length - MAX_STORAGE_DETAIL}件`);
        }
        storageRows.push({ label: 'sessionStorage詳細', value: limited.join('\n') });
      } else {
        storageRows.push({ label: 'sessionStorage詳細', value: '登録なし' });
      }
    } else {
      storageRows.push({ label: 'sessionStorage', value: '利用不可' });
    }
    sections.push({ title: 'ストレージ', rows: storageRows });

    const cacheRows = [];
    if (win.caches && typeof win.caches.keys === 'function') {
      try {
        const cacheNames = await win.caches.keys();
        cacheRows.push({ label: 'キャッシュ数', value: `${cacheNames.length}件` });
        const limitedNames = cacheNames.slice(0, 5);
        for (const name of limitedNames) {
          try {
            const cache = await win.caches.open(name);
            const requests = await cache.keys();
            cacheRows.push({ label: `• ${name}`, value: `${requests.length}件` });
          } catch (error) {
            cacheRows.push({ label: `• ${name}`, value: '詳細取得失敗' });
          }
        }
        if (cacheNames.length > limitedNames.length) {
          cacheRows.push({ label: '…', value: `他${cacheNames.length - limitedNames.length}件` });
        }
      } catch (error) {
        cacheRows.push({ label: 'キャッシュ', value: '取得失敗' });
      }
    } else {
      cacheRows.push({ label: 'キャッシュ', value: '未対応' });
    }
    if ('serviceWorker' in nav) {
      try {
        const registrations = typeof nav.serviceWorker.getRegistrations === 'function'
          ? await nav.serviceWorker.getRegistrations()
          : [];
        cacheRows.push({ label: 'SW登録数', value: `${registrations.length}件` });
        registrations.slice(0, 3).forEach((reg, index) => {
          const scope = reg && reg.scope ? reg.scope : '(scope不明)';
          cacheRows.push({ label: `• scope#${index + 1}`, value: scope });
        });
        if (registrations.length > 3) {
          cacheRows.push({ label: '…', value: `他${registrations.length - 3}件` });
        }
        const controller = nav.serviceWorker.controller;
        cacheRows.push({ label: '制御中スクリプト', value: controller ? controller.scriptURL : 'なし' });
      } catch (error) {
        cacheRows.push({ label: 'Service Worker', value: '取得失敗' });
      }
    } else {
      cacheRows.push({ label: 'Service Worker', value: '未対応' });
    }
    sections.push({ title: 'キャッシュ・サービスワーカー', rows: cacheRows });

    const networkRows = [];
    if (typeof nav.onLine === 'boolean') {
      networkRows.push({ label: 'オンライン', value: nav.onLine ? 'オンライン' : 'オフライン' });
    }
    const connection = nav.connection || nav.mozConnection || nav.webkitConnection;
    if (connection) {
      if (connection.effectiveType) {
        networkRows.push({ label: '回線タイプ', value: connection.effectiveType });
      }
      const downlink = formatDecimal(connection.downlink);
      if (downlink) {
        networkRows.push({ label: '下り推定', value: `${downlink}Mbps` });
      }
      if (Number.isFinite(connection.rtt)) {
        networkRows.push({ label: '推定RTT', value: `${Math.round(connection.rtt)}ms` });
      }
      if (typeof connection.saveData === 'boolean') {
        networkRows.push({ label: 'データセーバー', value: connection.saveData ? '有効' : '無効' });
      }
    } else {
      networkRows.push({ label: '接続情報', value: '未対応' });
    }
    sections.push({ title: 'ネットワーク', rows: networkRows });

    const performanceRows = [];
    if (typeof performance !== 'undefined') {
      if (typeof performance.now === 'function') {
        performanceRows.push({ label: '経過時間', value: `${Math.round(performance.now())}ms` });
      }
      if (typeof performance.timeOrigin === 'number') {
        try {
          const originDate = new Date(performance.timeOrigin);
          performanceRows.push({ label: 'timeOrigin', value: originDate.toISOString() });
        } catch (error) {
          performanceRows.push({ label: 'timeOrigin', value: String(performance.timeOrigin) });
        }
      }
      const memory = performance.memory;
      if (memory && Number.isFinite(memory.usedJSHeapSize)) {
        const used = formatBytes(memory.usedJSHeapSize);
        const total = Number.isFinite(memory.totalJSHeapSize) ? formatBytes(memory.totalJSHeapSize) : '?';
        const limit = Number.isFinite(memory.jsHeapSizeLimit) ? formatBytes(memory.jsHeapSizeLimit) : '?';
        performanceRows.push({ label: 'JS Heap', value: `${used} / ${total} (上限 ${limit})` });
      }
    }
    sections.push({ title: 'パフォーマンス', rows: performanceRows });

    const accessibilityRows = [];
    if (win.matchMedia && typeof win.matchMedia === 'function') {
      accessibilityRows.push({ label: 'モーション軽減', value: evaluateMediaQuery('(prefers-reduced-motion: reduce)') });
      accessibilityRows.push({ label: 'コントラスト', value: evaluateMediaQuery('(prefers-contrast: more)') });
      accessibilityRows.push({ label: 'データ節約', value: evaluateMediaQuery('(prefers-reduced-data: reduce)') });
      accessibilityRows.push({ label: '透過効果', value: evaluateMediaQuery('(prefers-reduced-transparency: reduce)') });
    } else {
      accessibilityRows.push({ label: 'メディアクエリ', value: '未対応' });
    }
    if (typeof nav.doNotTrack !== 'undefined') {
      accessibilityRows.push({ label: 'Do Not Track', value: nav.doNotTrack === '1' ? '有効' : '無効' });
    }
    if (win.visualViewport) {
      try {
        const vp = win.visualViewport;
        if (Number.isFinite(vp.scale)) {
          accessibilityRows.push({ label: 'ビジュアルビューポート倍率', value: vp.scale.toFixed(2) });
        }
        if (Number.isFinite(vp.width) && Number.isFinite(vp.height)) {
          accessibilityRows.push({ label: 'ビジュアルビューポート', value: `${Math.round(vp.width)}×${Math.round(vp.height)}` });
        }
      } catch (error) {
        accessibilityRows.push({ label: 'ビジュアルビューポート', value: '取得失敗' });
      }
    }
    sections.push({ title: 'アクセシビリティ・表示設定', rows: accessibilityRows });

    const inputRows = [];
    if (win.matchMedia && typeof win.matchMedia === 'function') {
      inputRows.push({ label: '細かいポインタ', value: evaluateMediaQuery('(pointer: fine)') });
      inputRows.push({ label: '粗いポインタ', value: evaluateMediaQuery('(pointer: coarse)') });
      inputRows.push({ label: 'any-hover', value: evaluateMediaQuery('(any-hover: hover)') });
      inputRows.push({ label: 'any-pointer', value: evaluateMediaQuery('(any-pointer: coarse)') });
    }
    if (typeof nav.maxTouchPoints === 'number') {
      inputRows.push({ label: 'マルチタッチポイント', value: `${nav.maxTouchPoints}点` });
    }
    if (win.screen && typeof win.screen.orientation === 'object' && win.screen.orientation) {
      if (typeof win.screen.orientation.angle === 'number') {
        inputRows.push({ label: '画面角度', value: `${win.screen.orientation.angle}°` });
      }
    }
    sections.push({ title: '入力・インタラクション', rows: inputRows });

    const mediaDeviceRows = [];
    if (nav.mediaDevices && typeof nav.mediaDevices.enumerateDevices === 'function') {
      try {
        const devices = await nav.mediaDevices.enumerateDevices();
        mediaDeviceRows.push({ label: 'デバイス数', value: `${devices.length}件` });
        devices.slice(0, 5).forEach((device, index) => {
          const name = device.label ? device.label : `(名称不明#${index + 1})`;
          mediaDeviceRows.push({ label: `• ${device.kind}`, value: name });
        });
        if (devices.length > 5) {
          mediaDeviceRows.push({ label: '…', value: `他${devices.length - 5}件` });
        }
      } catch (error) {
        mediaDeviceRows.push({ label: 'メディアデバイス', value: '取得失敗' });
      }
    } else {
      mediaDeviceRows.push({ label: 'メディアデバイス', value: '未対応' });
    }
    sections.push({ title: 'メディアデバイス', rows: mediaDeviceRows });

    const permissionRows = [];
    const permissionNames = [
      'geolocation',
      'notifications',
      'camera',
      'microphone',
      'clipboard-read',
      'clipboard-write',
      'push'
    ];
    if (nav.permissions && typeof nav.permissions.query === 'function') {
      for (const name of permissionNames) {
        try {
          const status = await nav.permissions.query({ name });
          permissionRows.push({ label: name, value: status && status.state ? status.state : '不明' });
        } catch (error) {
          permissionRows.push({ label: name, value: '取得失敗' });
        }
      }
    } else {
      permissionRows.push({ label: '権限API', value: '未対応' });
    }
    sections.push({ title: '権限状態', rows: permissionRows });

    if (typeof nav.getBattery === 'function') {
      try {
        const battery = await nav.getBattery();
        if (battery) {
          const batteryRows = [];
          batteryRows.push({ label: '充電中', value: battery.charging ? 'はい' : 'いいえ' });
          if (Number.isFinite(battery.level)) {
            batteryRows.push({ label: '残量', value: `${Math.round(battery.level * 100)}%` });
          }
          if (Number.isFinite(battery.chargingTime)) {
            batteryRows.push({ label: '満充電まで', value: battery.chargingTime === Infinity ? '不明' : `${Math.round(battery.chargingTime / 60)}分` });
          }
          if (Number.isFinite(battery.dischargingTime)) {
            batteryRows.push({ label: '推定残り時間', value: battery.dischargingTime === Infinity ? '不明' : `${Math.round(battery.dischargingTime / 60)}分` });
          }
          sections.push({ title: 'バッテリー', rows: batteryRows });
        }
      } catch (error) {
        sections.push({ title: 'バッテリー', rows: [{ label: '情報', value: '取得失敗' }] });
      }
    }

    return sections
      .map(section => ({
        title: section.title,
        rows: (section.rows || []).filter(row => row && row.value !== undefined && row.value !== null)
      }))
      .filter(section => section.rows.length > 0);
  };

  const showPopover = async () => {
    header = header || document.querySelector('header');
    const element = ensurePopover();
    if (!header || !element) {
      return;
    }
    lastRenderToken += 1;
    const renderToken = lastRenderToken;
    element.hidden = false;
    element.setAttribute('aria-hidden', 'false');
    header.classList.add('debug-popover-open');
    const content = element.querySelector('.debug-popover__content');
    if (content) {
      content.textContent = '読み込み中…';
    }
    removeDocumentListeners();
    outsideClickHandler = event => {
      if (!element.contains(event.target)) {
        hidePopover();
      }
    };
    keydownHandler = event => {
      if (event.key === 'Escape') {
        event.preventDefault();
        hidePopover();
      }
    };
    document.addEventListener('pointerdown', outsideClickHandler, true);
    document.addEventListener('keydown', keydownHandler, true);
    const closeBtn = element.querySelector('.debug-popover__close');
    if (closeBtn) {
      try {
        closeBtn.focus({ preventScroll: true });
      } catch (error) {
        closeBtn.focus();
      }
    }
    try {
      const sections = await gatherDebugSections();
      if (renderToken !== lastRenderToken || element.hidden) {
        return;
      }
      renderSections(content, sections);
    } catch (error) {
      console.error('Failed to collect debug information', error);
      if (renderToken !== lastRenderToken || !content) {
        return;
      }
      content.textContent = 'デバッグ情報の取得に失敗しました。';
    }
  };

  const resetTapCounter = () => {
    tapCount = 0;
    lastTapTime = 0;
    firstTapTime = 0;
  };

  const handleTap = event => {
    if (!header || !header.contains(event.target)) {
      return;
    }
    if (event.pointerType === 'mouse' && typeof event.button === 'number' && event.button !== 0) {
      return;
    }
    const interactiveTarget = event.target && event.target.closest
      ? event.target.closest(INTERACTIVE_SELECTOR)
      : null;
    if (interactiveTarget) {
      return;
    }

    const now = Date.now();
    if (!firstTapTime || now - lastTapTime > RAPID_TAP_INTERVAL || now - firstTapTime > RAPID_TAP_WINDOW) {
      resetTapCounter();
      firstTapTime = now;
    }

    tapCount += 1;
    lastTapTime = now;

    if (tapCount >= TAP_COUNT_TO_OPEN && now - firstTapTime <= RAPID_TAP_WINDOW) {
      resetTapCounter();
      showPopover();
    }
  };

  const init = () => {
    header = document.querySelector('header');
    if (!header || header.dataset.debugMenuBound === 'true') {
      return;
    }
    header.dataset.debugMenuBound = 'true';
    header.addEventListener('pointerdown', handleTap, true);
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'hidden') {
        hidePopover();
        resetTapCounter();
      }
    });
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();

(function setupNetworkStatusIndicator() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  const ONLINE_MESSAGE_DURATION = 3000;
  const FADE_OUT_DURATION = 400;
  let statusElement = null;
  let hideTimer = null;
  let clearTimer = null;

  function clearTimers() {
    if (hideTimer) {
      clearTimeout(hideTimer);
      hideTimer = null;
    }
    if (clearTimer) {
      clearTimeout(clearTimer);
      clearTimer = null;
    }
  }

  function showOffline() {
    if (!statusElement) {
      return;
    }
    clearTimers();
    statusElement.textContent = '🔴オフライン';
    statusElement.classList.add('is-visible', 'network-status--offline');
    statusElement.classList.remove('network-status--online');
    if (typeof window.setHeaderStatusMessageActive === 'function') {
      window.setHeaderStatusMessageActive(true);
    }
  }

  function hideOnlineMessage() {
    if (!statusElement) {
      return;
    }
    hideTimer = window.setTimeout(() => {
      if (!statusElement) {
        return;
      }
      statusElement.classList.remove('is-visible');
      clearTimer = window.setTimeout(() => {
        if (!statusElement) {
          return;
        }
        statusElement.textContent = '';
        statusElement.classList.remove('network-status--online');
        if (typeof window.setHeaderStatusMessageActive === 'function') {
          window.setHeaderStatusMessageActive(false);
        }
        if (typeof window.refreshHeaderClock === 'function') {
          window.refreshHeaderClock();
        }
        clearTimer = null;
      }, FADE_OUT_DURATION);
      hideTimer = null;
    }, ONLINE_MESSAGE_DURATION);
  }

  function showOnlineRecovered() {
    if (!statusElement) {
      return;
    }
    clearTimers();
    statusElement.textContent = '🟢オンラインに復帰しました';
    statusElement.classList.add('is-visible', 'network-status--online');
    statusElement.classList.remove('network-status--offline');
    if (typeof window.setHeaderStatusMessageActive === 'function') {
      window.setHeaderStatusMessageActive(true);
    }
    hideOnlineMessage();
  }

  function handleOffline() {
    showOffline();
  }

  function handleOnline() {
    showOnlineRecovered();
  }

  function initNetworkStatus() {
    statusElement = document.getElementById('network-status');
    if (!statusElement) {
      return;
    }

    if (!navigator.onLine) {
      showOffline();
    } else {
      statusElement.textContent = '';
      statusElement.classList.remove('is-visible', 'network-status--offline', 'network-status--online');
      if (typeof window.setHeaderStatusMessageActive === 'function') {
        window.setHeaderStatusMessageActive(false);
      }
      if (typeof window.refreshHeaderClock === 'function') {
        window.refreshHeaderClock();
      }
    }

    window.addEventListener('offline', handleOffline);
    window.addEventListener('online', handleOnline);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initNetworkStatus, { once: true });
  } else {
    initNetworkStatus();
  }
})();

let PASSWORD = '3963';
window.settingsError = false;
const SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTKnnQY1d5BXnOstLwIhJOn7IX8aqHXC98XzreJoFscTUFPJXhef7jO2-0KKvZ7_fPF0uZwpbdcEpcV/pub?output=xlsx';
const SETTINGS_SHEET_NAME = '給与計算_設定';

// Simple password gate to restrict access
function initPasswordGate() {
  if (sessionStorage.getItem('pwAuth')) return;

  const overlay = document.createElement('div');
  overlay.id = 'pw-overlay';

  const container = document.createElement('div');
  container.className = 'pw-container';

  const message = document.createElement('div');
  message.id = 'pw-message';
  message.className = 'pw-message';
  message.textContent = 'パスワードを入力してください。';
  container.appendChild(message);

  const display = document.createElement('div');
  display.id = 'pw-display';
  display.className = 'pw-display';
  const slots = [];
  for (let i = 0; i < 4; i++) {
    const span = document.createElement('span');
    span.className = 'pw-slot';
    display.appendChild(span);
    slots.push(span);
  }
  container.appendChild(display);

  if (!settingsLoaded) {
    const loadingCover = document.createElement('div');
    loadingCover.className = 'pw-loading';
    loadingCover.textContent = 'パスワード問い合わせ中・・・';
    display.appendChild(loadingCover);
    settingsLoadPromise.finally(() => {
      loadingCover.remove();
    });
  }

  const keypad = document.createElement('div');
  keypad.className = 'pw-keypad';
  const keys = ['7','8','9','4','5','6','1','2','3','', '0','del'];
  keys.forEach(k => {
    if (k === '') {
      const placeholder = document.createElement('div');
      placeholder.className = 'pw-empty';
      keypad.appendChild(placeholder);
      return;
    }
    const btn = document.createElement('button');
    btn.dataset.key = k;
    btn.textContent = k === 'del' ? 'del' : k;
    keypad.appendChild(btn);
  });
  container.appendChild(keypad);

  const note = document.createElement('div');
  note.className = 'pw-note';
  note.innerHTML = 'デフォルトのパスワードはATMの売上金入金と同じです。<br>パスワードは設定用スプレッドシートから変更できます。';
  container.appendChild(note);

  overlay.appendChild(container);
  document.body.appendChild(overlay);

  let input = '';
  function updateDisplay() {
    slots.forEach((s, i) => {
      s.textContent = input[i] || '';
    });
  }
  function clearInput(msg) {
    input = '';
    updateDisplay();
    if (msg) {
      message.textContent = msg;
      setTimeout(() => {
        message.textContent = 'パスワードを入力してください。';
      }, 2000);
    }
  }
  function handleDigit(d) {
    if (input.length >= 4) return;
    input += d;
    updateDisplay();
    if (input.length === 4) {
      setTimeout(() => {
        if (input === PASSWORD) {
          sessionStorage.setItem('pwAuth', '1');
          window.removeEventListener('keydown', onKey);
          overlay.remove();
        } else {
          clearInput('パスワードが間違っています。');
        }
      }, 150);
    }
  }
  function delDigit() {
    input = input.slice(0, -1);
    updateDisplay();
  }
  function onKey(e) {
    if (e.key >= '0' && e.key <= '9') {
      handleDigit(e.key);
    } else if (e.key === 'Backspace' || e.key === 'Delete') {
      delDigit();
    }
  }
  window.addEventListener('keydown', onKey);

  keypad.addEventListener('click', e => {
    const btn = e.target.closest('button');
    if (!btn) return;
    const k = btn.dataset.key;
    if (k === 'del') {
      delDigit();
    } else {
      handleDigit(k);
    }
  });

  updateDisplay();
}

// Shared regex for time ranges such as "9:00-17:30"
const TIME_RANGE_REGEX = /^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/;

// Break deduction rules (minHours or more => deduct hours)
const BREAK_DEDUCTIONS = [
  { minHours: 8, deduct: 1 },
  { minHours: 7, deduct: 0.75 },
  { minHours: 6, deduct: 0.5 }
];

// Manage loading timers without mutating DOM elements
const loadingMap = new WeakMap();


let DEFAULT_STORES = {
  night: {
    name: '夜勤',
    url: 'https://docs.google.com/spreadsheets/d/1gCGyxiXXxOOhgHG2tk3BlzMpXuaWQULacySlIhhoWRY/edit?gid=601593061#gid=601593061',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  sagamihara_higashi: {
    name: '相模原東大沼店',
    url: 'https://docs.google.com/spreadsheets/d/1fEMEasqSGU30DuvCx6O6D0nJ5j6m6WrMkGTAaSQuqBY/edit?gid=358413717#gid=358413717',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  kobuchi: {
    name: '古淵駅前店',
    url: 'https://docs.google.com/spreadsheets/d/1hSD3sdIQftusWcNegZnGbCtJmByZhzpAvLJegDoJckQ/edit?gid=946573079#gid=946573079',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  hashimoto: {
    name: '相模原橋本五丁目店',
    url: 'https://docs.google.com/spreadsheets/d/1YYvWZaF9Li_RHDLevvOm2ND8ASJ3864uHRkDAiWBEDc/edit?gid=2000770170#gid=2000770170',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  isehara: {
    name: '伊勢原高森七丁目店',
    url: 'https://docs.google.com/spreadsheets/d/1PfEQRnvHcKS5hJ6gkpJQc0VFjDoJUBhHl7JTTyJheZc/edit?gid=34390331#gid=34390331',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  }
};

async function fetchRemoteSettings() {
  try {
    let res;
    const hasOutput = /[?&]output=/.test(SETTINGS_URL);

    if (hasOutput) {
      const direct = await fetch(SETTINGS_URL, { cache: 'no-store' });
      if (direct.ok) {
        res = direct;
      }
    }

    if (!res) {
      const exportUrl = toXlsxExportUrl(SETTINGS_URL);
      if (!exportUrl) {
        window.settingsError = true;
        return;
      }
      try {
        const converted = await fetch(exportUrl, { cache: 'no-store' });
        if (converted.ok) {
          res = converted;
        } else {
          window.settingsError = true;
          return;
        }
      } catch (e) {
        window.settingsError = true;
        return;
      }
    }


    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const sheet = wb.Sheets[SETTINGS_SHEET_NAME];
    if (!sheet) {
      window.settingsError = true;
      window.settingsErrorDetails = [`シート「${SETTINGS_SHEET_NAME}」が見つかりません`];
      return;
    }
    const rawStatus = sheet['B4']?.v;
    const status = rawStatus != null ? String(rawStatus).trim().toUpperCase() : null;
    if (status !== 'ALL_OK') {
      window.settingsError = true;
      const details = [];
      if (rawStatus != null && status !== 'ALL_OK') details.push(String(rawStatus));
      const cells = [
        { addr: 'B5', label: 'URL設定' },
        { addr: 'B6', label: '基本時給設定' },
        { addr: 'B7', label: '時間外倍率設定' },
        { addr: 'B8', label: 'パスワード' },
      ];
      cells.forEach(c => {
        const val = sheet[c.addr]?.v;
        if (val && String(val) !== 'OK') details.push(`${c.label}：${val}`);
      });
      window.settingsErrorDetails = details;
      return;
    }

    const baseWage = Number(sheet['D11']?.v);
    const overtime = Number(sheet['F11']?.v);
    const passwordCell = sheet['J11'];
    const password = passwordCell ? String(passwordCell.v) : null;
    const excludeWords = [];
    for (let r = 11, count = 0; ; r++) {
      const cell = sheet[`H${r}`];
      if (!cell || cell.v === undefined || cell.v === '') {
        if (count > 0) break;
        continue;
      }
      excludeWords.push(String(cell.v));
      count++;
    }
    const stores = {};
    let idx = 1;
    for (let r = 11; ; r++) {
      const nameCell = sheet[`A${r}`];
      const urlCell = sheet[`B${r}`];
      if ((!nameCell || nameCell.v === undefined || nameCell.v === '') && (!urlCell || urlCell.v === undefined || urlCell.v === '')) {
        if (idx > 1) break;
        continue;
      }
      if (nameCell && nameCell.v && urlCell && urlCell.v) {
        const key = `store${idx}`;
        stores[key] = { name: String(nameCell.v), url: String(urlCell.v), baseWage, overtime, excludeWords };
        idx++;
      }
    }
    if (Object.keys(stores).length) {
      window.settingsError = false;
      window.settingsErrorDetails = undefined;
      if (typeof document !== 'undefined') {
        const err = document.getElementById('settings-error');
        if (err) err.textContent = '';
      }
      DEFAULT_STORES = stores;
      if (password) PASSWORD = password;
    } else {
      window.settingsError = true;
    }
  } catch (e) {
    console.error('fetchRemoteSettings failed', e);
    window.settingsError = true;
  }
}

let settingsLoadPromise;
let settingsLoaded = false;
let cacheApplied = false;

function applySettingsRecord(record) {
  if (!record) return false;
  if (record.stores) DEFAULT_STORES = record.stores;
  if (record.password) PASSWORD = record.password;
  if (record.settingsError) window.settingsError = true;
  if (record.settingsErrorDetails) window.settingsErrorDetails = record.settingsErrorDetails;
  return true;
}

function loadSettingsFromCache() {
  let raw;
  try {
    raw = localStorage.getItem(SETTINGS_CACHE_KEY);
  } catch (e) {
    console.error('loadSettingsFromCache failed', e);
    return false;
  }
  if (!raw) return false;
  try {
    const data = JSON.parse(raw);
    return applySettingsRecord(data);
  } catch (e) {
    console.error('loadSettingsFromCache parse failed', e);
    try {
      localStorage.removeItem(SETTINGS_CACHE_KEY);
    } catch (removeErr) {
      console.error('loadSettingsFromCache cleanup failed', removeErr);
    }
    return false;
  }
}

function saveSettingsToCache() {
  const payload = {
    fetchedAt: Date.now(),
    stores: DEFAULT_STORES,
    password: PASSWORD,
    settingsError: window.settingsError || undefined,
    settingsErrorDetails: window.settingsErrorDetails || undefined,
  };
  try {
    localStorage.setItem(SETTINGS_CACHE_KEY, JSON.stringify(payload));
  } catch (e) {
    console.error('saveSettingsToCache failed', e);
  }
}

function shouldFetchRemoteSettings() {
  if (typeof navigator !== 'undefined' && 'onLine' in navigator) {
    return navigator.onLine !== false;
  }
  return true;
}

function ensureSettingsLoaded() {
  if (!cacheApplied) {
    const loaded = loadSettingsFromCache();
    if (loaded) {
      settingsLoaded = true;
    }
    cacheApplied = true;
  }

  if (!settingsLoadPromise) {
    const needsFetch = shouldFetchRemoteSettings();
    const promise = needsFetch
      ? fetchRemoteSettings().then(() => {
          saveSettingsToCache();
        })
      : Promise.resolve();

    settingsLoadPromise = promise.finally(() => {
      settingsLoadPromise = null;
    });
    if (typeof window !== 'undefined') {
      window.settingsLoadPromise = settingsLoadPromise;
    }
  }
  return settingsLoadPromise;
}

settingsLoadPromise = ensureSettingsLoaded();
settingsLoadPromise.then(() => {
  settingsLoaded = true;
});

if (typeof window !== 'undefined' && window.addEventListener) {
  window.addEventListener('online', () => {
    const promise = ensureSettingsLoaded();
    window.settingsLoadPromise = promise;
    promise.then(() => {
      settingsLoaded = true;
    });
  });
}
window.settingsLoadPromise = settingsLoadPromise;
document.addEventListener('DOMContentLoaded', initPasswordGate);

function loadStores() {
  let stored = {};
  const raw = localStorage.getItem('stores');
  if (raw) {
    try {
      stored = JSON.parse(raw);
    } catch (e) {
      console.error('loadStores parse failed', e);
      localStorage.removeItem('stores');
      alert('保存された店舗データが破損していたため、初期設定に戻しました。');
      stored = {};
    }
  }
  const merged = {};
  Object.keys(DEFAULT_STORES).forEach(key => {
    const base = DEFAULT_STORES[key];
    const custom = stored[key] || {};
    const customBaseWage = Number(custom.baseWage);
    const baseWage = Number.isFinite(customBaseWage) ? customBaseWage : base.baseWage;
    const customOvertime = Number(custom.overtime);
    const overtime = Number.isFinite(customOvertime) ? customOvertime : base.overtime;
    const excludeWords = Array.isArray(custom.excludeWords) ? custom.excludeWords : base.excludeWords;
    merged[key] = { ...base, ...custom, baseWage, overtime, excludeWords };
  });
  return merged;
}

function saveStores(stores) {
  try {
    localStorage.setItem('stores', JSON.stringify(stores));
  } catch (e) {
    console.error('saveStores failed', e);
    throw e;
  }
}

function getStore(key) {
  const stores = loadStores();
  return stores[key];
}

function updateStore(key, values) {
  const stores = loadStores();
  stores[key] = { ...stores[key], ...values };
  saveStores(stores);
}

function startLoading(el, text, options = {}) {
  if (!el) return;
  stopLoading(el);
  const rawText = (text || '').replace(/[・.]+$/, '');
  const displayText = rawText === '読込中' ? '読み込み中' : rawText;
  el.textContent = '';

  const { disableSlowNote = false } = options;

  const container = document.createElement('div');
  container.className = 'loading-container';

  const loader = document.createElement('div');
  loader.className = 'loader';
  container.appendChild(loader);

  const message = document.createElement('div');
  message.className = 'loading-message';
  message.textContent = displayText || '読み込み中';
  container.appendChild(message);

  el.appendChild(container);

  let timeout = null;
  if (!disableSlowNote) {
    timeout = setTimeout(() => {
      message.textContent = '通常よりも読み込みに時間がかかっていますが、正常に動作していますのでそのままお待ち下さい。';
      message.classList.remove('loading-message');
      message.classList.add('loading-note');
    }, 5000);
  }

  loadingMap.set(el, { timeout });
}

function stopLoading(el) {
  if (!el) return;
  const timers = loadingMap.get(el);
  if (timers) {
    clearInterval(timers.interval);
    clearTimeout(timers.timeout);
    loadingMap.delete(el);
  }
  el.textContent = '';
}


function extractFileId(url) {
  const match = url.match(/\/d\/(?:e\/)?([a-zA-Z0-9_-]+)(?:\/|$)/);
  return match ? match[1] : null;
}

function toXlsxExportUrl(url) {
  const fileId = extractFileId(url);
  return fileId ? `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx` : null;
}

function getWorkbookCacheKey(url) {
  return `workbookCache:${url}`;
}

function tryGetSessionStorage() {
  try {
    if (typeof sessionStorage === 'undefined') {
      return null;
    }
    // Safari private mode can throw on setItem, so test availability.
    const testKey = '__workbook_cache_test__';
    sessionStorage.setItem(testKey, '1');
    sessionStorage.removeItem(testKey);
    return sessionStorage;
  } catch (e) {
    return null;
  }
}

const workbookCacheStorage = tryGetSessionStorage();

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return btoa(binary);
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

const OFFLINE_WORKBOOK_STORAGE_KEY = 'offlineWorkbook';
let offlineWorkbookCache = null;

function getOfflineWorkbookEntry() {
  if (!workbookCacheStorage) {
    return null;
  }
  let raw;
  try {
    raw = workbookCacheStorage.getItem(OFFLINE_WORKBOOK_STORAGE_KEY);
  } catch (e) {
    return null;
  }
  if (!raw) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (e) {
    try {
      workbookCacheStorage.removeItem(OFFLINE_WORKBOOK_STORAGE_KEY);
    } catch (removeError) {
      // Ignore cleanup failures.
    }
    return null;
  }
}

function clearOfflineWorkbook() {
  offlineWorkbookCache = null;
  if (!workbookCacheStorage) {
    return;
  }
  try {
    workbookCacheStorage.removeItem(OFFLINE_WORKBOOK_STORAGE_KEY);
  } catch (e) {
    // Ignore storage errors when clearing offline data.
  }
}

function setOfflineWorkbookActive(active) {
  if (!workbookCacheStorage) {
    return;
  }
  const entry = getOfflineWorkbookEntry();
  if (!entry) {
    return;
  }
  const nextEntry = { ...entry, active: !!active };
  try {
    workbookCacheStorage.setItem(OFFLINE_WORKBOOK_STORAGE_KEY, JSON.stringify(nextEntry));
  } catch (e) {
    // Ignore storage errors when updating state.
  }
  if (!active) {
    offlineWorkbookCache = null;
  }
}

function setOfflineWorkbook(buffer, meta = {}) {
  if (!(buffer instanceof ArrayBuffer)) {
    throw new Error('Invalid workbook buffer');
  }
  const clonedBuffer = buffer.slice ? buffer.slice(0) : buffer;
  if (!workbookCacheStorage) {
    throw new Error('Offline workbook storage is unavailable');
  }
  const payload = {
    data: arrayBufferToBase64(clonedBuffer),
    fileName: typeof meta.fileName === 'string' ? meta.fileName : '',
    timestamp: Date.now(),
    active: true,
  };
  try {
    workbookCacheStorage.setItem(OFFLINE_WORKBOOK_STORAGE_KEY, JSON.stringify(payload));
    offlineWorkbookCache = clonedBuffer;
  } catch (e) {
    offlineWorkbookCache = null;
    throw e;
  }
}

function getOfflineWorkbookBuffer() {
  const entry = getOfflineWorkbookEntry();
  if (!entry || entry.active === false || !entry.data) {
    offlineWorkbookCache = null;
    return null;
  }
  if (offlineWorkbookCache) {
    return offlineWorkbookCache;
  }
  try {
    offlineWorkbookCache = base64ToArrayBuffer(entry.data);
    return offlineWorkbookCache;
  } catch (e) {
    clearOfflineWorkbook();
    return null;
  }
}

function getOfflineWorkbookInfo() {
  const entry = getOfflineWorkbookEntry();
  if (!entry) {
    return null;
  }
  return {
    fileName: typeof entry.fileName === 'string' ? entry.fileName : '',
    timestamp: typeof entry.timestamp === 'number' ? entry.timestamp : null,
  };
}

function isOfflineWorkbookActive() {
  const entry = getOfflineWorkbookEntry();
  return !!(entry && entry.data && entry.active !== false);
}

function cacheWorkbookBuffer(url, buffer) {
  if (!workbookCacheStorage) {
    return;
  }
  const key = getWorkbookCacheKey(url);
  try {
    const base64 = arrayBufferToBase64(buffer);
    workbookCacheStorage.setItem(key, base64);
  } catch (e) {
    try {
      workbookCacheStorage.removeItem(key);
    } catch (removeError) {
      // Ignore storage cleanup errors.
    }
  }
}

function getCachedWorkbookBuffer(url) {
  if (!workbookCacheStorage) {
    return null;
  }
  const key = getWorkbookCacheKey(url);
  try {
    const base64 = workbookCacheStorage.getItem(key);
    return base64 ? base64ToArrayBuffer(base64) : null;
  } catch (e) {
    return null;
  }
}

function buildSheetMetadata(wb) {
  const metaSheets = wb.Workbook && wb.Workbook.Sheets;
  return wb.SheetNames.map((name, index) => ({
    name,
    index,
    sheetId: metaSheets && metaSheets[index] ? metaSheets[index].sheetId : undefined,
  }));
}

function buildWorkbookResult(wb, sheetIndex = 0) {
  const targetIndex = (sheetIndex >= 0 && sheetIndex < wb.SheetNames.length) ? sheetIndex : 0;
  const sheetName = wb.SheetNames[targetIndex];
  const metaSheets = wb.Workbook && wb.Workbook.Sheets;
  const sheetId = metaSheets && metaSheets[targetIndex] ? metaSheets[targetIndex].sheetId : undefined;
  const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
  return { sheetName, data, sheetId, sheetIndex: targetIndex };
}

async function fetchWorkbook(url, sheetIndex = 0, options = {}) {
  const allowOffline = options && options.allowOffline !== undefined ? options.allowOffline : true;
  const offlineBuffer = allowOffline ? getOfflineWorkbookBuffer() : null;
  if (offlineBuffer) {
    const wb = XLSX.read(offlineBuffer, { type: 'array' });
    return buildWorkbookResult(wb, sheetIndex);
  }
  const exportUrl = toXlsxExportUrl(url);
  if (!exportUrl) {
    throw new Error('Invalid spreadsheet URL');
  }
  let buffer = getCachedWorkbookBuffer(url);
  if (!buffer) {
    const res = await fetch(exportUrl, { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`HTTP ${res.status}`);
    }

    buffer = await res.arrayBuffer();
    cacheWorkbookBuffer(url, buffer);
  }
  const wb = XLSX.read(buffer, { type: 'array' });
  return buildWorkbookResult(wb, sheetIndex);
}

async function fetchSheetList(url, options = {}) {
  const allowOffline = options && options.allowOffline !== undefined ? options.allowOffline : true;
  const offlineBuffer = allowOffline ? getOfflineWorkbookBuffer() : null;
  if (offlineBuffer) {
    const wb = XLSX.read(offlineBuffer, { type: 'array', bookSheets: true });
    return buildSheetMetadata(wb);
  }
  const exportUrl = toXlsxExportUrl(url);
  if (!exportUrl) {
    throw new Error('Invalid spreadsheet URL');
  }
  const res = await fetch(exportUrl, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  const buffer = await res.arrayBuffer();
  cacheWorkbookBuffer(url, buffer);
  const wb = XLSX.read(buffer, { type: 'array', bookSheets: true });
  return buildSheetMetadata(wb);
}

function calculateSalaryFromBreakdown(breakdown, baseWage, overtime) {
  if (!breakdown) return 0;
  const regular = Number(breakdown.regularHours ?? breakdown.regular ?? 0) || 0;
  const overtimeHours = Number(breakdown.overtimeHours ?? breakdown.overtime ?? 0) || 0;
  const overtimeRate = Number.isFinite(overtime) ? Number(overtime) : 1;
  const base = regular * baseWage;
  const extra = overtimeHours * baseWage * overtimeRate;
  return Math.floor(base + extra);
}

function calculateEmployee(schedule, baseWage, overtime) {
  let total = 0;
  let workdays = 0;
  let regularTotal = 0;
  let overtimeTotal = 0;
  let absentDays = 0;
  schedule.forEach(cell => {
    if (cell === null || cell === undefined) return;
    const text = cell.toString().trim();
    if (!text) return;
    if (text === '欠勤') {
      absentDays += 1;
      return;
    }
    const segments = text.split(',');
    let dayHours = 0;
    let hasValid = false;
    segments.forEach(seg => {
      const m = seg.trim().match(TIME_RANGE_REGEX);
      if (!m) return;
      const sh = parseInt(m[1], 10);
      const sm = m[2] ? parseInt(m[2], 10) : 0;
      const eh = parseInt(m[3], 10);
      const em = m[4] ? parseInt(m[4], 10) : 0;
      if (
        sh < 0 || sh > 24 || eh < 0 || eh > 24 ||
        sm < 0 || sm >= 60 || em < 0 || em >= 60
      ) return;
      hasValid = true;
      const start = sh * 60 + sm;
      const end = eh * 60 + em;
      const diff = end >= start ? end - start : 24 * 60 - start + end;
      dayHours += diff / 60;
    });
    if (!hasValid || dayHours <= 0) return;
    workdays++;
    for (const rule of BREAK_DEDUCTIONS) {
      if (dayHours >= rule.minHours) {
        dayHours -= rule.deduct;
        break;
      }
    }
    total += dayHours;
    const regular = Math.min(dayHours, 8);
    const over = Math.max(dayHours - 8, 0);
    regularTotal += regular;
    overtimeTotal += over;
  });
  const breakdown = { regularHours: regularTotal, overtimeHours: overtimeTotal };
  const salary = calculateSalaryFromBreakdown(breakdown, baseWage, overtime);
  return {
    hours: total,
    days: workdays,
    absentDays,
    salary,
    breakdown,
    regularHours: regularTotal,
    overtimeHours: overtimeTotal
  };
}

function calculatePayroll(data, baseWage, overtime, excludeWords = []) {
  const header = data[2];
  const names = [];
  const schedules = [];
  for (let col = 3; col < header.length; col++) {
    const name = header[col];
    if (name && !excludeWords.some(word => name.includes(word))) {
      names.push(name);
      // rows 4-34 contain daily schedules
      schedules.push(data.slice(3, 34).map(row => row[col]));
    }
  }

  const results = names.map((name, idx) => {
    const r = calculateEmployee(schedules[idx], baseWage, overtime);
    return {
      name,
      baseWage,
      hours: r.hours,
      days: r.days,
      absentDays: r.absentDays,
      salary: r.salary,
      baseSalary: r.salary,
      transport: 0,
      breakdown: r.breakdown,
      regularHours: r.regularHours,
      overtimeHours: r.overtimeHours
    };
  });

  const totalSalary = results.reduce((sum, r) => sum + r.salary, 0);
  return { results, totalSalary, schedules };
}

