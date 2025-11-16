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
  function createDeferred() {
    let resolved = false;
    let resolver = null;
    const promise = new Promise(resolve => {
      resolver = value => {
        if (!resolved) {
          resolved = true;
          resolve(value);
        }
      };
    });
    return {
      promise,
      resolve: value => {
        if (resolver) {
          resolver(value);
        }
      },
      isResolved: () => resolved
    };
  }

  const tutorialReady = createDeferred();
  const list = document.getElementById('store-list');
  const status = document.getElementById('store-status');
  const offlineControls = document.getElementById('offline-controls');
  const offlineButton = document.getElementById('offline-load-button');
  const offlineInfo = document.getElementById('offline-workbook-info');
  const todayAttendanceBox = document.getElementById('today-attendance');
  const LAST_STORE_STORAGE_KEY = 'lastSelectedStore';
  const IOS_PWA_PROMPT_DISMISSED_KEY = 'iosPwaPromptDismissed';
  let showMultiStoreOverlayForTutorial = () => {
    const overlay = document.getElementById('multi-store-overlay');
    if (overlay && overlay.style.display !== 'flex') {
      overlay.style.display = 'flex';
      return;
    }
    const button = document.getElementById('multi-store-mode-button');
    if (button) {
      button.click();
    }
  };
  let hideMultiStoreOverlayForTutorial = () => {
    const closeButton = document.getElementById('multi-store-close');
    if (closeButton) {
      closeButton.click();
      if (multiStoreTutorialState && typeof multiStoreTutorialState.clearPreview === 'function') {
        multiStoreTutorialState.clearPreview();
      }
      return;
    }
    const overlay = document.getElementById('multi-store-overlay');
    if (overlay) {
      overlay.style.display = 'none';
    }
    if (multiStoreTutorialState && typeof multiStoreTutorialState.clearPreview === 'function') {
      multiStoreTutorialState.clearPreview();
    }
  };
  let multiStoreTutorialState = null;
  let storeButtonsHighlightTarget = null;

  const DAY_IN_MS = 24 * 60 * 60 * 1000;

  function formatPeriodLabel(startDate, endDate) {
    if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
      return '';
    }
    const startMonth = startDate.getMonth() + 1;
    const endMonth = endDate.getMonth() + 1;
    const startYear = startDate.getFullYear();
    const endYear = endDate.getFullYear();
    const needsEndYear = startYear !== endYear;
    const startText = `${startYear}年${startMonth}月${startDate.getDate()}日`;
    const endText = `${needsEndYear ? `${endYear}年` : ''}${endMonth}月${endDate.getDate()}日`;
    return `${startText}～${endText}`;
  }

  function formatTimeLabel(totalMinutes) {
    const minutes = Math.max(0, Math.floor(totalMinutes) % (24 * 60));
    const h = Math.floor(minutes / 60);
    const m = minutes % 60;
    if (m === 0) {
      return String(h);
    }
    return `${h}:${String(m).padStart(2, '0')}`;
  }

  function buildMemberColumns(header, excludeWords = []) {
    const columns = [];
    if (!Array.isArray(header)) {
      return columns;
    }
    for (let col = 3; col < header.length; col++) {
      const rawName = header[col];
      if (!rawName && rawName !== 0) continue;
      const name = rawName.toString().trim();
      if (!name) continue;
      if (excludeWords.some(word => name.includes(word))) continue;
      columns.push({ col, name });
    }
    return columns;
  }

  function parseShiftSegment(segment) {
    if (!segment) return null;
    const trimmed = segment.trim();
    if (!trimmed || trimmed === '欠勤') return null;
    const m = trimmed.match(TIME_RANGE_REGEX);
    if (!m) return null;
    const sh = parseInt(m[1], 10);
    const sm = m[2] ? parseInt(m[2], 10) : 0;
    const eh = parseInt(m[3], 10);
    const em = m[4] ? parseInt(m[4], 10) : 0;
    if (
      !Number.isFinite(sh) || !Number.isFinite(sm) || !Number.isFinite(eh) || !Number.isFinite(em)
      || sh < 0 || sh > 24 || eh < 0 || eh > 24
      || sm < 0 || sm >= 60 || em < 0 || em >= 60
    ) {
      return null;
    }
    const startMinutes = sh * 60 + sm;
    const endMinutes = eh * 60 + em;
    const crossesMidnight = endMinutes < startMinutes;
    const duration = crossesMidnight
      ? (24 * 60 - startMinutes + endMinutes)
      : (endMinutes - startMinutes);
    if (duration <= 0) {
      return null;
    }
    return {
      startMinutes,
      endMinutes,
      crossesMidnight,
      label: `${formatTimeLabel(startMinutes)}-${formatTimeLabel(endMinutes)}`
    };
  }

  function collectEntriesForRow(data, memberColumns, rowIndex, dayOffset) {
    const entries = [];
    const row = Array.isArray(data) ? data[rowIndex] : null;
    if (!row) return entries;

    memberColumns.forEach(({ col, name }) => {
      const cell = row[col];
      if (cell === null || cell === undefined) {
        return;
      }
      const text = cell.toString().trim();
      if (!text || text === '欠勤') {
        return;
      }
      text.split(',').forEach(segment => {
        const parsed = parseShiftSegment(segment);
        if (!parsed) return;
        if (dayOffset < 0 && !parsed.crossesMidnight) {
          return;
        }
        entries.push({
          ...parsed,
          name,
          fromPreviousDay: dayOffset < 0,
          dayOffset
        });
      });
    });

    return entries;
  }

  function extractTodayAttendance(workbook, store, targetDate) {
    if (!workbook || !workbook.data || !Array.isArray(workbook.data)) {
      return null;
    }
    const data = workbook.data;
    const header = data[2] || [];
    const memberColumns = buildMemberColumns(header, store.excludeWords || []);
    if (memberColumns.length === 0) {
      return null;
    }

    const rawYear = data[1] && data[1][2];
    const rawStartMonth = data[1] && data[1][4];
    const year = Number.parseInt(rawYear, 10);
    const startMonth = Number.parseInt(rawStartMonth, 10);
    if (!Number.isFinite(year) || !Number.isFinite(startMonth)) {
      return null;
    }

    const startDate = new Date(year, startMonth - 1, 16);
    const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 15);
    const targetTime = targetDate.getTime();
    if (targetTime < startDate.getTime() || targetTime > endDate.getTime()) {
      return null;
    }

    const dayOffset = Math.floor((targetDate - startDate) / DAY_IN_MS);
    if (dayOffset < 0 || dayOffset > 30) {
      return null;
    }

    const todayRowIndex = 3 + dayOffset;
    const entries = collectEntriesForRow(data, memberColumns, todayRowIndex, 0);
    if (dayOffset > 0) {
      entries.push(...collectEntriesForRow(data, memberColumns, todayRowIndex - 1, -1));
    }

    entries.sort((a, b) => {
      if (a.dayOffset !== b.dayOffset) return a.dayOffset - b.dayOffset;
      if (a.startMinutes !== b.startMinutes) return a.startMinutes - b.startMinutes;
      return a.name.localeCompare(b.name, 'ja');
    });

    return {
      entries,
      periodLabel: formatPeriodLabel(startDate, endDate),
      sheetName: workbook.sheetName || ''
    };
  }

  async function findTodaySheetForStore(store, targetDate) {
    if (!store || !store.url) {
      return null;
    }
    const sheetList = await fetchSheetList(store.url, { allowOffline: true });
    const candidates = Array.isArray(sheetList) ? sheetList.slice().reverse() : [];
    for (const sheet of candidates) {
      const workbook = await fetchWorkbook(store.url, sheet.index, { allowOffline: true });
      const attendance = extractTodayAttendance(workbook, store, targetDate);
      if (attendance) {
        return {
          ...attendance,
          sheetName: attendance.sheetName || sheet.name || ''
        };
      }
    }
    return null;
  }

  function renderTodayAttendance(results, container, statusEl) {
    if (!container || !statusEl) {
      return;
    }
    container.textContent = '';
    const hasSuccessful = results.some(result => result && result.periodLabel !== undefined);
    if (!hasSuccessful) {
      statusEl.textContent = '出勤情報を取得できませんでした。';
    } else {
      statusEl.textContent = '今日の出勤予定を表示しています。';
    }

    results.forEach(result => {
      if (!result) return;
      const storeBlock = document.createElement('div');
      storeBlock.className = 'today-attendance-store';

      const title = document.createElement('div');
      title.className = 'today-attendance-store-title';
      const labels = [];
      if (result.storeName) labels.push(result.storeName);
      if (result.periodLabel) {
        labels.push(result.periodLabel);
      } else if (result.sheetName) {
        labels.push(result.sheetName);
      }
      title.textContent = labels.join(' ／ ');
      storeBlock.appendChild(title);

      if (result.entries && result.entries.length > 0) {
        const listEl = document.createElement('ul');
        listEl.className = 'today-attendance-list';
        result.entries.forEach(entry => {
          const item = document.createElement('li');
          item.className = 'today-attendance-item';
          const range = document.createElement('span');
          range.className = 'today-attendance-range';
          range.textContent = entry.label + (entry.fromPreviousDay ? '（前日開始）' : '');
          const name = document.createElement('span');
          name.className = 'today-attendance-name';
          name.textContent = entry.name;
          item.appendChild(range);
          item.appendChild(name);
          listEl.appendChild(item);
        });
        storeBlock.appendChild(listEl);
      } else {
        const empty = document.createElement('p');
        empty.className = 'today-attendance-empty';
        empty.textContent = result.message || '本日の出勤予定はありません。';
        storeBlock.appendChild(empty);
      }

      container.appendChild(storeBlock);
    });
  }

  async function loadTodayAttendance(stores) {
    if (!todayAttendanceBox) {
      return;
    }

    todayAttendanceBox.textContent = '';
    const header = document.createElement('div');
    header.className = 'today-attendance-header';
    header.textContent = '●今日の出勤者●';
    todayAttendanceBox.appendChild(header);

    const statusEl = document.createElement('p');
    statusEl.className = 'today-attendance-status';
    statusEl.textContent = '出勤者を読み込み中…';
    todayAttendanceBox.appendChild(statusEl);

    const container = document.createElement('div');
    container.className = 'today-attendance-body';
    todayAttendanceBox.appendChild(container);

    const storeKeys = stores ? Object.keys(stores) : [];
    if (storeKeys.length === 0) {
      statusEl.textContent = '店舗情報がありません。設定を確認してください。';
      return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const results = [];

    for (const key of storeKeys) {
      const store = stores[key];
      const result = { storeName: store.name };
      try {
        const attendance = await findTodaySheetForStore(store, today);
        if (attendance) {
          Object.assign(result, attendance);
        } else {
          result.message = '今日に対応するシートが見つかりませんでした。';
        }
      } catch (error) {
        console.error('Failed to load today attendance', error);
        result.message = '出勤情報を取得できませんでした。';
      }
      results.push(result);
    }

    renderTodayAttendance(results, container, statusEl);
  }

  function isIos() {
    const ua = window.navigator.userAgent.toLowerCase();
    const isIpadOs13 = ua.includes('macintosh') && 'ontouchend' in document;
    return /iphone|ipod|ipad/.test(ua) || isIpadOs13;
  }

  function isStandaloneMode() {
    if (window.matchMedia && window.matchMedia('(display-mode: standalone)').matches) {
      return true;
    }
    if (typeof window.navigator.standalone === 'boolean' && window.navigator.standalone) {
      return true;
    }
    return false;
  }

  function isPwaPromptDismissed() {
    try {
      return localStorage.getItem(IOS_PWA_PROMPT_DISMISSED_KEY) === '1';
    } catch (e) {
      return false;
    }
  }

  function setPwaPromptDismissed() {
    try {
      localStorage.setItem(IOS_PWA_PROMPT_DISMISSED_KEY, '1');
    } catch (e) {
      // Ignore storage access issues.
    }
  }

  function showIosPwaPrompt() {
    if (!document.body) {
      return;
    }

    const overlay = document.createElement('div');
    overlay.id = 'ios-pwa-overlay';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');

    const dialog = document.createElement('div');
    dialog.className = 'ios-pwa-dialog';

    const title = document.createElement('h2');
    title.className = 'ios-pwa-title';
    title.textContent = 'ホーム画面に追加すると便利です';
    dialog.appendChild(title);

    const description = document.createElement('p');
    description.className = 'ios-pwa-description';
    description.textContent = '簡易給与計算をホーム画面にアプリとして追加することができます。';
    dialog.appendChild(description);

    const steps = document.createElement('ol');
    steps.className = 'ios-pwa-steps';

    const stepItems = [
      () => {
        const li = document.createElement('li');
        li.append('Safariの');
        const icon = document.createElement('img');
        icon.src = 'icons/share_icon.png';
        icon.alt = '共有アイコン';
        icon.className = 'ios-pwa-step-icon';
        li.appendChild(icon);
        li.append('をタップします。');
        return li;
      },
      () => {
        const li = document.createElement('li');
        li.append('表示されたメニューの下の');
        const icon = document.createElement('img');
        icon.src = 'icons/3point_icon.png';
        icon.alt = 'その他アイコン';
        icon.className = 'ios-pwa-step-icon';
        li.appendChild(icon);
        const label = document.createElement('strong');
        label.textContent = 'その他';
        li.appendChild(label);
        li.append('をタップします。');
        return li;
      },
      () => {
        const li = document.createElement('li');
        li.append('広がったメニューの中に');
        const icon = document.createElement('img');
        icon.src = 'icons/plus_icon.png';
        icon.alt = 'ホーム画面に追加アイコン';
        icon.className = 'ios-pwa-step-icon';
        li.appendChild(icon);
        const label = document.createElement('strong');
        label.textContent = 'ホーム画面に追加';
        li.appendChild(label);
        li.append('をタップします。');
        return li;
      },
      () => {
        const li = document.createElement('li');
        li.append('Webアプリとして開くを');
        const labelOn = document.createElement('strong');
        labelOn.textContent = 'オン';
        li.appendChild(labelOn);
        li.append('にしたまま');
        const labelAdd = document.createElement('strong');
        labelAdd.textContent = '追加';
        li.appendChild(labelAdd);
        li.append('をタップします。');
        return li;
      }
    ];

    stepItems.forEach(createStep => {
      steps.appendChild(createStep());
    });
    dialog.appendChild(steps);

    const footer = document.createElement('div');
    footer.className = 'ios-pwa-footer';

    const checkboxLabel = document.createElement('label');
    checkboxLabel.className = 'ios-pwa-checkbox';

    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = 'ios-pwa-dismiss-checkbox';
    checkboxLabel.appendChild(checkbox);

    const checkboxText = document.createElement('span');
    checkboxText.textContent = '再び表示しない';
    checkboxLabel.appendChild(checkboxText);

    footer.appendChild(checkboxLabel);

    const actions = document.createElement('div');
    actions.className = 'ios-pwa-actions';

    const closeButton = document.createElement('button');
    closeButton.type = 'button';
    closeButton.className = 'ios-pwa-close';
    closeButton.textContent = '閉じる';
    actions.appendChild(closeButton);
    footer.appendChild(actions);

    dialog.appendChild(footer);
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    function closePrompt() {
      if (checkbox.checked) {
        setPwaPromptDismissed();
      }
      window.removeEventListener('keydown', onKeyDown, true);
      overlay.remove();
    }

    function onKeyDown(event) {
      if (event.key === 'Escape') {
        event.preventDefault();
        closePrompt();
      }
    }

    window.addEventListener('keydown', onKeyDown, true);

    closeButton.addEventListener('click', closePrompt);
    overlay.addEventListener('click', event => {
      if (event.target === overlay) {
        closePrompt();
      }
    });

    setTimeout(() => {
      closeButton.focus();
    }, 0);
  }

  if (isIos() && !isStandaloneMode() && !isPwaPromptDismissed()) {
    showIosPwaPrompt();
  }

  function getLastSelectedStoreKey(stores) {
    if (!stores) {
      return null;
    }
    try {
      const key = localStorage.getItem(LAST_STORE_STORAGE_KEY);
      if (key && stores[key]) {
        return key;
      }
    } catch (e) {
      // Ignore storage access issues.
    }
    return null;
  }

  function setLastSelectedStoreKey(key) {
    if (!key) {
      return;
    }
    try {
      localStorage.setItem(LAST_STORE_STORAGE_KEY, key);
    } catch (e) {
      // Ignore storage access issues.
    }
  }

  function updateOfflineIndicator(message, variant) {
    if (!offlineInfo) return;
    offlineInfo.textContent = message || '';
    offlineInfo.classList.remove('is-success', 'is-error');
    if (variant === 'success') {
      offlineInfo.classList.add('is-success');
    } else if (variant === 'error') {
      offlineInfo.classList.add('is-error');
    }
  }

  function resetOfflineState() {
    if (typeof clearOfflineWorkbook === 'function') {
      try {
        clearOfflineWorkbook();
      } catch (e) {
        // Ignore cleanup failures; the offline cache is best-effort only.
      }
    }
    updateOfflineIndicator('');
    if (status) {
      status.textContent = '';
    }
  }

  resetOfflineState();

  window.addEventListener('pageshow', event => {
    if (event.persisted) {
      resetOfflineState();
    }
  });

  let offlineFileInput = null;
  if (offlineControls && offlineButton) {
    offlineFileInput = document.createElement('input');
    offlineFileInput.type = 'file';
    offlineFileInput.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    offlineFileInput.style.display = 'none';
    offlineControls.appendChild(offlineFileInput);

    offlineButton.addEventListener('click', () => {
      updateOfflineIndicator('');
      if (offlineFileInput) {
        offlineFileInput.value = '';
        offlineFileInput.click();
      }
    });

    offlineFileInput.addEventListener('change', () => {
      const file = offlineFileInput.files && offlineFileInput.files[0];
      if (!file) {
        return;
      }
      updateOfflineIndicator('');
      if (status) {
        startLoading(status, 'ローカルファイルを読み込み中・・・');
      }
      const reader = new FileReader();
      reader.onload = () => {
        stopLoading(status);
        try {
          if (typeof setOfflineWorkbook !== 'function') {
            throw new Error('Offline workbook is unavailable');
          }
          setOfflineWorkbook(reader.result, { fileName: file.name });
          updateOfflineIndicator('ローカルファイルを使用しています', 'success');
          const stores = typeof loadStores === 'function' ? loadStores() : null;
          const availableKeys = stores ? Object.keys(stores) : [];
          const storedKey = getLastSelectedStoreKey(stores);
          const targetKey = storedKey && availableKeys.includes(storedKey) ? storedKey : availableKeys[0];
          if (targetKey) {
            setLastSelectedStoreKey(targetKey);
            window.location.href = `sheets.html?store=${encodeURIComponent(targetKey)}&offline=1`;
          } else if (status) {
            status.textContent = '店舗情報が設定されていないためローカルファイルを開けません。設定を確認してください。';
          }
        } catch (e) {
          console.error('Failed to store offline workbook', e);
          updateOfflineIndicator('ローカルファイルを保存できませんでした。', 'error');
          if (status) {
            status.textContent = 'ローカルファイルを保存できませんでした。';
          }
          if (typeof window.notifyPlatformFeedback === 'function') {
            window.notifyPlatformFeedback(null, { feedbackLevel: 'error' });
          }
        } finally {
          offlineFileInput.value = '';
        }
      };
      reader.onerror = () => {
        stopLoading(status);
        console.error('Failed to read offline workbook', reader.error);
        updateOfflineIndicator('ローカルファイルの読み込みに失敗しました。', 'error');
        if (status) {
          status.textContent = 'ローカルファイルの読み込みに失敗しました。';
        }
        offlineFileInput.value = '';
        if (typeof window.notifyPlatformFeedback === 'function') {
          window.notifyPlatformFeedback(null, { feedbackLevel: 'error' });
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  startLoading(status, '読込中・・・');

  try {
    await ensureSettingsLoaded();
  } catch (e) {
    stopLoading(status);
    tutorialReady.resolve();
    if (list) {
      list.style.color = 'red';
      list.style.whiteSpace = 'pre-line';
      list.textContent = '店舗一覧の読み込みに失敗しました。\n通信環境をご確認のうえ、再度お試しください。';
    }
    if (typeof window.notifyPlatformFeedback === 'function') {
      window.notifyPlatformFeedback(null, { feedbackLevel: 'error' });
    }
    return;
  }
  document.getElementById('version').textContent = `ver.${APP_VERSION}`;
  const stores = loadStores();
  stopLoading(status);
  if (typeof window.notifyPlatformFeedback === 'function') {
    try {
      window.notifyPlatformFeedback(null, { feedbackLevel: 'success' });
    } catch (error) {
      console.warn('notifyPlatformFeedback failed', error);
    }
  }
  tutorialReady.resolve();
  if (status && typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive()) {
    status.textContent = '店舗を選択して続行してください。';
  }
  if (list) {
    list.textContent = '';
    list.style.color = '';
    list.style.whiteSpace = '';
  }
  storeButtonsHighlightTarget = null;
  const err = document.getElementById('settings-error');
  if (window.settingsError && err) {
    err.textContent = '設定が読み込めませんでした。\nデフォルトの値を使用します。\n設定からエラーを確認してください。';
  }
  const storeKeysForToast = stores ? Object.keys(stores) : [];
  const offlineActiveNow = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();
  let toastMessage = '';
  if (window.settingsError) {
    toastMessage = '設定を読み込めませんでした。デフォルトの店舗一覧を表示します。';
  } else if (storeKeysForToast.length === 0) {
    toastMessage = '店舗情報が見つかりませんでした。設定から店舗を登録してください。';
  } else if (offlineActiveNow) {
    toastMessage = '店舗一覧を読み込みました。ローカルファイルを利用できます。';
  } else {
    toastMessage = '店舗一覧の読み込みが完了しました。';
  }
  const feedbackLevelForToast = (window.settingsError || storeKeysForToast.length === 0)
    ? 'error'
    : 'success';
  showToastWithNativeNotice(toastMessage, { duration: 3200, feedbackLevel: feedbackLevelForToast });
  const storeKeys = Object.keys(stores);
  if (list && storeKeys.length > 0) {
    if (!document.getElementById('multi-store-overlay')) {
      const overlay = document.createElement('div');
      overlay.id = 'multi-store-overlay';
      overlay.style.display = 'none';

      const popup = document.createElement('div');
      popup.id = 'multi-store-popup';

      const title = document.createElement('h2');
      title.id = 'multi-store-title';
      title.textContent = '店舗横断計算モード';

      const description = document.createElement('p');
      description.id = 'multi-store-description';
      description.textContent = '計算したい店舗を複数選択してください。';

      const optionList = document.createElement('div');
      optionList.id = 'multi-store-list';

      const actions = document.createElement('div');
      actions.id = 'multi-store-actions';

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.id = 'multi-store-close';
      closeBtn.textContent = '閉じる';

      const startBtn = document.createElement('button');
      startBtn.type = 'button';
      startBtn.id = 'multi-store-start';
      startBtn.textContent = '読み込み開始';
      startBtn.disabled = true;

      actions.appendChild(closeBtn);
      actions.appendChild(startBtn);

      popup.appendChild(title);
      popup.appendChild(description);
      popup.appendChild(optionList);
      popup.appendChild(actions);

      overlay.appendChild(popup);
      document.body.appendChild(overlay);

      const selectedStores = new Set();
      const optionButtons = new Map();
      let tutorialPreviousSelection = null;

      function updateStartButton() {
        const count = selectedStores.size;
        startBtn.disabled = count === 0;
        startBtn.textContent = count === 0
          ? '読み込み開始'
          : `読み込み開始 (${count}件)`;
      }

      function toggleOverlay(show) {
        overlay.style.display = show ? 'flex' : 'none';
        if (!show && tutorialPreviousSelection) {
          restoreTutorialSelection();
        }
      }

      showMultiStoreOverlayForTutorial = () => toggleOverlay(true);
      hideMultiStoreOverlayForTutorial = () => toggleOverlay(false);

      closeBtn.addEventListener('click', () => {
        toggleOverlay(false);
      });

      overlay.addEventListener('click', event => {
        if (event.target === overlay) {
          toggleOverlay(false);
        }
      });

      function setOptionSelected(key, selected) {
        const button = optionButtons.get(key);
        if (!button) {
          return;
        }
        if (selected) {
          selectedStores.add(key);
          button.classList.add('is-selected');
        } else {
          selectedStores.delete(key);
          button.classList.remove('is-selected');
        }
        updateStartButton();
      }

      function applyTutorialSelection() {
        if (tutorialPreviousSelection) {
          return;
        }
        tutorialPreviousSelection = new Set(selectedStores);
        optionButtons.forEach(button => button.classList.remove('is-selected'));
        selectedStores.clear();
        const previewKeys = storeKeys.slice(0, 2);
        previewKeys.forEach(key => {
          const button = optionButtons.get(key);
          if (button) {
            button.classList.add('is-selected');
            selectedStores.add(key);
          }
        });
        updateStartButton();
      }

      function restoreTutorialSelection() {
        if (!tutorialPreviousSelection) {
          return;
        }
        optionButtons.forEach(button => button.classList.remove('is-selected'));
        selectedStores.clear();
        tutorialPreviousSelection.forEach(key => {
          const button = optionButtons.get(key);
          if (button) {
            button.classList.add('is-selected');
            selectedStores.add(key);
          }
        });
        tutorialPreviousSelection = null;
        updateStartButton();
      }

      storeKeys.forEach(key => {
        const storeInfo = stores[key];
        if (!storeInfo) {
          return;
        }
        const optionBtn = document.createElement('button');
        optionBtn.type = 'button';
        optionBtn.className = 'multi-store-option';
        optionBtn.textContent = storeInfo.name;
        optionBtn.dataset.storeKey = key;
        optionBtn.addEventListener('click', () => {
          const nextSelected = !selectedStores.has(key);
          setOptionSelected(key, nextSelected);
          if (tutorialPreviousSelection) {
            tutorialPreviousSelection = new Set(selectedStores);
          }
        });
        optionButtons.set(key, optionBtn);
        optionList.appendChild(optionBtn);
      });

      startBtn.addEventListener('click', () => {
        if (selectedStores.size === 0) {
          return;
        }
        const orderedKeys = Array.from(selectedStores);
        const params = new URLSearchParams();
        params.set('stores', orderedKeys.join(','));
        window.location.href = `sheets.html?${params.toString()}`;
      });

      const modeButton = document.createElement('button');
      modeButton.type = 'button';
      modeButton.id = 'multi-store-mode-button';
      modeButton.textContent = '店舗横断計算モード';
      modeButton.addEventListener('click', () => {
        toggleOverlay(true);
      });

      list.appendChild(modeButton);
      updateStartButton();

      multiStoreTutorialState = {
        getFocusElement: () => optionList,
        showPreview: applyTutorialSelection,
        clearPreview: restoreTutorialSelection
      };
    }

    if (!storeButtonsHighlightTarget) {
      const container = document.createElement('div');
      container.id = 'store-buttons-container';
      list.appendChild(container);
      storeButtonsHighlightTarget = container;
    }

    storeKeys.forEach(key => {
      const btn = document.createElement('button');
      btn.classList.add('store-button');
      btn.textContent = stores[key].name;
      btn.addEventListener('click', () => {
        setLastSelectedStoreKey(key);
        if (typeof setOfflineWorkbookActive === 'function') {
          try {
            setOfflineWorkbookActive(false);
          } catch (e) {
            // Ignore failures disabling offline mode.
          }
        }
        window.location.href = `sheets.html?store=${key}`;
      });
      const targetContainer = storeButtonsHighlightTarget || list;
      targetContainer.appendChild(btn);
    });
  }

  await loadTodayAttendance(stores);

  const infoBox = document.getElementById('announcements');
  if (infoBox) {
    infoBox.textContent = '';
    const header = document.createElement('div');
    header.style.textAlign = 'center';
    header.style.fontSize = '1.2rem';
    header.textContent = '●お知らせ●';
    infoBox.appendChild(header);

    const controlWrapper = document.createElement('div');
    infoBox.appendChild(controlWrapper);

    const messageDiv = document.createElement('div');
    messageDiv.textContent = 'お知らせの読み込み中';
    infoBox.appendChild(messageDiv);

    fetch('announcements.txt', { cache: 'no-store' })
      .then(res => res.text())
      .then(text => {
        const blocks = text
          .split(/\n\s*\n/)
          .map(block => block.trim())
          .filter(Boolean);
        const issues = [];
        const notes = [];
        blocks.forEach(block => {
          const lines = block.split('\n').map(l => l.trim()).filter(line => line.length > 0);
          if (!lines.length) {
            return;
          }
          const header = lines.shift();
          if (/^ver\./i.test(header)) {
            const version = header.replace(/^ver\./i, '').trim();
            notes.push({
              type: 'version',
              version,
              messages: lines
            });
            return;
          }
          if (header === '現在確認できている不具合') {
            issues.push({
              type: 'issues',
              title: header,
              messages: lines
            });
          }
        });

        notes.sort((a, b) => {
          const pa = a.version.split('.').map(Number);
          const pb = b.version.split('.').map(Number);
          for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
            const diff = (pb[i] || 0) - (pa[i] || 0);
            if (diff !== 0) return diff;
          }
          return 0;
        });

        if (!notes.length && !issues.length) {
          controlWrapper.textContent = '';
          messageDiv.textContent = '現在お知らせはありません。';
          return;
        }
        controlWrapper.textContent = '';
        const select = document.createElement('select');
        const latestOpt = document.createElement('option');
        latestOpt.value = '';
        latestOpt.textContent = '最新3件';
        select.appendChild(latestOpt);
        if (issues.length) {
          const issuesOpt = document.createElement('option');
          issuesOpt.value = '__issues__';
          issuesOpt.textContent = '現在確認できている不具合';
          select.appendChild(issuesOpt);
        }
        notes.forEach(n => {
          const opt = document.createElement('option');
          opt.value = n.version;
          opt.textContent = `ver.${n.version}`;
          select.appendChild(opt);
        });
        controlWrapper.appendChild(select);
        const latestNotes = notes.slice(0, 3);

        function buildNoteFragment(note) {
          const frag = document.createDocumentFragment();
          const strong = document.createElement('strong');
          strong.textContent = `ver.${note.version}`;
          frag.appendChild(strong);
          note.messages.forEach(msg => {
            frag.appendChild(document.createElement('br'));
            const span = document.createElement('span');
            span.textContent = msg;
            frag.appendChild(span);
          });
          return frag;
        }

        function buildIssuesFragment(block) {
          const frag = document.createDocumentFragment();
          const strong = document.createElement('strong');
          strong.textContent = block.title;
          frag.appendChild(strong);
          if (!block.messages.length) {
            return frag;
          }
          block.messages.forEach(msg => {
            frag.appendChild(document.createElement('br'));
            const span = document.createElement('span');
            span.textContent = msg;
            frag.appendChild(span);
          });
          return frag;
        }

        function render(ver) {
          messageDiv.textContent = '';
          if (!ver) {
            if (!latestNotes.length) {
              if (issues.length) {
                issues.forEach((block, idx) => {
                  messageDiv.appendChild(buildIssuesFragment(block));
                  if (idx < issues.length - 1) {
                    messageDiv.appendChild(document.createElement('br'));
                    messageDiv.appendChild(document.createElement('br'));
                  }
                });
              } else {
                const span = document.createElement('span');
                span.textContent = '現在お知らせはありません。';
                messageDiv.appendChild(span);
              }
              return;
            }
            latestNotes.forEach((n, idx) => {
              messageDiv.appendChild(buildNoteFragment(n));
              if (idx < latestNotes.length - 1) {
                messageDiv.appendChild(document.createElement('br'));
                messageDiv.appendChild(document.createElement('br'));
              }
            });
            return;
          }
          if (ver === '__issues__') {
            issues.forEach((block, idx) => {
              messageDiv.appendChild(buildIssuesFragment(block));
              if (idx < issues.length - 1) {
                messageDiv.appendChild(document.createElement('br'));
                messageDiv.appendChild(document.createElement('br'));
              }
            });
            if (!issues.length) {
              const span = document.createElement('span');
              span.textContent = '現在確認できている不具合はありません。';
              messageDiv.appendChild(span);
            }
            return;
          }
          const note = notes.find(n => n.version === ver);
          if (!note) return;
          messageDiv.appendChild(buildNoteFragment(note));
        }
        select.addEventListener('change', () => render(select.value));
        select.value = '';
        render(select.value);
      })
      .catch(() => {
        controlWrapper.textContent = '';
        messageDiv.textContent = 'お知らせを取得できませんでした。';
      });
  }

  initializeHelp('help/top.txt', {
    pageKey: 'top',
    showPrompt: true,
    enableAutoStartOnComplete: true,
    waitForReady: () => (tutorialReady.isResolved() ? true : tutorialReady.promise),
    onFinish: () => {
      hideMultiStoreOverlayForTutorial();
    },
    steps: {
      mode1: '#multi-store-mode-button',
      modePopup: {
        selector: '#multi-store-popup',
        onEnter: () => showMultiStoreOverlayForTutorial(),
        onExit: context => {
          if (!context || context.direction !== 'next') {
            hideMultiStoreOverlayForTutorial();
            if (multiStoreTutorialState && typeof multiStoreTutorialState.clearPreview === 'function') {
              multiStoreTutorialState.clearPreview();
            }
          }
        }
      },
      modeOptions: {
        getElement: () => {
          if (multiStoreTutorialState && typeof multiStoreTutorialState.getFocusElement === 'function') {
            const element = multiStoreTutorialState.getFocusElement();
            if (element instanceof Element) {
              return element;
            }
          }
          return document.getElementById('multi-store-list');
        },
        onEnter: () => {
          showMultiStoreOverlayForTutorial();
          if (multiStoreTutorialState && typeof multiStoreTutorialState.showPreview === 'function') {
            multiStoreTutorialState.showPreview();
          }
        },
        onExit: context => {
          if (!context || context.direction !== 'next') {
            hideMultiStoreOverlayForTutorial();
            if (multiStoreTutorialState && typeof multiStoreTutorialState.clearPreview === 'function') {
              multiStoreTutorialState.clearPreview();
            }
          }
        }
      },
      modeStart: {
        selector: '#multi-store-start',
        onEnter: () => showMultiStoreOverlayForTutorial(),
        onExit: context => {
          if (!context || context.direction !== 'next') {
            hideMultiStoreOverlayForTutorial();
            if (multiStoreTutorialState && typeof multiStoreTutorialState.clearPreview === 'function') {
              multiStoreTutorialState.clearPreview();
            }
          }
        }
      },
      modeClose: {
        selector: '#multi-store-close',
        onEnter: () => showMultiStoreOverlayForTutorial(),
        onExit: () => hideMultiStoreOverlayForTutorial()
      },
      stores: () => {
        if (storeButtonsHighlightTarget instanceof Element) {
          return storeButtonsHighlightTarget;
        }
        return document.querySelector('#store-list button.store-button') || document.getElementById('store-list');
      },
      local: '#offline-load-button',
      setting: '#settings',
      announcements: {
        selectors: ['#announcements select', '#announcements']
      },
      help: () => document.getElementById('help-button')
    }
  });

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.ready
      .then(registration => {
        if (!registration || !registration.active) {
          return;
        }
        registration.active.postMessage({
          type: 'WARMUP_CACHE',
          paths: [
            '/payroll.html',
            '/settings.html',
            '/sheets.html',
            '/payroll.js',
            '/settings.js',
            '/sheets.js',
            '/calc.js',
            '/help.js',
            '/help/payroll.txt',
            '/help/settings.txt',
            '/help/sheets.txt'
          ]
        });
      })
      .catch(() => {
        // Ignore failures warming the cache; navigation will still work without it.
      });
  }
});
