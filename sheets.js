function formatSheetName(name) {
  const withYear = name.match(/^(\d{2})\.(\d{1,2})16-(\d{1,2})15$/);
  if (withYear) {
    const [, yy, sm, em] = withYear;
    return `20${yy}年${parseInt(sm)}月16日～${parseInt(em)}月15日`;
  }
  const withoutYear = name.match(/^(\d{1,2})16-(\d{1,2})15$/);
  if (withoutYear) {
    const [, sm, em] = withoutYear;
    return `${parseInt(sm)}月16日～${parseInt(em)}月15日`;
  }
  return name;
}

const CROSS_STORE_LOADING_MESSAGE = [
  '店舗横断計算モードでは複数店舗のデータを読み込むため、通常の計算よりも時間がかかります。',
  'しばらくお待ち下さい。'
].join('\n');

let sheetButtonsHighlightTarget = null;
let multiMonthTutorialState = null;

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
  const statusEl = document.getElementById('status');

  const params = new URLSearchParams(location.search);
  const storesParamRaw = params.get('stores');
  const crossStoreMode = storesParamRaw !== null;

  const ensureMultiMonthOverlayOpen = () => {
    const overlay = document.getElementById('multi-month-overlay');
    if (!overlay) {
      return;
    }
    const style = window.getComputedStyle(overlay);
    if (style.display !== 'flex') {
      const modeBtn = document.getElementById('multi-month-mode-button');
      if (modeBtn) {
        modeBtn.click();
      }
    }
  };

  const closeMultiMonthOverlay = () => {
    const overlay = document.getElementById('multi-month-overlay');
    if (!overlay) {
      if (multiMonthTutorialState && typeof multiMonthTutorialState.restorePreview === 'function') {
        multiMonthTutorialState.restorePreview();
      }
      return;
    }
    const style = window.getComputedStyle(overlay);
    if (style.display === 'flex') {
      const closeBtn = document.getElementById('multi-month-close');
      if (closeBtn) {
        closeBtn.click();
      } else {
        overlay.style.display = 'none';
      }
    }
    if (multiMonthTutorialState && typeof multiMonthTutorialState.restorePreview === 'function') {
      multiMonthTutorialState.restorePreview();
    }
  };

  const resolveSheetButtonsHighlightTarget = () => {
    if (sheetButtonsHighlightTarget instanceof Element) {
      return sheetButtonsHighlightTarget;
    }
    return document.getElementById('sheet-buttons-container')
      || document.querySelector('.store-sheet-buttons')
      || document.querySelector('#sheet-list .sheet-button')
      || document.getElementById('sheet-list');
  };

  startLoading(
    statusEl,
    crossStoreMode ? CROSS_STORE_LOADING_MESSAGE : '読込中・・・',
    { disableSlowNote: crossStoreMode }
  );
  initializeHelp('help/sheets.txt', {
    pageKey: 'sheets',
    showPrompt: false,
    autoStartIf: ({ hasAutoStartFlag }) => hasAutoStartFlag,
    waitForReady: () => (tutorialReady.isResolved() ? true : tutorialReady.promise),
    onStart: closeMultiMonthOverlay,
    onFinish: () => {
      closeMultiMonthOverlay();
    },
    steps: {
      back: '#sheets-back',
      restart: {
        selector: '#sheets-home',
        onExit: context => {
          if (context && context.direction === 'next') {
            closeMultiMonthOverlay();
          }
        }
      },
      mode: {
        selector: '#multi-month-mode-button',
        onEnter: closeMultiMonthOverlay
      },
      modePopup: {
        selector: '#multi-month-popup',
        onEnter: ensureMultiMonthOverlayOpen,
        onExit: context => {
          if (!context || context.direction !== 'next') {
            closeMultiMonthOverlay();
          }
        }
      },
      modeSelection: {
        selector: '#multi-month-list',
        onEnter: () => {
          ensureMultiMonthOverlayOpen();
          if (multiMonthTutorialState && typeof multiMonthTutorialState.previewSelectTopTwo === 'function') {
            multiMonthTutorialState.previewSelectTopTwo();
          }
        },
        onExit: context => {
          if (context && context.direction === 'next') {
            if (multiMonthTutorialState && typeof multiMonthTutorialState.previewClearSelection === 'function') {
              multiMonthTutorialState.previewClearSelection();
            }
            return;
          }
          closeMultiMonthOverlay();
        }
      },
      selectAll: {
        selector: '#multi-month-select-all',
        onEnter: () => {
          ensureMultiMonthOverlayOpen();
          if (multiMonthTutorialState && typeof multiMonthTutorialState.previewClearSelection === 'function') {
            multiMonthTutorialState.previewClearSelection();
          }
        },
        onExit: context => {
          if (multiMonthTutorialState && typeof multiMonthTutorialState.previewSelectAll === 'function' && context && context.direction === 'next') {
            multiMonthTutorialState.previewSelectAll();
            return;
          }
          if (multiMonthTutorialState && typeof multiMonthTutorialState.restorePreview === 'function') {
            multiMonthTutorialState.restorePreview();
          }
        }
      },
      selectAllToggle: {
        selector: '#multi-month-select-all',
        onEnter: () => {
          ensureMultiMonthOverlayOpen();
          if (multiMonthTutorialState && typeof multiMonthTutorialState.previewSelectAll === 'function') {
            multiMonthTutorialState.previewSelectAll();
          }
        },
        onExit: () => {
          if (multiMonthTutorialState && typeof multiMonthTutorialState.previewClearSelection === 'function') {
            multiMonthTutorialState.previewClearSelection();
          }
        }
      },
      start: {
        selector: '#multi-month-start',
        onEnter: () => {
          ensureMultiMonthOverlayOpen();
          if (multiMonthTutorialState && typeof multiMonthTutorialState.previewSelectTopTwo === 'function') {
            multiMonthTutorialState.previewSelectTopTwo();
          }
        },
        onExit: context => {
          if (!context || context.direction !== 'next') {
            if (multiMonthTutorialState && typeof multiMonthTutorialState.restorePreview === 'function') {
              multiMonthTutorialState.restorePreview();
            }
          }
        }
      },
      close: {
        selector: '#multi-month-close',
        onEnter: ensureMultiMonthOverlayOpen,
        onExit: context => {
          if (context && context.direction === 'next') {
            if (multiMonthTutorialState && typeof multiMonthTutorialState.previewClearSelection === 'function') {
              multiMonthTutorialState.previewClearSelection();
            }
          } else if (multiMonthTutorialState && typeof multiMonthTutorialState.restorePreview === 'function') {
            multiMonthTutorialState.restorePreview();
          }
          closeMultiMonthOverlay();
        }
      },
      sheets: () => resolveSheetButtonsHighlightTarget(),
      help: () => document.getElementById('help-button')
    }
  });
  await ensureSettingsLoaded();

  const storeParamRaw = params.get('store');
  const offlineRequested = params.get('offline') === '1';
  const offlineMode = offlineRequested && !crossStoreMode;
  const info = typeof getOfflineWorkbookInfo === 'function' ? getOfflineWorkbookInfo() : null;
  const offlineActive = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();

  function normalizeStoreKey(value) {
    if (value === null || value === undefined) {
      return '';
    }
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }
    try {
      return decodeURIComponent(trimmed);
    } catch (e) {
      return trimmed;
    }
  }

  const requestedKeys = [];
  if (storesParamRaw) {
    storesParamRaw.split(',').forEach(part => {
      const key = normalizeStoreKey(part);
      if (key) {
        requestedKeys.push(key);
      }
    });
  }
  if (!crossStoreMode && storeParamRaw) {
    const key = normalizeStoreKey(storeParamRaw);
    if (key) {
      requestedKeys.push(key);
    }
  }

  const normalizedKeys = [];
  const seenKeys = new Set();
  requestedKeys.forEach(key => {
    if (!key || seenKeys.has(key)) {
      return;
    }
    seenKeys.add(key);
    normalizedKeys.push(key);
  });

  if (normalizedKeys.length === 0) {
    stopLoading(statusEl);
    tutorialReady.resolve();
    statusEl.textContent = '店舗が選択されていません。トップページに戻ってやり直してください。';
    return;
  }

  const storeRecords = normalizedKeys
    .map(key => ({ key, store: getStore(key) }))
    .filter(record => !!record.store);

  if (storeRecords.length === 0) {
    stopLoading(statusEl);
    tutorialReady.resolve();
    statusEl.textContent = '店舗情報を取得できませんでした。設定を確認してください。';
    return;
  }

  if (offlineMode && !offlineActive) {
    stopLoading(statusEl);
    tutorialReady.resolve();
    statusEl.textContent = 'ローカルファイルを利用できません。もう一度トップに戻って読み込み直してください。';
    return;
  }

  if (!offlineMode && typeof setOfflineWorkbookActive === 'function') {
    try {
      setOfflineWorkbookActive(false);
    } catch (e) {
      // Ignore errors while disabling offline mode.
    }
  }

  const titleEl = document.getElementById('store-name');
  if (crossStoreMode) {
    if (titleEl) {
      titleEl.textContent = '店舗横断計算モード';
    }
    document.title = '店舗横断計算モード - シート選択';
  } else {
    const primaryStore = storeRecords[0].store;
    if (titleEl) {
      if (offlineMode && offlineActive && info && info.fileName) {
        titleEl.textContent = info.fileName;
        document.title = `${info.fileName} - シート選択`;
      } else {
        titleEl.textContent = primaryStore.name;
        document.title = `${primaryStore.name} - シート選択`;
      }
    }
  }

  try {
    let targetStores = [];
    const failures = [];

    if (crossStoreMode) {
      const results = await Promise.allSettled(
        storeRecords.map(record => fetchSheetList(record.store.url, { allowOffline: false }))
      );
      results.forEach((result, idx) => {
        const record = storeRecords[idx];
        if (!record) {
          return;
        }
        if (result.status === 'fulfilled') {
          targetStores.push({ ...record, sheets: result.value });
        } else {
          failures.push({ record, reason: result.reason });
        }
      });
    } else {
      const record = storeRecords[0];
      const sheets = await fetchSheetList(record.store.url, { allowOffline: offlineMode });
      targetStores.push({ ...record, sheets });
    }

    targetStores = targetStores.filter(entry => Array.isArray(entry.sheets));

    if (targetStores.length === 0) {
      throw new Error('no-sheets');
    }

    stopLoading(statusEl);
    tutorialReady.resolve();

    const list = document.getElementById('sheet-list');
    if (list) {
      buildSheetSelectionInterface({
        list,
        stores: targetStores,
        crossStoreMode,
        offlineMode,
      });
    }

    const availableStoreNames = targetStores
      .map(entry => (entry && entry.store && entry.store.name ? entry.store.name : ''))
      .filter(name => !!name);
    let toastMessage = '';
    if (offlineMode && offlineActive && info && info.fileName) {
      toastMessage = `${info.fileName} のシート一覧を読み込みました。`;
    } else if (crossStoreMode) {
      if (availableStoreNames.length === 0) {
        toastMessage = 'シート一覧の読み込みが完了しました。';
      } else if (availableStoreNames.length <= 3) {
        toastMessage = `${availableStoreNames.join('・')} のシート一覧を読み込みました。`;
      } else {
        toastMessage = `${availableStoreNames.length}店舗のシート一覧を読み込みました。`;
      }
    } else {
      const primaryName = availableStoreNames[0] || '';
      toastMessage = primaryName
        ? `${primaryName} のシート一覧を読み込みました。`
        : 'シート一覧の読み込みが完了しました。';
    }
    if (failures.length > 0) {
      toastMessage += '（一部の店舗は読み込めませんでした）';
    }
    showToastWithNativeNotice(toastMessage, { duration: 3200 });

    if (failures.length > 0 && statusEl) {
      statusEl.textContent = '一部の店舗のシート一覧を読み込めませんでした。';
    } else if (statusEl) {
      statusEl.textContent = '';
    }
  } catch (e) {
    stopLoading(statusEl);
    tutorialReady.resolve();
    const listEl = document.getElementById('sheet-list');
    if (listEl) {
      listEl.style.color = 'red';
      listEl.style.whiteSpace = 'pre-line';
      listEl.textContent = 'シート一覽が読み込めませんでした。\nURLが間違っていないか設定から確認ください。';
    }
  }
});

function buildSheetSelectionInterface({ list, stores, crossStoreMode, offlineMode }) {
  list.innerHTML = '';
  sheetButtonsHighlightTarget = null;

  const overlay = document.createElement('div');
  overlay.id = 'multi-month-overlay';
  overlay.style.display = 'none';

  const popup = document.createElement('div');
  popup.id = 'multi-month-popup';

  const popupTitle = document.createElement('h2');
  popupTitle.id = 'multi-month-title';
  popupTitle.textContent = '月横断計算モード';

  const popupDescription = document.createElement('p');
  popupDescription.id = 'multi-month-description';
  popupDescription.textContent = crossStoreMode
    ? '計算したい店舗・月を複数選択してください。'
    : '計算したい月を複数選択してください。';

  const popupList = document.createElement('div');
  popupList.id = 'multi-month-list';

  const popupActions = document.createElement('div');
  popupActions.id = 'multi-month-actions';

  const selectAllBtn = document.createElement('button');
  selectAllBtn.type = 'button';
  selectAllBtn.id = 'multi-month-select-all';
  selectAllBtn.textContent = '全選択';

  const closeBtn = document.createElement('button');
  closeBtn.type = 'button';
  closeBtn.id = 'multi-month-close';
  closeBtn.textContent = '閉じる';

  const startBtn = document.createElement('button');
  startBtn.type = 'button';
  startBtn.id = 'multi-month-start';
  startBtn.textContent = '計算開始';
  startBtn.disabled = true;

  popupActions.appendChild(closeBtn);
  popupActions.appendChild(startBtn);

  popup.appendChild(popupTitle);
  popup.appendChild(popupDescription);
  popup.appendChild(popupList);
  popup.appendChild(selectAllBtn);
  popup.appendChild(popupActions);

  overlay.appendChild(popup);
  document.body.appendChild(overlay);

  const selectedSheets = new Set();
  const popupButtons = [];
  const popupButtonMap = new Map();
  let tutorialSnapshot = null;
  let tutorialPreviewActive = false;
  let lastPreviewClearTimestamp = 0;
  let pendingTopTwoTimer = null;

  function cancelPendingTopTwo() {
    if (pendingTopTwoTimer !== null) {
      clearTimeout(pendingTopTwoTimer);
      pendingTopTwoTimer = null;
    }
  }

  function setSheetSelectedByKey(key, selected) {
    const btn = popupButtonMap.get(key);
    if (!btn) {
      return;
    }
    if (selected) {
      selectedSheets.add(key);
      btn.classList.add('is-selected');
    } else {
      selectedSheets.delete(key);
      btn.classList.remove('is-selected');
    }
  }

  function clearAllSheets() {
    selectedSheets.clear();
    popupButtons.forEach(btn => btn.classList.remove('is-selected'));
  }

  function selectAllSheets() {
    popupButtons.forEach(btn => {
      const key = `${btn.dataset.storeKey}|${btn.dataset.sheetIndex}`;
      selectedSheets.add(key);
      btn.classList.add('is-selected');
    });
  }

  function ensureTutorialSnapshot() {
    if (!tutorialPreviewActive) {
      tutorialSnapshot = new Set(selectedSheets);
      tutorialPreviewActive = true;
    }
  }

  function restoreTutorialPreview() {
    if (!tutorialPreviewActive) {
      return;
    }
    cancelPendingTopTwo();
    lastPreviewClearTimestamp = 0;
    clearAllSheets();
    if (tutorialSnapshot) {
      tutorialSnapshot.forEach(key => {
        setSheetSelectedByKey(key, true);
      });
    }
    tutorialSnapshot = null;
    tutorialPreviewActive = false;
    updateStartButton();
    updateSelectAllState();
  }

  function previewClearSelection() {
    ensureTutorialSnapshot();
    cancelPendingTopTwo();
    clearAllSheets();
    updateStartButton();
    updateSelectAllState();
    lastPreviewClearTimestamp = Date.now();
  }

  function previewSelectAll() {
    ensureTutorialSnapshot();
    cancelPendingTopTwo();
    selectAllSheets();
    updateStartButton();
    updateSelectAllState();
    lastPreviewClearTimestamp = 0;
  }

  function previewSelectTopTwo() {
    ensureTutorialSnapshot();
    const focusTarget = document.activeElement instanceof HTMLElement
      && document.activeElement !== document.body
      ? document.activeElement
      : null;
    const restoreFocus = () => {
      if (!focusTarget || !document.contains(focusTarget)) {
        return;
      }
      requestAnimationFrame(() => {
        if (!document.contains(focusTarget)) {
          return;
        }
        try {
          focusTarget.focus({ preventScroll: true });
        } catch (error) {
          focusTarget.focus();
        }
      });
    };
    const applySelection = () => {
      cancelPendingTopTwo();
      clearAllSheets();
      for (let i = 0; i < popupButtons.length && i < 2; i += 1) {
        const btn = popupButtons[i];
        if (!btn) {
          continue;
        }
        const key = `${btn.dataset.storeKey}|${btn.dataset.sheetIndex}`;
        setSheetSelectedByKey(key, true);
      }
      updateStartButton();
      updateSelectAllState();
      lastPreviewClearTimestamp = 0;
      restoreFocus();
    };

    const now = Date.now();
    if (lastPreviewClearTimestamp && now - lastPreviewClearTimestamp < 150) {
      cancelPendingTopTwo();
      pendingTopTwoTimer = setTimeout(applySelection, 320);
      return;
    }

    applySelection();
  }

  function updateStartButton() {
    const count = selectedSheets.size;
    startBtn.disabled = count === 0;
    startBtn.textContent = count === 0 ? '計算開始' : `計算開始 (${count}件)`;
  }

  function toggleOverlay(show) {
    overlay.style.display = show ? 'flex' : 'none';
    if (!show) {
      restoreTutorialPreview();
    }
  }

  closeBtn.addEventListener('click', () => {
    toggleOverlay(false);
  });

  overlay.addEventListener('click', event => {
    if (event.target === overlay) {
      toggleOverlay(false);
    }
  });

  stores.forEach((entry, storeOrder) => {
    const sheetMeta = Array.isArray(entry.sheets) ? entry.sheets : [];
    sheetMeta.forEach(meta => {
      const key = `${entry.key}|${meta.index}`;
      const popupBtn = document.createElement('button');
      popupBtn.type = 'button';
      popupBtn.className = 'multi-month-option';
      popupBtn.textContent = crossStoreMode
        ? `${entry.store.name}：${formatSheetName(meta.name)}`
        : formatSheetName(meta.name);
      popupBtn.dataset.storeKey = entry.key;
      popupBtn.dataset.sheetIndex = String(meta.index);
      popupBtn.dataset.sheetId = meta.sheetId !== undefined && meta.sheetId !== null ? String(meta.sheetId) : '';
      popupBtn.dataset.storeOrder = String(storeOrder);
      popupBtn.addEventListener('click', () => {
        const nextSelected = !selectedSheets.has(key);
        setSheetSelectedByKey(key, nextSelected);
        updateStartButton();
        updateSelectAllState();
      });
      popupButtons.push(popupBtn);
      popupButtonMap.set(key, popupBtn);
      popupList.appendChild(popupBtn);
    });
  });

  function areAllSelected() {
    if (popupButtons.length === 0) {
      return false;
    }
    return popupButtons.every(btn => {
      const key = `${btn.dataset.storeKey}|${btn.dataset.sheetIndex}`;
      return selectedSheets.has(key);
    });
  }

  function updateSelectAllState() {
    selectAllBtn.textContent = areAllSelected() ? '選択解除' : '全選択';
  }

  selectAllBtn.addEventListener('click', () => {
    if (areAllSelected()) {
      clearAllSheets();
    } else {
      selectAllSheets();
    }
    updateStartButton();
    updateSelectAllState();
  });

  updateSelectAllState();
  updateStartButton();

  startBtn.addEventListener('click', () => {
    if (selectedSheets.size === 0) {
      return;
    }
    const selectionDetails = Array.from(selectedSheets)
      .map(identifier => {
        const [storeKey, sheetIndexStr] = identifier.split('|');
        const sheetIndex = Number(sheetIndexStr);
        if (!storeKey || !Number.isFinite(sheetIndex)) {
          return null;
        }
        const btn = popupButtons.find(
          candidate => candidate.dataset.storeKey === storeKey && Number(candidate.dataset.sheetIndex) === sheetIndex
        );
        if (!btn) {
          return null;
        }
        const storeOrder = Number(btn.dataset.storeOrder || '0');
        const sheetId = btn.dataset.sheetId || '';
        return { storeKey, sheetIndex, sheetId, storeOrder };
      })
      .filter(Boolean);

    if (selectionDetails.length === 0) {
      return;
    }

    selectionDetails.sort((a, b) => {
      if (a.storeOrder !== b.storeOrder) {
        return a.storeOrder - b.storeOrder;
      }
      return a.sheetIndex - b.sheetIndex;
    });

    const query = new URLSearchParams();
    if (crossStoreMode) {
      const orderedStores = [];
      selectionDetails.forEach(detail => {
        if (!orderedStores.includes(detail.storeKey)) {
          orderedStores.push(detail.storeKey);
        }
      });
      query.set('stores', orderedStores.join(','));
      const segments = selectionDetails.map(detail => {
        let segment = `${encodeURIComponent(detail.storeKey)}:${detail.sheetIndex}`;
        if (detail.sheetId) {
          segment += `:${detail.sheetId}`;
        }
        return segment;
      });
      query.set('selections', segments.join(','));
    } else {
      const storeKey = stores[0] ? stores[0].key : '';
      query.set('store', storeKey);
      const sheetIndices = selectionDetails.map(detail => detail.sheetIndex);
      query.set('sheets', sheetIndices.join(','));
      if (offlineMode) {
        query.set('offline', '1');
      }
      const gidValues = selectionDetails.map(detail => detail.sheetId || '');
      if (gidValues.some(id => id !== '')) {
        query.set('gids', gidValues.join(','));
      }
    }

    toggleOverlay(false);
    window.location.href = `payroll.html?${query.toString()}`;
  });

  multiMonthTutorialState = {
    previewClearSelection,
    previewSelectAll,
    previewSelectTopTwo,
    restorePreview: restoreTutorialPreview
  };

  const modeButton = document.createElement('button');
  modeButton.type = 'button';
  modeButton.id = 'multi-month-mode-button';
  modeButton.textContent = '月横断計算モード';
  modeButton.addEventListener('click', () => {
    toggleOverlay(true);
  });
  list.appendChild(modeButton);

  if (crossStoreMode) {
    requestAnimationFrame(() => toggleOverlay(true));
  }

  if (crossStoreMode) {
    const sectionsContainer = document.createElement('div');
    sectionsContainer.id = 'store-sheet-sections';
    list.appendChild(sectionsContainer);

    stores.forEach(entry => {
      const section = document.createElement('section');
      section.className = 'store-section';

      const heading = document.createElement('h3');
      heading.className = 'store-section-title';
      heading.textContent = entry.store.name;
      section.appendChild(heading);

      const buttonWrapper = document.createElement('div');
      buttonWrapper.className = 'store-sheet-buttons';

      const sheetMeta = Array.isArray(entry.sheets) ? entry.sheets : [];
      if (sheetMeta.length === 0) {
        const empty = document.createElement('p');
        empty.className = 'store-section-empty';
        empty.textContent = 'シートが見つかりません。';
        section.appendChild(empty);
      } else {
        sheetMeta.forEach(meta => {
          const btn = document.createElement('button');
          btn.classList.add('sheet-button');
          btn.textContent = formatSheetName(meta.name);
          btn.addEventListener('click', () => {
            const params = new URLSearchParams({ store: entry.key, sheet: meta.index });
            if (meta.sheetId !== undefined && meta.sheetId !== null) {
              params.set('gid', String(meta.sheetId));
            }
            window.location.href = `payroll.html?${params.toString()}`;
          });
          buttonWrapper.appendChild(btn);
        });
        section.appendChild(buttonWrapper);
        if (!(sheetButtonsHighlightTarget instanceof Element)) {
          sheetButtonsHighlightTarget = buttonWrapper;
        }
      }

      sectionsContainer.appendChild(section);
    });
  } else {
    const sheetButtonsContainer = document.createElement('div');
    sheetButtonsContainer.id = 'sheet-buttons-container';
    list.appendChild(sheetButtonsContainer);
    sheetButtonsHighlightTarget = sheetButtonsContainer;

    const sheetMeta = Array.isArray(stores[0].sheets) ? stores[0].sheets : [];
    if (sheetMeta.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'store-section-empty';
      empty.textContent = 'シートが見つかりません。';
      sheetButtonsContainer.appendChild(empty);
    } else {
      sheetMeta.forEach(meta => {
        const btn = document.createElement('button');
        btn.classList.add('sheet-button');
        btn.textContent = formatSheetName(meta.name);
        btn.addEventListener('click', () => {
          const params = new URLSearchParams({ store: stores[0].key, sheet: meta.index });
          if (meta.sheetId !== undefined && meta.sheetId !== null) {
            params.set('gid', String(meta.sheetId));
          }
          if (offlineMode) {
            params.set('offline', '1');
          }
          window.location.href = `payroll.html?${params.toString()}`;
        });
        sheetButtonsContainer.appendChild(btn);
      });
    }
  }
}
