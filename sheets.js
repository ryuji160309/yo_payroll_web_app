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

document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('status');

  const params = new URLSearchParams(location.search);
  const storesParamRaw = params.get('stores');
  const crossStoreMode = storesParamRaw !== null;

  startLoading(statusEl, crossStoreMode ? CROSS_STORE_LOADING_MESSAGE : '読込中・・・');
  initializeHelp('help/sheets.txt');
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
    statusEl.textContent = '店舗が選択されていません。トップページに戻ってやり直してください。';
    return;
  }

  const storeRecords = normalizedKeys
    .map(key => ({ key, store: getStore(key) }))
    .filter(record => !!record.store);

  if (storeRecords.length === 0) {
    stopLoading(statusEl);
    statusEl.textContent = '店舗情報を取得できませんでした。設定を確認してください。';
    return;
  }

  if (offlineMode && !offlineActive) {
    stopLoading(statusEl);
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

    const list = document.getElementById('sheet-list');
    if (list) {
      buildSheetSelectionInterface({
        list,
        stores: targetStores,
        crossStoreMode,
        offlineMode,
      });
    }

    if (failures.length > 0 && statusEl) {
      statusEl.textContent = '一部の店舗のシート一覧を読み込めませんでした。';
    } else if (statusEl) {
      statusEl.textContent = '';
    }
  } catch (e) {
    stopLoading(statusEl);
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

  function updateStartButton() {
    const count = selectedSheets.size;
    startBtn.disabled = count === 0;
    startBtn.textContent = count === 0 ? '計算開始' : `計算開始 (${count}件)`;
  }

  function toggleOverlay(show) {
    overlay.style.display = show ? 'flex' : 'none';
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
        if (selectedSheets.has(key)) {
          selectedSheets.delete(key);
          popupBtn.classList.remove('is-selected');
        } else {
          selectedSheets.add(key);
          popupBtn.classList.add('is-selected');
        }
        updateStartButton();
        updateSelectAllState();
      });
      popupButtons.push(popupBtn);
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
      selectedSheets.clear();
      popupButtons.forEach(btn => btn.classList.remove('is-selected'));
    } else {
      popupButtons.forEach(btn => {
        const key = `${btn.dataset.storeKey}|${btn.dataset.sheetIndex}`;
        selectedSheets.add(key);
        btn.classList.add('is-selected');
      });
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

  const modeButton = document.createElement('button');
  modeButton.type = 'button';
  modeButton.id = 'multi-month-mode-button';
  modeButton.textContent = '月横断計算モード';
  modeButton.addEventListener('click', () => {
    toggleOverlay(true);
  });
  list.appendChild(modeButton);

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
      }

      sectionsContainer.appendChild(section);
    });
  } else {
    const sheetButtonsContainer = document.createElement('div');
    sheetButtonsContainer.id = 'sheet-buttons-container';
    list.appendChild(sheetButtonsContainer);

    const sheetMeta = Array.isArray(stores[0].sheets) ? stores[0].sheets : [];
    if (sheetMeta.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'store-section-empty';
      empty.textContent = 'シートが見つかりません。';
      sheetButtonsContainer.appendChild(empty);
    } else {
      sheetMeta.forEach(meta => {
        const btn = document.createElement('button');
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
