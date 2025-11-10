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

document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  initializeHelp('help/sheets.txt');
  await ensureSettingsLoaded();
  const params = new URLSearchParams(location.search);
  const offlineMode = params.get('offline') === '1';
  const info = typeof getOfflineWorkbookInfo === 'function' ? getOfflineWorkbookInfo() : null;
  const offlineActive = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();
  const storeKey = params.get('store');
  const store = getStore(storeKey);
  if (!store) {
    stopLoading(statusEl);
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
  if (titleEl) {
    if (offlineMode && offlineActive && info && info.fileName) {
      titleEl.textContent = info.fileName;
      document.title = `${info.fileName} - シート選択`;
    } else {
      titleEl.textContent = store.name;
      document.title = `${store.name} - シート選択`;
    }
  }
  if (offlineMode && !offlineActive) {
    stopLoading(statusEl);
    statusEl.textContent = 'ローカルファイルを利用できません。もう一度トップに戻って読み込み直してください。';
    return;
  }
  try {
    const sheets = await fetchSheetList(store.url, { allowOffline: offlineMode });
    stopLoading(statusEl);
    const list = document.getElementById('sheet-list');

    const overlay = document.createElement('div');
    overlay.id = 'multi-month-overlay';
    overlay.style.display = 'none';

    const popup = document.createElement('div');
    popup.id = 'multi-month-popup';

    const popupTitle = document.createElement('h2');
    popupTitle.id = 'multi-month-title';
    popupTitle.textContent = '計算する月を選択';

    const popupDescription = document.createElement('p');
    popupDescription.id = 'multi-month-description';
    popupDescription.textContent = '計算したい月を選択してください。複数選択できます。';

    const popupList = document.createElement('div');
    popupList.id = 'multi-month-list';

    const popupActions = document.createElement('div');
    popupActions.id = 'multi-month-actions';

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
    popup.appendChild(popupActions);

    overlay.appendChild(popup);
    document.body.appendChild(overlay);

    const selectedSheets = new Set();

    function updateStartButton() {
      startBtn.disabled = selectedSheets.size === 0;
      if (selectedSheets.size === 0) {
        startBtn.textContent = '計算開始';
      } else {
        startBtn.textContent = `計算開始 (${selectedSheets.size}件)`;
      }
    }

    function toggleOverlay(show) {
      overlay.style.display = show ? 'flex' : 'none';
    }

    closeBtn.addEventListener('click', () => {
      toggleOverlay(false);
    });

    overlay.addEventListener('click', e => {
      if (e.target === overlay) {
        toggleOverlay(false);
      }
    });

    const modeButton = document.createElement('button');
    modeButton.type = 'button';
    modeButton.id = 'multi-month-mode-button';
    modeButton.textContent = '月横断計算モード';
    modeButton.addEventListener('click', () => {
      toggleOverlay(true);
    });

    list.appendChild(modeButton);

    const sheetButtonsContainer = document.createElement('div');
    sheetButtonsContainer.id = 'sheet-buttons-container';
    list.appendChild(sheetButtonsContainer);

    const popupButtons = [];

    sheets.forEach(({ name, index, sheetId }) => {
      const btn = document.createElement('button');
      btn.textContent = formatSheetName(name);
      btn.addEventListener('click', () => {
        const params = new URLSearchParams({ store: storeKey, sheet: index });
        if (sheetId !== undefined && sheetId !== null) {
          params.set('gid', sheetId);
        }
        if (offlineMode) {
          params.set('offline', '1');
        }
        window.location.href = `payroll.html?${params.toString()}`;
      });
      sheetButtonsContainer.appendChild(btn);

      const popupBtn = document.createElement('button');
      popupBtn.type = 'button';
      popupBtn.className = 'multi-month-option';
      popupBtn.textContent = formatSheetName(name);
      popupBtn.dataset.sheetIndex = String(index);
      popupBtn.dataset.sheetId = sheetId !== undefined && sheetId !== null ? String(sheetId) : '';
      popupBtn.addEventListener('click', () => {
        const sheetKey = index;
        if (selectedSheets.has(sheetKey)) {
          selectedSheets.delete(sheetKey);
          popupBtn.classList.remove('is-selected');
        } else {
          selectedSheets.add(sheetKey);
          popupBtn.classList.add('is-selected');
        }
        updateStartButton();
      });
      popupButtons.push(popupBtn);
      popupList.appendChild(popupBtn);
    });

    startBtn.addEventListener('click', () => {
      if (selectedSheets.size === 0) {
        return;
      }
      const sortedIndices = Array.from(selectedSheets).sort((a, b) => a - b);
      const params = new URLSearchParams({ store: storeKey, sheets: sortedIndices.join(',') });
      if (offlineMode) {
        params.set('offline', '1');
      }
      const gidValues = sortedIndices.map(idx => {
        const btn = popupButtons.find(b => Number(b.dataset.sheetIndex) === idx);
        return btn ? btn.dataset.sheetId || '' : '';
      });
      if (gidValues.some(id => id !== '')) {
        params.set('gids', gidValues.join(','));
      }
      toggleOverlay(false);
      window.location.href = `payroll.html?${params.toString()}`;
    });
  } catch (e) {
    stopLoading(statusEl);
    const listEl = document.getElementById('sheet-list');
    listEl.style.color = 'red';
    listEl.style.whiteSpace = 'pre-line';
    listEl.textContent = 'シート一覽が読み込めませんでした。\nURLが間違っていないか設定から確認ください。';
  }
});
