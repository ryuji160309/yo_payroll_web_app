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
  const offlineIndicator = document.getElementById('offline-file-indicator');
  if (offlineIndicator) {
    offlineIndicator.classList.remove('is-success', 'is-error');
    if (offlineMode && offlineActive) {
      offlineIndicator.textContent = 'ローカルファイルを使用しています';
      offlineIndicator.classList.add('is-success');
    } else {
      offlineIndicator.textContent = '';
    }
  }
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
      list.appendChild(btn);
    });
  } catch (e) {
    stopLoading(statusEl);
    const listEl = document.getElementById('sheet-list');
    listEl.style.color = 'red';
    listEl.style.whiteSpace = 'pre-line';
    listEl.textContent = 'シート一覽が読み込めませんでした。\nURLが間違っていないか設定から確認ください。';
  }
});
