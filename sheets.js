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
  const offlineIndicator = document.getElementById('offline-file-indicator');
  if (offlineIndicator) {
    offlineIndicator.classList.remove('is-success', 'is-error');
    const offlineActive = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();
    const info = typeof getOfflineWorkbookInfo === 'function' ? getOfflineWorkbookInfo() : null;
    if (offlineActive) {
      const label = info && info.fileName ? `ローカルファイル：${info.fileName}` : 'ローカルファイルを使用しています';
      offlineIndicator.textContent = label;
      offlineIndicator.classList.add('is-success');
    } else {
      offlineIndicator.textContent = '';
    }
  }
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');
  const store = getStore(storeKey);
  if (!store) {
    stopLoading(statusEl);
    return;
  }
  document.getElementById('store-name').textContent = store.name;
  try {
    const sheets = await fetchSheetList(store.url);
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
