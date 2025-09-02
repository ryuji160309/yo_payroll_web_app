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
  await settingsLoadPromise;
  initializeHelp('help/sheets.txt');
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');
  const store = getStore(storeKey);
  if (!store) return;
  document.getElementById('store-name').textContent = store.name;
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  try {
    const sheets = await fetchSheetList(store.url, storeKey);
    stopLoading(statusEl);
    const list = document.getElementById('sheet-list');

    sheets.forEach(({ name, index }) => {
      const btn = document.createElement('button');
      btn.textContent = formatSheetName(name);
      btn.addEventListener('click', () => {
        window.location.href = `payroll.html?store=${storeKey}&sheet=${index}`;

      });
      list.appendChild(btn);
    });
  } catch (e) {
    stopLoading(statusEl);
    document.getElementById('sheet-list').textContent = 'URLが変更された可能性があります。設定からURL変更をお試しください。';
  }
});
