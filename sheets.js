document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('version').textContent = `ver.${APP_VERSION}`;
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');
  const store = getStore(storeKey);
  if (!store) return;
  try {
    const sheets = await fetchSheetList(store.url);
    const list = document.getElementById('sheet-list');
    sheets.forEach(({ name, index }) => {
      const btn = document.createElement('button');
      btn.textContent = name;
      btn.addEventListener('click', () => {
        window.location.href = `payroll.html?store=${storeKey}&sheet=${index}`;
      });
      list.appendChild(btn);
    });
  } catch (e) {
    document.getElementById('sheet-list').textContent = 'URLが変更された可能性があります。設定からURL変更をお試しください。';
  }
});
