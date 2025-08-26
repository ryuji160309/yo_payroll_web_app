document.addEventListener('DOMContentLoaded', async () => {
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');
  const store = getStore(storeKey);
  if (!store) return;
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  try {
    const sheets = await fetchSheetList(store.url);
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
