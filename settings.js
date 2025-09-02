document.addEventListener('DOMContentLoaded', async () => {
  await settingsLoadPromise;
  initializeHelp('help/settings.txt');
  const select = document.getElementById('store-select');
  Object.keys(DEFAULT_STORES).forEach(key => {
    const opt = document.createElement('option');
    opt.value = key;
    opt.textContent = DEFAULT_STORES[key].name;
    select.appendChild(opt);
  });

  function load(key) {
    const store = DEFAULT_STORES[key] || {};
    document.getElementById('url').value = store.url || '';
    document.getElementById('baseWage').value = store.baseWage;
    document.getElementById('overtime').value = store.overtime;
    document.getElementById('excludeWords').value = (store.excludeWords || []).join(',');
  }

  select.addEventListener('change', () => load(select.value));
  select.value = Object.keys(DEFAULT_STORES)[0];
  load(select.value);
});
