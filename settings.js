document.addEventListener('DOMContentLoaded', () => {
  const select = document.getElementById('store-select');
  const stores = loadStores();
  Object.keys(DEFAULT_STORES).forEach(key => {
    const opt = document.createElement('option');
    opt.value = key;
    opt.textContent = stores[key].name;
    select.appendChild(opt);
  });

  function load(key) {
    const store = getStore(key);
    document.getElementById('url').value = store.url;
    document.getElementById('baseWage').value = store.baseWage;
    document.getElementById('overtime').value = store.overtime;
    document.getElementById('excludeWords').value = (store.excludeWords || []).join(',');
  }

  select.addEventListener('change', () => load(select.value));
  select.value = Object.keys(DEFAULT_STORES)[0];
  load(select.value);

  document.getElementById('save').addEventListener('click', () => {
    updateStore(select.value, {
      url: document.getElementById('url').value,
      baseWage: parseFloat(document.getElementById('baseWage').value),
      overtime: parseFloat(document.getElementById('overtime').value),
      excludeWords: document.getElementById('excludeWords').value.split(',').map(s => s.trim()).filter(s => s)
    });
    alert('保存しました');
  });

  document.getElementById('reset').addEventListener('click', () => {
    updateStore(select.value, DEFAULT_STORES[select.value]);
    load(select.value);
  });
});
