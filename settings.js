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
    const def = DEFAULT_STORES[key];
    document.getElementById('url').value = store.url;
    document.getElementById('baseWage').value = Number.isFinite(store.baseWage) ? store.baseWage : def.baseWage;
    document.getElementById('overtime').value = Number.isFinite(store.overtime) ? store.overtime : def.overtime;
    const words = Array.isArray(store.excludeWords) ? store.excludeWords : def.excludeWords;
    document.getElementById('excludeWords').value = (words || []).join(',');
  }

  select.addEventListener('change', () => load(select.value));
  select.value = Object.keys(DEFAULT_STORES)[0];
  load(select.value);

  document.getElementById('save').addEventListener('click', () => {
    const baseWage = Number(document.getElementById('baseWage').value);
    const overtime = Number(document.getElementById('overtime').value);
    if (!Number.isFinite(baseWage) || !Number.isFinite(overtime)) {
      alert('数値を入力してください');
      return;
    }
    try {
      updateStore(select.value, {
        url: document.getElementById('url').value,
        baseWage,
        overtime,
        excludeWords: document.getElementById('excludeWords').value.split(',').map(s => s.trim()).filter(s => s)
      });
      const saved = getStore(select.value);
      if (saved.baseWage !== baseWage || saved.overtime !== overtime) {
        throw new Error('verify failed');
      }
      alert('保存しました');
    } catch (e) {
      alert('保存に失敗しました。ブラウザの設定を確認してください。');
    }
  });

  document.getElementById('reset').addEventListener('click', () => {
    updateStore(select.value, DEFAULT_STORES[select.value]);
    load(select.value);
  });
});
