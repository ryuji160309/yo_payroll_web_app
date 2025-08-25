document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('version').textContent = `ver.${APP_VERSION}`;
  const list = document.getElementById('store-list');
  const stores = loadStores();
  Object.keys(stores).forEach(key => {
    const btn = document.createElement('button');
    btn.textContent = stores[key].name;
    btn.addEventListener('click', () => {
      window.location.href = `sheets.html?store=${key}`;
    });
    list.appendChild(btn);
  });
});
