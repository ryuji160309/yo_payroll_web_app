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

  // Populate announcements from ANNOUNCEMENTS defined in app.js
  const infoBox = document.getElementById('announcements');
  if (infoBox && Array.isArray(ANNOUNCEMENTS)) {
    let html = '<div style="text-align:center;font-size:1.2rem;">●お知らせ●</div><div>';
    ANNOUNCEMENTS.forEach((note, idx) => {
      html += `ver.${note.version}<br>${note.messages.join('<br>')}`;
      if (idx !== ANNOUNCEMENTS.length - 1) {
        html += '<br><br>';
      }
    });
    html += '</div>';
    infoBox.innerHTML = html;
  }
});
