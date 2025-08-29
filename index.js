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

  const infoBox = document.getElementById('announcements');
  if (infoBox) {
    fetch('announcements.txt', { cache: 'no-store' })
      .then(res => res.text())
      .then(text => {
        const blocks = text.trim().split(/\n\s*\n/);
        const notes = blocks.map(block => {
          const lines = block.trim().split('\n');
          const versionLine = lines.shift();
          const version = versionLine.replace(/^ver\./i, '').trim();
          const messages = lines.map(l => l.trim()).filter(Boolean);
          return { version, messages };
        }).sort((a, b) => {
          const pa = a.version.split('.').map(Number);
          const pb = b.version.split('.').map(Number);
          for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
            const diff = (pb[i] || 0) - (pa[i] || 0);
            if (diff !== 0) return diff;
          }
          return 0;
        });
        if (!notes.length) return;
        infoBox.innerHTML = '<div style="text-align:center;font-size:1.2rem;">●お知らせ●</div>';
        const select = document.createElement('select');
        const latestOpt = document.createElement('option');
        latestOpt.value = '';
        latestOpt.textContent = '最新3件';
        select.appendChild(latestOpt);
        notes.forEach(n => {
          const opt = document.createElement('option');
          opt.value = n.version;
          opt.textContent = `ver.${n.version}`;
          select.appendChild(opt);
        });
        infoBox.appendChild(select);
        const messageDiv = document.createElement('div');
        infoBox.appendChild(messageDiv);
        const latestNotes = notes.slice(0, 3);
        function render(ver) {
          if (!ver) {
            messageDiv.innerHTML = latestNotes.map(n => {
              return ['<strong>ver.' + n.version + '</strong>', ...n.messages].join('<br>');
            }).join('<br><br>');
            return;
          }
          const note = notes.find(n => n.version === ver);
          if (!note) return;
          messageDiv.innerHTML = note.messages.join('<br>');
        }
        select.addEventListener('change', () => render(select.value));
        select.value = '';
        render(select.value);
      })
      .catch(() => {
        infoBox.textContent = 'お知らせを取得できませんでした。';
      });
  }
});
