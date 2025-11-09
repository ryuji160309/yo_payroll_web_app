document.addEventListener('DOMContentLoaded', async () => {
  const list = document.getElementById('store-list');
  startLoading(list, '読込中・・・');

  try {
    await ensureSettingsLoaded();
  } catch (e) {
    stopLoading(list);
    if (list) {
      list.style.color = 'red';
      list.style.whiteSpace = 'pre-line';
      list.textContent = '店舗一覧の読み込みに失敗しました。\n通信環境をご確認のうえ、再度お試しください。';
    }
    return;
  }
  document.getElementById('version').textContent = `ver.${APP_VERSION}`;
  const stores = loadStores();
  stopLoading(list);
  const err = document.getElementById('settings-error');
  if (window.settingsError && err) {
    err.textContent = '設定が読み込めませんでした。\nデフォルトの値を使用します。\n設定からエラーを確認してください。';
  }
  initializeHelp('help/top.txt');
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
        infoBox.textContent = '';
        const header = document.createElement('div');
        header.style.textAlign = 'center';
        header.style.fontSize = '1.2rem';
        header.textContent = '●お知らせ●';
        infoBox.appendChild(header);
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

        function buildNoteFragment(note) {
          const frag = document.createDocumentFragment();
          const strong = document.createElement('strong');
          strong.textContent = `ver.${note.version}`;
          frag.appendChild(strong);
          note.messages.forEach(msg => {
            frag.appendChild(document.createElement('br'));
            const span = document.createElement('span');
            span.textContent = msg;
            frag.appendChild(span);
          });
          return frag;
        }

        function render(ver) {
          messageDiv.textContent = '';
          if (!ver) {
            latestNotes.forEach((n, idx) => {
              messageDiv.appendChild(buildNoteFragment(n));
              if (idx < latestNotes.length - 1) {
                messageDiv.appendChild(document.createElement('br'));
                messageDiv.appendChild(document.createElement('br'));
              }
            });
            return;
          }
          const note = notes.find(n => n.version === ver);
          if (!note) return;
          messageDiv.appendChild(buildNoteFragment(note));
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
