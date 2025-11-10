document.addEventListener('DOMContentLoaded', async () => {
  const list = document.getElementById('store-list');
  const status = document.getElementById('store-status');
  startLoading(status, '読込中・・・');

  initializeHelp('help/top.txt');

  try {
    await ensureSettingsLoaded();
  } catch (e) {
    stopLoading(status);
    if (list) {
      list.style.color = 'red';
      list.style.whiteSpace = 'pre-line';
      list.textContent = '店舗一覧の読み込みに失敗しました。\n通信環境をご確認のうえ、再度お試しください。';
    }
    return;
  }
  document.getElementById('version').textContent = `ver.${APP_VERSION}`;
  const stores = loadStores();
  stopLoading(status);
  if (list) {
    list.textContent = '';
    list.style.color = '';
    list.style.whiteSpace = '';
  }
  const err = document.getElementById('settings-error');
  if (window.settingsError && err) {
    err.textContent = '設定が読み込めませんでした。\nデフォルトの値を使用します。\n設定からエラーを確認してください。';
  }
  if (list) {
    Object.keys(stores).forEach(key => {
      const btn = document.createElement('button');
      btn.textContent = stores[key].name;
      btn.addEventListener('click', () => {
        window.location.href = `sheets.html?store=${key}`;
      });
      list.appendChild(btn);
    });
  }

  const infoBox = document.getElementById('announcements');
  if (infoBox) {
    infoBox.textContent = '';
    const header = document.createElement('div');
    header.style.textAlign = 'center';
    header.style.fontSize = '1.2rem';
    header.textContent = '●お知らせ●';
    infoBox.appendChild(header);

    const controlWrapper = document.createElement('div');
    infoBox.appendChild(controlWrapper);

    const messageDiv = document.createElement('div');
    messageDiv.textContent = 'お知らせの読み込み中';
    infoBox.appendChild(messageDiv);

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
        if (!notes.length) {
          controlWrapper.textContent = '';
          messageDiv.textContent = '現在お知らせはありません。';
          return;
        }
        controlWrapper.textContent = '';
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
        controlWrapper.appendChild(select);
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
        controlWrapper.textContent = '';
        messageDiv.textContent = 'お知らせを取得できませんでした。';
      });
  }
});
