document.addEventListener('DOMContentLoaded', async () => {
  const list = document.getElementById('store-list');
  const status = document.getElementById('store-status');
  const offlineControls = document.getElementById('offline-controls');
  const offlineButton = document.getElementById('offline-load-button');
  const offlineInfo = document.getElementById('offline-workbook-info');
  const LAST_STORE_STORAGE_KEY = 'lastSelectedStore';

  function getLastSelectedStoreKey(stores) {
    if (!stores) {
      return null;
    }
    try {
      const key = localStorage.getItem(LAST_STORE_STORAGE_KEY);
      if (key && stores[key]) {
        return key;
      }
    } catch (e) {
      // Ignore storage access issues.
    }
    return null;
  }

  function setLastSelectedStoreKey(key) {
    if (!key) {
      return;
    }
    try {
      localStorage.setItem(LAST_STORE_STORAGE_KEY, key);
    } catch (e) {
      // Ignore storage access issues.
    }
  }

  function updateOfflineIndicator(message, variant) {
    if (!offlineInfo) return;
    offlineInfo.textContent = message || '';
    offlineInfo.classList.remove('is-success', 'is-error');
    if (variant === 'success') {
      offlineInfo.classList.add('is-success');
    } else if (variant === 'error') {
      offlineInfo.classList.add('is-error');
    }
  }

  function resetOfflineState() {
    if (typeof clearOfflineWorkbook === 'function') {
      try {
        clearOfflineWorkbook();
      } catch (e) {
        // Ignore cleanup failures; the offline cache is best-effort only.
      }
    }
    updateOfflineIndicator('');
    if (status) {
      status.textContent = '';
    }
  }

  resetOfflineState();

  window.addEventListener('pageshow', event => {
    if (event.persisted) {
      resetOfflineState();
    }
  });

  let offlineFileInput = null;
  if (offlineControls && offlineButton) {
    offlineFileInput = document.createElement('input');
    offlineFileInput.type = 'file';
    offlineFileInput.accept = '.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    offlineFileInput.style.display = 'none';
    offlineControls.appendChild(offlineFileInput);

    offlineButton.addEventListener('click', () => {
      updateOfflineIndicator('');
      if (offlineFileInput) {
        offlineFileInput.value = '';
        offlineFileInput.click();
      }
    });

    offlineFileInput.addEventListener('change', () => {
      const file = offlineFileInput.files && offlineFileInput.files[0];
      if (!file) {
        return;
      }
      updateOfflineIndicator('');
      if (status) {
        startLoading(status, 'ローカルファイルを読み込み中・・・');
      }
      const reader = new FileReader();
      reader.onload = () => {
        stopLoading(status);
        try {
          if (typeof setOfflineWorkbook !== 'function') {
            throw new Error('Offline workbook is unavailable');
          }
          setOfflineWorkbook(reader.result, { fileName: file.name });
          updateOfflineIndicator('ローカルファイルを使用しています', 'success');
          const stores = typeof loadStores === 'function' ? loadStores() : null;
          const availableKeys = stores ? Object.keys(stores) : [];
          const storedKey = getLastSelectedStoreKey(stores);
          const targetKey = storedKey && availableKeys.includes(storedKey) ? storedKey : availableKeys[0];
          if (targetKey) {
            setLastSelectedStoreKey(targetKey);
            window.location.href = `sheets.html?store=${encodeURIComponent(targetKey)}&offline=1`;
          } else if (status) {
            status.textContent = '店舗情報が設定されていないためローカルファイルを開けません。設定を確認してください。';
          }
        } catch (e) {
          console.error('Failed to store offline workbook', e);
          updateOfflineIndicator('ローカルファイルを保存できませんでした。', 'error');
          if (status) {
            status.textContent = 'ローカルファイルを保存できませんでした。';
          }
        } finally {
          offlineFileInput.value = '';
        }
      };
      reader.onerror = () => {
        stopLoading(status);
        console.error('Failed to read offline workbook', reader.error);
        updateOfflineIndicator('ローカルファイルの読み込みに失敗しました。', 'error');
        if (status) {
          status.textContent = 'ローカルファイルの読み込みに失敗しました。';
        }
        offlineFileInput.value = '';
      };
      reader.readAsArrayBuffer(file);
    });
  }

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
  if (status && typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive()) {
    status.textContent = '店舗を選択して続行してください。';
  }
  if (list) {
    list.textContent = '';
    list.style.color = '';
    list.style.whiteSpace = '';
  }
  const err = document.getElementById('settings-error');
  if (window.settingsError && err) {
    err.textContent = '設定が読み込めませんでした。\nデフォルトの値を使用します。\n設定からエラーを確認してください。';
  }
  const storeKeys = Object.keys(stores);
  if (list && storeKeys.length > 0) {
    if (!document.getElementById('multi-store-overlay')) {
      const overlay = document.createElement('div');
      overlay.id = 'multi-store-overlay';
      overlay.style.display = 'none';

      const popup = document.createElement('div');
      popup.id = 'multi-store-popup';

      const title = document.createElement('h2');
      title.id = 'multi-store-title';
      title.textContent = '店舗横断計算モード';

      const description = document.createElement('p');
      description.id = 'multi-store-description';
      description.textContent = '計算したい店舗を選択してください。';

      const optionList = document.createElement('div');
      optionList.id = 'multi-store-list';

      const actions = document.createElement('div');
      actions.id = 'multi-store-actions';

      const closeBtn = document.createElement('button');
      closeBtn.type = 'button';
      closeBtn.id = 'multi-store-close';
      closeBtn.textContent = '閉じる';

      const startBtn = document.createElement('button');
      startBtn.type = 'button';
      startBtn.id = 'multi-store-start';
      startBtn.textContent = '読み込み開始';
      startBtn.disabled = true;

      actions.appendChild(closeBtn);
      actions.appendChild(startBtn);

      popup.appendChild(title);
      popup.appendChild(description);
      popup.appendChild(optionList);
      popup.appendChild(actions);

      overlay.appendChild(popup);
      document.body.appendChild(overlay);

      const selectedStores = new Set();

      function updateStartButton() {
        const count = selectedStores.size;
        startBtn.disabled = count === 0;
        startBtn.textContent = count === 0
          ? '読み込み開始'
          : `読み込み開始 (${count}件)`;
      }

      function toggleOverlay(show) {
        overlay.style.display = show ? 'flex' : 'none';
      }

      closeBtn.addEventListener('click', () => {
        toggleOverlay(false);
      });

      overlay.addEventListener('click', event => {
        if (event.target === overlay) {
          toggleOverlay(false);
        }
      });

      storeKeys.forEach(key => {
        const storeInfo = stores[key];
        if (!storeInfo) {
          return;
        }
        const optionBtn = document.createElement('button');
        optionBtn.type = 'button';
        optionBtn.className = 'multi-store-option';
        optionBtn.textContent = storeInfo.name;
        optionBtn.dataset.storeKey = key;
        optionBtn.addEventListener('click', () => {
          if (selectedStores.has(key)) {
            selectedStores.delete(key);
            optionBtn.classList.remove('is-selected');
          } else {
            selectedStores.add(key);
            optionBtn.classList.add('is-selected');
          }
          updateStartButton();
        });
        optionList.appendChild(optionBtn);
      });

      startBtn.addEventListener('click', () => {
        if (selectedStores.size === 0) {
          return;
        }
        const orderedKeys = Array.from(selectedStores);
        const params = new URLSearchParams();
        params.set('stores', orderedKeys.join(','));
        window.location.href = `sheets.html?${params.toString()}`;
      });

      const modeButton = document.createElement('button');
      modeButton.type = 'button';
      modeButton.id = 'multi-store-mode-button';
      modeButton.textContent = '店舗横断計算モード';
      modeButton.addEventListener('click', () => {
        toggleOverlay(true);
      });

      list.appendChild(modeButton);
      updateStartButton();
    }

    storeKeys.forEach(key => {
      const btn = document.createElement('button');
      btn.textContent = stores[key].name;
      btn.addEventListener('click', () => {
        setLastSelectedStoreKey(key);
        if (typeof setOfflineWorkbookActive === 'function') {
          try {
            setOfflineWorkbookActive(false);
          } catch (e) {
            // Ignore failures disabling offline mode.
          }
        }
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

  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.ready
      .then(registration => {
        if (!registration || !registration.active) {
          return;
        }
        registration.active.postMessage({
          type: 'WARMUP_CACHE',
          paths: [
            '/payroll.html',
            '/settings.html',
            '/sheets.html',
            '/payroll.js',
            '/settings.js',
            '/sheets.js',
            '/calc.js',
            '/help.js',
            '/help/payroll.txt',
            '/help/settings.txt',
            '/help/sheets.txt'
          ]
        });
      })
      .catch(() => {
        // Ignore failures warming the cache; navigation will still work without it.
      });
  }
});
