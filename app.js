const APP_VERSION = '1.6.2';
const SETTINGS_CACHE_KEY = 'remoteSettingsCache';
const VERSION_CHECK_URL = 'version.json';
const UPDATE_DISMISS_KEY = 'updateNoticeDismissedVersion';

(function setupUpdateChecker() {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    return;
  }

  async function clearCachesAndReload(button, statusLabel) {
    const originalText = button.textContent;
    button.disabled = true;
    button.textContent = '更新中…';
    if (statusLabel) {
      statusLabel.textContent = '更新の準備をしています…';
    }

    try {
      if (window.caches && typeof caches.keys === 'function') {
        const keys = await caches.keys();
        await Promise.all(keys.map(key => caches.delete(key)));
      }
      if ('serviceWorker' in navigator && typeof navigator.serviceWorker.getRegistrations === 'function') {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(
          registrations.map(reg =>
            reg.unregister().catch(() => {
              /* ignore */
            })
          )
        );
      }
    } catch (error) {
      console.error('Failed to prepare update', error);
      button.disabled = false;
      button.textContent = originalText;
      if (statusLabel) {
        statusLabel.textContent = '更新の準備中にエラーが発生しました。ページを再読み込みしてください。';
      }
      return;
    }

    if (statusLabel) {
      statusLabel.textContent = '最新バージョンを読み込み直しています…';
    }

    const url = new URL(window.location.href);
    url.searchParams.set('forceReload', Date.now().toString());
    window.location.replace(url.toString());
  }

  function showUpdateNotice(latestVersion) {
    if (document.getElementById('update-overlay')) {
      return;
    }

    const overlay = document.createElement('div');
    overlay.id = 'update-overlay';

    const popup = document.createElement('div');
    popup.id = 'update-popup';

    const heading = document.createElement('h2');
    heading.id = 'update-title';
    heading.textContent = '更新のお知らせ';
    popup.appendChild(heading);

    const message = document.createElement('p');
    message.id = 'update-message';
    message.textContent = '簡易給与計算ソフトの更新があります。\n下のボタンを押して更新してください。';
    popup.appendChild(message);

    const versionInfo = document.createElement('p');
    versionInfo.id = 'update-version';
    versionInfo.textContent = `現在: ver.${APP_VERSION}\n最新: ver.${latestVersion}`;
    popup.appendChild(versionInfo);

    const status = document.createElement('p');
    status.id = 'update-status';
    popup.appendChild(status);

    const button = document.createElement('button');
    button.id = 'update-confirm';
    button.textContent = '更新する';
    button.addEventListener('click', () => clearCachesAndReload(button, status));
    popup.appendChild(button);

    const continueButton = document.createElement('button');
    continueButton.id = 'update-continue';
    continueButton.textContent = 'このまま続ける';
    popup.appendChild(continueButton);

    const dismissOption = document.createElement('label');
    dismissOption.id = 'update-dismiss-option';
    dismissOption.style.display = 'none';

    const dismissCheckbox = document.createElement('input');
    dismissCheckbox.type = 'checkbox';
    dismissCheckbox.id = 'update-dismiss';
    dismissOption.appendChild(dismissCheckbox);

    const dismissText = document.createElement('span');
    dismissText.textContent = '再び表示しない';
    dismissOption.appendChild(dismissText);
    popup.appendChild(dismissOption);

    const closeButton = document.createElement('button');
    closeButton.id = 'update-close';
    closeButton.textContent = '閉じる';
    closeButton.style.display = 'none';
    popup.appendChild(closeButton);

    continueButton.addEventListener('click', () => {
      message.textContent = '更新しないと不具合の修正や機能の追加が反映されない場合があります。';
      versionInfo.style.display = 'none';
      status.style.display = 'none';
      button.style.display = 'none';
      continueButton.style.display = 'none';
      dismissOption.style.display = 'flex';
      closeButton.style.display = 'block';
    });

    closeButton.addEventListener('click', () => {
      if (dismissCheckbox.checked) {
        try {
          localStorage.setItem(UPDATE_DISMISS_KEY, latestVersion);
        } catch (error) {
          console.warn('Failed to remember update dismissal', error);
        }
      }
      overlay.remove();
    });

    overlay.appendChild(popup);
    document.body.appendChild(overlay);
    requestAnimationFrame(() => {
      overlay.style.display = 'flex';
    });
  }

  async function checkForUpdates() {
    try {
      const response = await fetch(`${VERSION_CHECK_URL}?t=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(`Unexpected status ${response.status}`);
      }
      const data = await response.json();
      const latestVersion = data && typeof data.version === 'string' ? data.version.trim() : '';
      if (!latestVersion || latestVersion === APP_VERSION) {
        return;
      }
      try {
        if (localStorage.getItem(UPDATE_DISMISS_KEY) === latestVersion) {
          return;
        }
      } catch (error) {
        console.warn('Failed to read update dismissal', error);
      }
      showUpdateNotice(latestVersion);
    } catch (error) {
      console.warn('Version check failed', error);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', checkForUpdates, { once: true });
  } else {
    checkForUpdates();
  }
})();

(function registerServiceWorker() {
  if (!('serviceWorker' in navigator)) {
    return;
  }
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('service-worker.js').catch(err => {
      console.error('Service worker registration failed:', err);
    });
  });
})();

let PASSWORD = '3963';
window.settingsError = false;
const SETTINGS_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vTKnnQY1d5BXnOstLwIhJOn7IX8aqHXC98XzreJoFscTUFPJXhef7jO2-0KKvZ7_fPF0uZwpbdcEpcV/pub?output=xlsx';
const SETTINGS_SHEET_NAME = '給与計算_設定';

// Simple password gate to restrict access
function initPasswordGate() {
  if (sessionStorage.getItem('pwAuth')) return;

  const overlay = document.createElement('div');
  overlay.id = 'pw-overlay';

  const container = document.createElement('div');
  container.className = 'pw-container';

  const message = document.createElement('div');
  message.id = 'pw-message';
  message.className = 'pw-message';
  message.textContent = 'パスワードを入力してください。';
  container.appendChild(message);

  const display = document.createElement('div');
  display.id = 'pw-display';
  display.className = 'pw-display';
  const slots = [];
  for (let i = 0; i < 4; i++) {
    const span = document.createElement('span');
    span.className = 'pw-slot';
    display.appendChild(span);
    slots.push(span);
  }
  container.appendChild(display);

  if (!settingsLoaded) {
    const loadingCover = document.createElement('div');
    loadingCover.className = 'pw-loading';
    loadingCover.textContent = 'パスワード問い合わせ中・・・';
    display.appendChild(loadingCover);
    settingsLoadPromise.finally(() => {
      loadingCover.remove();
    });
  }

  const keypad = document.createElement('div');
  keypad.className = 'pw-keypad';
  const keys = ['7','8','9','4','5','6','1','2','3','', '0','del'];
  keys.forEach(k => {
    if (k === '') {
      const placeholder = document.createElement('div');
      placeholder.className = 'pw-empty';
      keypad.appendChild(placeholder);
      return;
    }
    const btn = document.createElement('button');
    btn.dataset.key = k;
    btn.textContent = k === 'del' ? 'del' : k;
    keypad.appendChild(btn);
  });
  container.appendChild(keypad);

  const note = document.createElement('div');
  note.className = 'pw-note';
  note.innerHTML = 'デフォルトのパスワードはATMの売上金入金と同じです。<br>パスワードは設定用スプレッドシートから変更できます。';
  container.appendChild(note);

  overlay.appendChild(container);
  document.body.appendChild(overlay);

  let input = '';
  function updateDisplay() {
    slots.forEach((s, i) => {
      s.textContent = input[i] || '';
    });
  }
  function clearInput(msg) {
    input = '';
    updateDisplay();
    if (msg) {
      message.textContent = msg;
      setTimeout(() => {
        message.textContent = 'パスワードを入力してください。';
      }, 2000);
    }
  }
  function handleDigit(d) {
    if (input.length >= 4) return;
    input += d;
    updateDisplay();
    if (input.length === 4) {
      setTimeout(() => {
        if (input === PASSWORD) {
          sessionStorage.setItem('pwAuth', '1');
          window.removeEventListener('keydown', onKey);
          overlay.remove();
        } else {
          clearInput('パスワードが間違っています。');
        }
      }, 150);
    }
  }
  function delDigit() {
    input = input.slice(0, -1);
    updateDisplay();
  }
  function onKey(e) {
    if (e.key >= '0' && e.key <= '9') {
      handleDigit(e.key);
    } else if (e.key === 'Backspace' || e.key === 'Delete') {
      delDigit();
    }
  }
  window.addEventListener('keydown', onKey);

  keypad.addEventListener('click', e => {
    const btn = e.target.closest('button');
    if (!btn) return;
    const k = btn.dataset.key;
    if (k === 'del') {
      delDigit();
    } else {
      handleDigit(k);
    }
  });

  updateDisplay();
}

// Shared regex for time ranges such as "9:00-17:30"
const TIME_RANGE_REGEX = /^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/;

// Break deduction rules (minHours or more => deduct hours)
const BREAK_DEDUCTIONS = [
  { minHours: 8, deduct: 1 },
  { minHours: 7, deduct: 0.75 },
  { minHours: 6, deduct: 0.5 }
];

// Manage loading timers without mutating DOM elements
const loadingMap = new WeakMap();


let DEFAULT_STORES = {
  night: {
    name: '夜勤',
    url: 'https://docs.google.com/spreadsheets/d/1gCGyxiXXxOOhgHG2tk3BlzMpXuaWQULacySlIhhoWRY/edit?gid=601593061#gid=601593061',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  sagamihara_higashi: {
    name: '相模原東大沼店',
    url: 'https://docs.google.com/spreadsheets/d/1fEMEasqSGU30DuvCx6O6D0nJ5j6m6WrMkGTAaSQuqBY/edit?gid=358413717#gid=358413717',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  kobuchi: {
    name: '古淵駅前店',
    url: 'https://docs.google.com/spreadsheets/d/1hSD3sdIQftusWcNegZnGbCtJmByZhzpAvLJegDoJckQ/edit?gid=946573079#gid=946573079',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  hashimoto: {
    name: '相模原橋本五丁目店',
    url: 'https://docs.google.com/spreadsheets/d/1YYvWZaF9Li_RHDLevvOm2ND8ASJ3864uHRkDAiWBEDc/edit?gid=2000770170#gid=2000770170',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  },
  isehara: {
    name: '伊勢原高森七丁目店',
    url: 'https://docs.google.com/spreadsheets/d/1PfEQRnvHcKS5hJ6gkpJQc0VFjDoJUBhHl7JTTyJheZc/edit?gid=34390331#gid=34390331',
    baseWage: 1225,
    overtime: 1.25,
    excludeWords: ['月', '日', '曜日', '空き', '予定', '.']
  }
};

async function fetchRemoteSettings() {
  try {
    let res;
    const hasOutput = /[?&]output=/.test(SETTINGS_URL);

    if (hasOutput) {
      const direct = await fetch(SETTINGS_URL, { cache: 'no-store' });
      if (direct.ok) {
        res = direct;
      }
    }

    if (!res) {
      const exportUrl = toXlsxExportUrl(SETTINGS_URL);
      if (!exportUrl) {
        window.settingsError = true;
        return;
      }
      try {
        const converted = await fetch(exportUrl, { cache: 'no-store' });
        if (converted.ok) {
          res = converted;
        } else {
          window.settingsError = true;
          return;
        }
      } catch (e) {
        window.settingsError = true;
        return;
      }
    }


    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const sheet = wb.Sheets[SETTINGS_SHEET_NAME];
    if (!sheet) {
      window.settingsError = true;
      window.settingsErrorDetails = [`シート「${SETTINGS_SHEET_NAME}」が見つかりません`];
      return;
    }
    const rawStatus = sheet['B4']?.v;
    const status = rawStatus != null ? String(rawStatus).trim().toUpperCase() : null;
    if (status !== 'ALL_OK') {
      window.settingsError = true;
      const details = [];
      if (rawStatus != null && status !== 'ALL_OK') details.push(String(rawStatus));
      const cells = [
        { addr: 'B5', label: 'URL設定' },
        { addr: 'B6', label: '基本時給設定' },
        { addr: 'B7', label: '時間外倍率設定' },
        { addr: 'B8', label: 'パスワード' },
      ];
      cells.forEach(c => {
        const val = sheet[c.addr]?.v;
        if (val && String(val) !== 'OK') details.push(`${c.label}：${val}`);
      });
      window.settingsErrorDetails = details;
      return;
    }

    const baseWage = Number(sheet['D11']?.v);
    const overtime = Number(sheet['F11']?.v);
    const passwordCell = sheet['J11'];
    const password = passwordCell ? String(passwordCell.v) : null;
    const excludeWords = [];
    for (let r = 11, count = 0; ; r++) {
      const cell = sheet[`H${r}`];
      if (!cell || cell.v === undefined || cell.v === '') {
        if (count > 0) break;
        continue;
      }
      excludeWords.push(String(cell.v));
      count++;
    }
    const stores = {};
    let idx = 1;
    for (let r = 11; ; r++) {
      const nameCell = sheet[`A${r}`];
      const urlCell = sheet[`B${r}`];
      if ((!nameCell || nameCell.v === undefined || nameCell.v === '') && (!urlCell || urlCell.v === undefined || urlCell.v === '')) {
        if (idx > 1) break;
        continue;
      }
      if (nameCell && nameCell.v && urlCell && urlCell.v) {
        const key = `store${idx}`;
        stores[key] = { name: String(nameCell.v), url: String(urlCell.v), baseWage, overtime, excludeWords };
        idx++;
      }
    }
    if (Object.keys(stores).length) {
      window.settingsError = false;
      window.settingsErrorDetails = undefined;
      if (typeof document !== 'undefined') {
        const err = document.getElementById('settings-error');
        if (err) err.textContent = '';
      }
      DEFAULT_STORES = stores;
      if (password) PASSWORD = password;
    } else {
      window.settingsError = true;
    }
  } catch (e) {
    console.error('fetchRemoteSettings failed', e);
    window.settingsError = true;
  }
}

let settingsLoadPromise;
let settingsLoaded = false;
let cacheApplied = false;

function applySettingsRecord(record) {
  if (!record) return false;
  if (record.stores) DEFAULT_STORES = record.stores;
  if (record.password) PASSWORD = record.password;
  if (record.settingsError) window.settingsError = true;
  if (record.settingsErrorDetails) window.settingsErrorDetails = record.settingsErrorDetails;
  return true;
}

function loadSettingsFromCache() {
  let raw;
  try {
    raw = localStorage.getItem(SETTINGS_CACHE_KEY);
  } catch (e) {
    console.error('loadSettingsFromCache failed', e);
    return false;
  }
  if (!raw) return false;
  try {
    const data = JSON.parse(raw);
    return applySettingsRecord(data);
  } catch (e) {
    console.error('loadSettingsFromCache parse failed', e);
    try {
      localStorage.removeItem(SETTINGS_CACHE_KEY);
    } catch (removeErr) {
      console.error('loadSettingsFromCache cleanup failed', removeErr);
    }
    return false;
  }
}

function saveSettingsToCache() {
  const payload = {
    fetchedAt: Date.now(),
    stores: DEFAULT_STORES,
    password: PASSWORD,
    settingsError: window.settingsError || undefined,
    settingsErrorDetails: window.settingsErrorDetails || undefined,
  };
  try {
    localStorage.setItem(SETTINGS_CACHE_KEY, JSON.stringify(payload));
  } catch (e) {
    console.error('saveSettingsToCache failed', e);
  }
}

function shouldFetchRemoteSettings() {
  if (typeof navigator !== 'undefined' && 'onLine' in navigator) {
    return navigator.onLine !== false;
  }
  return true;
}

function ensureSettingsLoaded() {
  if (!cacheApplied) {
    const loaded = loadSettingsFromCache();
    if (loaded) {
      settingsLoaded = true;
    }
    cacheApplied = true;
  }

  if (!settingsLoadPromise) {
    const needsFetch = shouldFetchRemoteSettings();
    const promise = needsFetch
      ? fetchRemoteSettings().then(() => {
          saveSettingsToCache();
        })
      : Promise.resolve();

    settingsLoadPromise = promise.finally(() => {
      settingsLoadPromise = null;
    });
    if (typeof window !== 'undefined') {
      window.settingsLoadPromise = settingsLoadPromise;
    }
  }
  return settingsLoadPromise;
}

settingsLoadPromise = ensureSettingsLoaded();
settingsLoadPromise.then(() => {
  settingsLoaded = true;
});

if (typeof window !== 'undefined' && window.addEventListener) {
  window.addEventListener('online', () => {
    const promise = ensureSettingsLoaded();
    window.settingsLoadPromise = promise;
    promise.then(() => {
      settingsLoaded = true;
    });
  });
}
window.settingsLoadPromise = settingsLoadPromise;
document.addEventListener('DOMContentLoaded', initPasswordGate);

function loadStores() {
  let stored = {};
  const raw = localStorage.getItem('stores');
  if (raw) {
    try {
      stored = JSON.parse(raw);
    } catch (e) {
      console.error('loadStores parse failed', e);
      localStorage.removeItem('stores');
      alert('保存された店舗データが破損していたため、初期設定に戻しました。');
      stored = {};
    }
  }
  const merged = {};
  Object.keys(DEFAULT_STORES).forEach(key => {
    const base = DEFAULT_STORES[key];
    const custom = stored[key] || {};
    const customBaseWage = Number(custom.baseWage);
    const baseWage = Number.isFinite(customBaseWage) ? customBaseWage : base.baseWage;
    const customOvertime = Number(custom.overtime);
    const overtime = Number.isFinite(customOvertime) ? customOvertime : base.overtime;
    const excludeWords = Array.isArray(custom.excludeWords) ? custom.excludeWords : base.excludeWords;
    merged[key] = { ...base, ...custom, baseWage, overtime, excludeWords };
  });
  return merged;
}

function saveStores(stores) {
  try {
    localStorage.setItem('stores', JSON.stringify(stores));
  } catch (e) {
    console.error('saveStores failed', e);
    throw e;
  }
}

function getStore(key) {
  const stores = loadStores();
  return stores[key];
}

function updateStore(key, values) {
  const stores = loadStores();
  stores[key] = { ...stores[key], ...values };
  saveStores(stores);
}

function startLoading(el, text) {
  if (!el) return;
  stopLoading(el);
  const rawText = (text || '').replace(/[・.]+$/, '');
  const displayText = rawText === '読込中' ? '読み込み中' : rawText;
  el.textContent = '';

  const container = document.createElement('div');
  container.className = 'loading-container';

  const loader = document.createElement('div');
  loader.className = 'loader';
  container.appendChild(loader);

  const message = document.createElement('div');
  message.className = 'loading-message';
  message.textContent = displayText || '読み込み中';
  container.appendChild(message);

  el.appendChild(container);

  const timeout = setTimeout(() => {
    message.textContent = '通常よりも読み込みに時間がかかっていますが、正常に動作していますのでそのままお待ち下さい。';
    message.classList.remove('loading-message');
    message.classList.add('loading-note');
  }, 5000);

  loadingMap.set(el, { timeout });
}

function stopLoading(el) {
  if (!el) return;
  const timers = loadingMap.get(el);
  if (timers) {
    clearInterval(timers.interval);
    clearTimeout(timers.timeout);
    loadingMap.delete(el);
  }
  el.textContent = '';
}


function extractFileId(url) {
  const match = url.match(/\/d\/(?:e\/)?([a-zA-Z0-9_-]+)(?:\/|$)/);
  return match ? match[1] : null;
}

function toXlsxExportUrl(url) {
  const fileId = extractFileId(url);
  return fileId ? `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx` : null;
}

async function fetchWorkbook(url, sheetIndex = 0) {
  const exportUrl = toXlsxExportUrl(url);

  const res = await fetch(exportUrl, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }

  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  const targetIndex = (sheetIndex >= 0 && sheetIndex < wb.SheetNames.length) ? sheetIndex : 0;
  const sheetName = wb.SheetNames[targetIndex];
  const metaSheets = wb.Workbook && wb.Workbook.Sheets;
  const sheetId = metaSheets && metaSheets[targetIndex] ? metaSheets[targetIndex].sheetId : undefined;

  const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
  return { sheetName, data, sheetId, sheetIndex: targetIndex };
}

async function fetchSheetList(url) {
  const exportUrl = toXlsxExportUrl(url);
  const res = await fetch(exportUrl, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array', bookSheets: true });
  const metaSheets = wb.Workbook && wb.Workbook.Sheets;
  return wb.SheetNames.map((name, index) => ({
    name,
    index,
    sheetId: metaSheets && metaSheets[index] ? metaSheets[index].sheetId : undefined
  }));

}

function calculateEmployee(schedule, baseWage, overtime) {
  let total = 0;
  let workdays = 0;
  let salary = 0;
  schedule.forEach(cell => {
    if (!cell) return;
    const segments = cell.toString().split(',');
    let dayHours = 0;
    let hasValid = false;
    segments.forEach(seg => {
      const m = seg.trim().match(TIME_RANGE_REGEX);
      if (!m) return;
      const sh = parseInt(m[1], 10);
      const sm = m[2] ? parseInt(m[2], 10) : 0;
      const eh = parseInt(m[3], 10);
      const em = m[4] ? parseInt(m[4], 10) : 0;
      if (
        sh < 0 || sh > 24 || eh < 0 || eh > 24 ||
        sm < 0 || sm >= 60 || em < 0 || em >= 60
      ) return;
      hasValid = true;
      const start = sh * 60 + sm;
      const end = eh * 60 + em;
      const diff = end >= start ? end - start : 24 * 60 - start + end;
      dayHours += diff / 60;
    });
    if (!hasValid || dayHours <= 0) return;
    workdays++;
    for (const rule of BREAK_DEDUCTIONS) {
      if (dayHours >= rule.minHours) {
        dayHours -= rule.deduct;
        break;
      }
    }
    total += dayHours;
    const regular = Math.min(dayHours, 8);
    const over = Math.max(dayHours - 8, 0);
    salary += regular * baseWage + over * baseWage * overtime;
  });
  return { hours: total, days: workdays, salary: Math.floor(salary) };
}

function calculatePayroll(data, baseWage, overtime, excludeWords = []) {
  const header = data[2];
  const names = [];
  const schedules = [];
  for (let col = 3; col < header.length; col++) {
    const name = header[col];
    if (name && !excludeWords.some(word => name.includes(word))) {
      names.push(name);
      // rows 4-34 contain daily schedules
      schedules.push(data.slice(3, 34).map(row => row[col]));
    }
  }

  const results = names.map((name, idx) => {
    const r = calculateEmployee(schedules[idx], baseWage, overtime);
    return { name, baseWage, hours: r.hours, days: r.days, salary: r.salary };
  });

  const totalSalary = results.reduce((sum, r) => sum + r.salary, 0);
  return { results, totalSalary, schedules };
}

