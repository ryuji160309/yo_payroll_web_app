const APP_VERSION = '1.3.1';

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
  note.textContent = 'パスワードはローソン銀行ATMの売上金入金と同じです。';
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
      if (input === '3963') {
        sessionStorage.setItem('pwAuth', '1');
        window.removeEventListener('keydown', onKey);
        overlay.remove();
      } else {
        clearInput('パスワードが間違っています。');
      }
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

document.addEventListener('DOMContentLoaded', initPasswordGate);

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


const DEFAULT_STORES = {
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
  const baseText = text.replace(/[・.]+$/, '');
  el.textContent = '';
  const mainSpan = document.createElement('span');
  mainSpan.textContent = baseText;
  const dotSpan = document.createElement('span');
  mainSpan.appendChild(dotSpan);
  el.appendChild(mainSpan);

  const note = document.createElement('div');
  note.className = 'loading-note';
  note.textContent = '通常よりも読み込みに時間がかかっていますが、正常に動作していますのでそのままお待ち下さい。';
  note.style.display = 'none';
  el.appendChild(note);

  let dotCount = 0;
  function updateDots() {
    dotSpan.textContent = '・'.repeat(dotCount);
    dotCount = (dotCount + 1) % 4;
  }
  updateDots();
  const interval = setInterval(updateDots, 500);
  const timeout = setTimeout(() => {
    note.style.display = 'block';
  }, 5000);
  loadingMap.set(el, { interval, timeout });
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
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)(?:\/|$)/);
  return match ? match[1] : null;
}

function toXlsxExportUrl(url) {
  const fileId = extractFileId(url);
  return fileId ? `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx` : null;
}

function bufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function base64ToBuffer(base64) {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

async function fetchWorkbook(url, sheetIndex = 0, storeKey) {
  const exportUrl = toXlsxExportUrl(url);

  const res = await fetch(exportUrl, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }

  const buffer = await res.arrayBuffer();
  if (storeKey) {
    sessionStorage.setItem(`workbook_${storeKey}`, bufferToBase64(buffer));
  }
  const wb = XLSX.read(buffer, { type: 'array' });
  const sheetName = wb.SheetNames[sheetIndex] || wb.SheetNames[0];

  const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
  return { sheetName, data };
}

async function fetchSheetList(url, storeKey) {
  const exportUrl = toXlsxExportUrl(url);
  const res = await fetch(exportUrl, { cache: 'no-store' });
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  const buffer = await res.arrayBuffer();
  if (storeKey) {
    sessionStorage.setItem(`workbook_${storeKey}`, bufferToBase64(buffer));
  }
  const wb = XLSX.read(buffer, { type: 'array', bookSheets: true });
  return wb.SheetNames.map((name, index) => ({ name, index }));

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

