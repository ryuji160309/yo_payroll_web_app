const APP_VERSION = '1.1.1';


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
  try {
    stored = JSON.parse(localStorage.getItem('stores') || '{}');
  } catch (e) {
    stored = {};
  }
  const merged = {};
  Object.keys(DEFAULT_STORES).forEach(key => {
    const base = DEFAULT_STORES[key];
    const custom = stored[key] || {};
    const baseWage = Number.isFinite(custom.baseWage) ? custom.baseWage : base.baseWage;
    const overtime = Number.isFinite(custom.overtime) ? custom.overtime : base.overtime;
    const excludeWords = Array.isArray(custom.excludeWords) ? custom.excludeWords : base.excludeWords;
    merged[key] = { ...base, ...custom, baseWage, overtime, excludeWords };
  });
  return merged;
}

function saveStores(stores) {
  try {
    localStorage.setItem('stores', JSON.stringify(stores));
    return true;
  } catch (e) {
    console.error('saveStores failed', e);
    return false;
  }
}

function getStore(key) {
  const stores = loadStores();
  return stores[key];
}

function updateStore(key, values) {
  const stores = loadStores();
  stores[key] = { ...stores[key], ...values };
  if (!saveStores(stores)) {
    throw new Error('Failed to save stores');
  }
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
    dotCount = (dotCount % 3) + 1;
    dotSpan.textContent = '・'.repeat(dotCount);
  }
  updateDots();
  const interval = setInterval(updateDots, 500);
  const timeout = setTimeout(() => {
    note.style.display = 'block';
  }, 5000);
  el._loadingInterval = interval;
  el._loadingTimeout = timeout;
}

function stopLoading(el) {
  if (!el) return;
  clearInterval(el._loadingInterval);
  clearTimeout(el._loadingTimeout);
  el.textContent = '';
  delete el._loadingInterval;
  delete el._loadingTimeout;
}


function extractFileId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)(?:\/|$)/);
  return match ? match[1] : null;
}

function toXlsxExportUrl(url) {
  const fileId = extractFileId(url);
  return fileId ? `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx` : null;
}

async function fetchWorkbook(url, sheetIndex = 0) {
  const exportUrl = toXlsxExportUrl(url);

  const res = await fetch(exportUrl);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }

  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
  const sheetName = wb.SheetNames[sheetIndex] || wb.SheetNames[0];

  const data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
  return { sheetName, data };
}

async function fetchSheetList(url) {
  const exportUrl = toXlsxExportUrl(url);
  const res = await fetch(exportUrl);
  if (!res.ok) {
    throw new Error(`HTTP ${res.status}`);
  }
  const buffer = await res.arrayBuffer();
  const wb = XLSX.read(buffer, { type: 'array' });
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
      const m = seg.trim().match(/^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/);
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
    if (dayHours >= 8) dayHours -= 1;
    else if (dayHours >= 7) dayHours -= 0.75;
    else if (dayHours >= 6) dayHours -= 0.5;
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

