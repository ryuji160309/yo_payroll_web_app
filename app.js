const APP_VERSION = '0.1.0';

const DEFAULT_STORES = {
  night: {
    name: '夜勤',
    url: 'https://docs.google.com/spreadsheets/d/1gCGyxiXXxOOhgHG2tk3BlzMpXuaWQULacySlIhhoWRY/edit?gid=601593061#gid=601593061',
    baseWage: 1000,
    overtime: 1.25
  },
  sagamihara_higashi: {
    name: '相模原東大沼店',
    url: 'https://docs.google.com/spreadsheets/d/1fEMEasqSGU30DuvCx6O6D0nJ5j6m6WrMkGTAaSQuqBY/edit?gid=358413717#gid=358413717',
    baseWage: 1000,
    overtime: 1.25
  },
  kobuchi: {
    name: '古淵駅前店',
    url: 'https://docs.google.com/spreadsheets/d/1hSD3sdIQftusWcNegZnGbCtJmByZhzpAvLJegDoJckQ/edit?gid=946573079#gid=946573079',
    baseWage: 1000,
    overtime: 1.25
  },
  hashimoto: {
    name: '相模原橋本五丁目店',
    url: 'https://docs.google.com/spreadsheets/d/1YYvWZaF9Li_RHDLevvOm2ND8ASJ3864uHRkDAiWBEDc/edit?gid=2000770170#gid=2000770170',
    baseWage: 1000,
    overtime: 1.25
  },
  isehara: {
    name: '伊勢原高森七丁目店',
    url: 'https://docs.google.com/spreadsheets/d/1PfEQRnvHcKS5hJ6gkpJQc0VFjDoJUBhHl7JTTyJheZc/edit?gid=34390331#gid=34390331',
    baseWage: 1000,
    overtime: 1.25
  }
};

function loadStores() {
  const stored = JSON.parse(localStorage.getItem('stores') || '{}');
  return { ...DEFAULT_STORES, ...stored };
}

function saveStores(stores) {
  localStorage.setItem('stores', JSON.stringify(stores));
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

function calculatePayroll(data, baseWage, overtime) {
  const header = data[2];
  const names = [];
  const schedules = [];
  for (let col = 3; col < header.length; col++) {
    const name = header[col];
    if (name && !['月', '日', '曜日', '空き', '予定', '.'].includes(name)) {
      names.push(name);
      schedules.push(data.slice(3).map(row => row[col]));
    }
  }

  const results = names.map((name, idx) => {
    let total = 0;
    let workdays = 0;
    let salary = 0;
    schedules[idx].forEach(cell => {
      if (!cell) return;
      workdays++;
      const segments = cell.toString().split(',');
      let dayHours = 0;
      segments.forEach(seg => {
        const [s, e] = seg.split('-').map(Number);
        if (!isNaN(s) && !isNaN(e)) {
          let h = e >= s ? e - s : 24 - s + e;
          dayHours += h;
        }
      });
      if (dayHours >= 8) dayHours -= 1;
      else if (dayHours >= 7) dayHours -= 0.75;
      else if (dayHours >= 6) dayHours -= 0.5;
      total += dayHours;
      const regular = Math.min(dayHours, 8);
      const over = Math.max(dayHours - 8, 0);
      salary += regular * baseWage + over * baseWage * overtime;
    });
    return {
      name,
      hours: total,
      days: workdays,
      salary: Math.floor(salary)
    };
  });

  const totalSalary = results.reduce((sum, r) => sum + r.salary, 0);
  return { results, totalSalary };
}

