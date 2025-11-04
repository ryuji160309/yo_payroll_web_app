document.addEventListener('DOMContentLoaded', async () => {
  await ensureSettingsLoaded();
  initializeHelp('help/payroll.txt');
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');

  const sheetIndex = parseInt(params.get('sheet'), 10) || 0;
  const store = getStore(storeKey);
  if (!store) return;
  const openSourceBtn = document.getElementById('open-source');
  if (openSourceBtn) {
    openSourceBtn.disabled = true;
  }
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  try {
    const key = `workbook_${storeKey}`;
    const cached = sessionStorage.getItem(key);
    let data;
    let sheetId;
    if (cached) {
      const buffer = base64ToBuffer(cached);
      const wb = XLSX.read(buffer, { type: 'array' });
      const targetIndex = (sheetIndex >= 0 && sheetIndex < wb.SheetNames.length) ? sheetIndex : 0;
      const sheetName = wb.SheetNames[targetIndex];
      const metaSheets = wb.Workbook && wb.Workbook.Sheets;
      sheetId = metaSheets && metaSheets[targetIndex] ? metaSheets[targetIndex].sheetId : undefined;
      data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
    } else {
      const result = await fetchWorkbook(store.url, sheetIndex, storeKey);
      data = result.data;
      sheetId = result.sheetId;
    }
    stopLoading(statusEl);
    const year = data[1] && data[1][2];
    const startMonthRaw = data[1] && data[1][4];
    const startMonth = parseInt(startMonthRaw, 10);
    const endMonth = (startMonth % 12) + 1;
    document.getElementById('period').textContent = `${year}年${startMonth}月16日～${endMonth}月15日`;
    const startDate = new Date(year, startMonth - 1, 16);
    const nameRow = data[36] || [];
    const storeName = nameRow[14] || store.name;
    document.getElementById('store-name').textContent = storeName;
    startLoading(statusEl, '計算中・・・');

    const { results, totalSalary, schedules } = calculatePayroll(data, store.baseWage, store.overtime, store.excludeWords || []);
    document.getElementById('total-salary').textContent = `合計支払い給与：${totalSalary.toLocaleString()}円`;
    const tbody = document.querySelector('#employees tbody');

    function recalc() {
      let total = 0;
      document.querySelectorAll('.wage-input').forEach(input => {
        const idx = parseInt(input.dataset.idx, 10);
        const wage = Number(input.value);
        if (Number.isFinite(wage)) {
          const calcResult = calculateEmployee(schedules[idx], wage, store.overtime);
          results[idx].baseWage = wage;
          results[idx].baseSalary = calcResult.salary;
        }
        const baseSalary = results[idx].baseSalary;
        const transportInput = document.querySelector(`.transport-input[data-idx="${idx}"]`);
        const transportRaw = transportInput ? Number(transportInput.value) : 0;
        const transport = Number.isFinite(transportRaw) ? transportRaw : 0;
        results[idx].transport = transport;
        const salary = baseSalary + transport;
        results[idx].salary = salary;
        const row = input.closest('tr');
        if (row) {
          const salaryCell = row.querySelector('.salary-cell');
          if (salaryCell) salaryCell.textContent = salary.toLocaleString();
        }
        total += salary;
      });
      document.getElementById('total-salary').textContent = `合計支払い給与：${total.toLocaleString()}円`;
    }

    results.forEach((r, idx) => {
      const tr = document.createElement('tr');

      const nameTd = document.createElement('td');
      nameTd.textContent = r.name;
      nameTd.addEventListener('click', () => {
        const lines = [];
        schedules[idx].forEach((cell, dayIdx) => {
          if (!cell) return;
          const segments = cell.toString().split(',')
            .map(s => s.trim())
            .filter(seg => TIME_RANGE_REGEX.test(seg));
          if (segments.length === 0) return;
          const d = new Date(startDate);
          d.setDate(startDate.getDate() + dayIdx);
          const mm = d.getMonth() + 1;
          const dd = d.getDate();
          lines.push(`${mm}/${dd} ${segments.join(', ')}`);
        });
        const message = lines.length ? `${r.name}\n${lines.join('\n')}` : `${r.name}\n出勤記録がありません`;
        alert(message);
      });

      const wageTd = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'number';
      input.value = r.baseWage;
      input.className = 'wage-input';
      input.dataset.idx = idx;
      input.addEventListener('input', recalc);
      wageTd.appendChild(input);

      const hoursTd = document.createElement('td');
      hoursTd.textContent = r.hours.toFixed(2);

      const daysTd = document.createElement('td');
      daysTd.textContent = r.days;

      const transportTd = document.createElement('td');
      const transportInput = document.createElement('input');
      transportInput.type = 'number';
      transportInput.value = r.transport;
      transportInput.className = 'transport-input';
      transportInput.dataset.idx = idx;
      transportInput.addEventListener('input', recalc);
      transportTd.appendChild(transportInput);

      const salaryTd = document.createElement('td');
      salaryTd.className = 'salary-cell';
      salaryTd.textContent = r.salary.toLocaleString();

      tr.appendChild(nameTd);
      tr.appendChild(wageTd);
      tr.appendChild(hoursTd);
      tr.appendChild(daysTd);
      tr.appendChild(transportTd);
      tr.appendChild(salaryTd);
      tbody.appendChild(tr);
    });
    recalc();
    stopLoading(statusEl);

    const baseWageInput = document.getElementById('base-wage-input');
    baseWageInput.value = store.baseWage;
    document.getElementById('set-base-wage').addEventListener('click', () => {
      const wage = Number(baseWageInput.value);
      document.querySelectorAll('.wage-input').forEach(input => {
        input.value = wage;
      });
      recalc();
    });

    const transportAllInput = document.getElementById('transport-input');
    transportAllInput.value = 0;
    document.getElementById('set-transport').addEventListener('click', () => {
      const transport = Number(transportAllInput.value);
      document.querySelectorAll('.transport-input').forEach(input => {
        input.value = transport;
      });
      recalc();
    });

    setupDownload(storeName, `${year}${startMonthRaw}`, results);
    setupSourceOpener(store.url, sheetId);
  } catch (e) {
    stopLoading(statusEl);
    document.getElementById('error').innerHTML = 'シートが読み込めませんでした。<br>シフト表ではないシートを選択しているか、表のデータが破損している可能性があります。';
  }
});

function setupDownload(storeName, period, results) {
  const button = document.getElementById('download');
  const overlay = document.createElement('div');
  overlay.id = 'download-overlay';
  const popup = document.createElement('div');
  popup.id = 'download-popup';
  const options = document.createElement('div');
  options.id = 'download-options';

  const txtBtn = document.createElement('button');
  txtBtn.textContent = 'テキスト形式';
  const xlsxBtn = document.createElement('button');
  xlsxBtn.textContent = 'EXCEL形式';
  const csvBtn = document.createElement('button');
  csvBtn.textContent = 'CSV形式';
  options.appendChild(txtBtn);
  options.appendChild(xlsxBtn);
  options.appendChild(csvBtn);

  const closeBtn = document.createElement('button');
  closeBtn.id = 'download-close';
  closeBtn.textContent = '閉じる';

  popup.appendChild(options);
  popup.appendChild(closeBtn);
  overlay.appendChild(popup);
  document.body.appendChild(overlay);

  function hide() {
    overlay.style.display = 'none';
  }

  button.addEventListener('click', () => {
    overlay.style.display = 'flex';
  });
  closeBtn.addEventListener('click', hide);
  overlay.addEventListener('click', e => { if (e.target === overlay) hide(); });

  txtBtn.addEventListener('click', () => { downloadResults(storeName, period, results, 'txt'); hide(); });
  xlsxBtn.addEventListener('click', () => { downloadResults(storeName, period, results, 'xlsx'); hide(); });
  csvBtn.addEventListener('click', () => { downloadResults(storeName, period, results, 'csv'); hide(); });
}

function setupSourceOpener(storeUrl, sheetId) {
  const button = document.getElementById('open-source');
  if (!button) return;
  if (!storeUrl) {
    button.disabled = true;
    button.title = '元シフトのURLが設定されていません。';
    return;
  }
  button.disabled = false;
  button.title = '';
  button.addEventListener('click', () => {
    const targetUrl = buildSheetUrl(storeUrl, sheetId);
    window.open(targetUrl, '_blank', 'noopener');
  });
}

function buildSheetUrl(baseUrl, sheetId) {
  if (sheetId === undefined || sheetId === null) return baseUrl;
  const gid = String(sheetId);
  try {
    const url = new URL(baseUrl);
    if (url.searchParams.has('gid')) {
      url.searchParams.set('gid', gid);
    }
    url.hash = `gid=${gid}`;
    return url.toString();
  } catch (e) {
    if (baseUrl.includes('#gid=')) {
      return baseUrl.replace(/#gid=\d+/, `#gid=${gid}`);
    }
    if (baseUrl.includes('gid=')) {
      return baseUrl.replace(/gid=\d+/, `gid=${gid}`);
    }
    return `${baseUrl}${baseUrl.includes('#') ? '' : '#'}gid=${gid}`;
  }
}

function downloadBlob(content, fileName, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

async function downloadResults(storeName, period, results, format) {
  const aoa = [
    ['従業員名', '基本時給', '勤務時間', '出勤日数', '交通費', '給与'],
    ...results.map(r => [r.name, r.baseWage, r.hours, r.days, r.transport, r.salary])
  ];
  const total = results.reduce((sum, r) => sum + r.salary, 0);
  aoa.push(['合計支払い給与', '', '', '', '', total]);

  if (format === 'csv') {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const csv = XLSX.utils.sheet_to_csv(ws);
    downloadBlob(csv, `${period}_${storeName}.csv`, 'text/csv');
  } else if (format === 'txt') {
    const text = aoa.map(row => row.join('\t')).join('\n');
    downloadBlob(text, `${period}_${storeName}.txt`, 'text/plain');
  } else {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '結果');
    XLSX.writeFile(wb, `${period}_${storeName}.xlsx`);
  }
}
