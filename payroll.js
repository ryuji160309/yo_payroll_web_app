document.addEventListener('DOMContentLoaded', async () => {
  await ensureSettingsLoaded();
  initializeHelp('help/payroll.txt');
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');

  const sheetIndex = parseInt(params.get('sheet'), 10) || 0;
  const store = getStore(storeKey);
  if (!store) return;
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  try {
    const key = `workbook_${storeKey}`;
    const cached = sessionStorage.getItem(key);
    let data;
    if (cached) {
      const buffer = base64ToBuffer(cached);
      const wb = XLSX.read(buffer, { type: 'array' });
      const sheetName = wb.SheetNames[sheetIndex] || wb.SheetNames[0];
      data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false });
    } else {
      const result = await fetchWorkbook(store.url, sheetIndex, storeKey);
      data = result.data;
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
    const totalSalaryEl = document.getElementById('total-salary');
    const transportationToggle = document.getElementById('transportation-toggle');
    const transportationBulk = document.getElementById('transportation-bulk');
    const transportationInput = document.getElementById('transportation-input');
    const setTransportationBtn = document.getElementById('set-transportation');
    transportationInput.value = 0;
    const transportationEnabled = () => transportationToggle && transportationToggle.checked;

    function getActiveTransportation(amount) {
      return transportationEnabled() ? amount : 0;
    }

    function updateSalaryCell(cell, idx) {
      const transport = getActiveTransportation(results[idx].transportation || 0);
      cell.textContent = (results[idx].salary + transport).toLocaleString();
    }

    function updateTotals() {
      const includeTransportation = transportationEnabled();
      const total = results.reduce((sum, r) => sum + r.salary + (includeTransportation ? (r.transportation || 0) : 0), 0);
      totalSalaryEl.textContent = `合計支払い給与：${total.toLocaleString()}円`;
    }

    totalSalaryEl.textContent = `合計支払い給与：${totalSalary.toLocaleString()}円`;
    const tbody = document.querySelector('#employees tbody');
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
      wageTd.appendChild(input);

      const hoursTd = document.createElement('td');
      hoursTd.textContent = r.hours.toFixed(2);

      const daysTd = document.createElement('td');
      daysTd.textContent = r.days;

      const salaryTd = document.createElement('td');
      salaryTd.className = 'salary-cell';
      salaryTd.dataset.idx = idx;
      updateSalaryCell(salaryTd, idx);

      const transportationTd = document.createElement('td');
      const transportationField = document.createElement('input');
      transportationField.type = 'number';
      transportationField.value = r.transportation || 0;
      transportationField.className = 'transportation-input';
      transportationField.dataset.idx = idx;
      transportationField.disabled = !transportationEnabled();
      transportationField.addEventListener('input', () => {
        const value = Number(transportationField.value);
        results[idx].transportation = Number.isFinite(value) ? value : 0;
        updateSalaryCell(salaryTd, idx);
        updateTotals();
      });
      transportationTd.appendChild(transportationField);

      tr.appendChild(nameTd);
      tr.appendChild(wageTd);
      tr.appendChild(hoursTd);
      tr.appendChild(daysTd);
      tr.appendChild(transportationTd);
      tr.appendChild(salaryTd);
      tbody.appendChild(tr);
    });
    stopLoading(statusEl);

    const baseWageInput = document.getElementById('base-wage-input');
    baseWageInput.value = store.baseWage;
    document.getElementById('set-base-wage').addEventListener('click', () => {
      const wage = Number(baseWageInput.value);
      document.querySelectorAll('.wage-input').forEach(input => {
        input.value = wage;
      });
      document.getElementById('recalc').click();
    });

    document.getElementById('recalc').addEventListener('click', () => {
      const inputs = document.querySelectorAll('.wage-input');
      let total = 0;
      const includeTransportation = transportationEnabled();
      inputs.forEach(input => {
        const idx = parseInt(input.dataset.idx, 10);
        const wage = Number(input.value);
        if (!Number.isFinite(wage)) {
          const transport = includeTransportation ? (results[idx].transportation || 0) : 0;
          total += results[idx].salary + transport;
          return;
        }
        const r = calculateEmployee(schedules[idx], wage, store.overtime);
        results[idx].baseWage = wage;
        results[idx].salary = r.salary;
        const salaryCell = input.closest('tr').querySelector('.salary-cell');
        const transport = includeTransportation ? (results[idx].transportation || 0) : 0;
        total += r.salary + transport;
        salaryCell.textContent = (r.salary + transport).toLocaleString();
      });
      totalSalaryEl.textContent = `合計支払い給与：${total.toLocaleString()}円`;
    });

    if (transportationToggle) {
      const setEnabledState = enabled => {
        transportationBulk.style.display = enabled ? 'inline-block' : 'none';
        transportationInput.disabled = !enabled;
        setTransportationBtn.disabled = !enabled;
        document.querySelectorAll('.transportation-input').forEach(input => {
          input.disabled = !enabled;
        });
        document.querySelectorAll('.salary-cell').forEach(cell => {
          const idx = parseInt(cell.dataset.idx, 10);
          updateSalaryCell(cell, idx);
        });
        updateTotals();
      };

      setEnabledState(transportationEnabled());

      transportationToggle.addEventListener('change', () => {
        setEnabledState(transportationToggle.checked);
      });

      setTransportationBtn.addEventListener('click', () => {
        const amount = Number(transportationInput.value);
        const value = Number.isFinite(amount) ? amount : 0;
        document.querySelectorAll('.transportation-input').forEach(input => {
          input.value = value;
          const idx = parseInt(input.dataset.idx, 10);
          results[idx].transportation = value;
          const salaryCell = input.closest('tr').querySelector('.salary-cell');
          updateSalaryCell(salaryCell, idx);
        });
        updateTotals();
      });
    }

    updateTotals();

    setupDownload(storeName, `${year}${startMonthRaw}`, results, transportationEnabled);
  } catch (e) {
    stopLoading(statusEl);
    document.getElementById('error').innerHTML = 'シートが読み込めませんでした。<br>シフト表ではないシートを選択しているか、表のデータが破損している可能性があります。';
  }
});

function setupDownload(storeName, period, results, transportationEnabledFn = () => true) {
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

  txtBtn.addEventListener('click', () => { downloadResults(storeName, period, results, 'txt', transportationEnabledFn()); hide(); });
  xlsxBtn.addEventListener('click', () => { downloadResults(storeName, period, results, 'xlsx', transportationEnabledFn()); hide(); });
  csvBtn.addEventListener('click', () => { downloadResults(storeName, period, results, 'csv', transportationEnabledFn()); hide(); });
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

async function downloadResults(storeName, period, results, format, includeTransportation = true) {
  const aoa = [['従業員名', '基本時給', '勤務時間', '出勤日数', '交通費', '給与'], ...results.map(r => {
    const transportation = includeTransportation ? (r.transportation || 0) : 0;
    return [r.name, r.baseWage, r.hours, r.days, transportation, r.salary + transportation];
  })];
  const total = results.reduce((sum, r) => {
    const transportation = includeTransportation ? (r.transportation || 0) : 0;
    return sum + r.salary + transportation;
  }, 0);
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
