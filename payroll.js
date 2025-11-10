document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  initializeHelp('help/payroll.txt');
  await ensureSettingsLoaded();
  const params = new URLSearchParams(location.search);
  const offlineMode = params.get('offline') === '1';
  const offlineInfo = typeof getOfflineWorkbookInfo === 'function' ? getOfflineWorkbookInfo() : null;
  const offlineActive = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();
  const offlineIndicator = document.getElementById('offline-file-indicator');
  if (offlineIndicator) {
    offlineIndicator.classList.remove('is-success', 'is-error');
    if (offlineMode && offlineActive) {
      offlineIndicator.textContent = 'ローカルファイルを使用しています';
      offlineIndicator.classList.add('is-success');
    } else {
      offlineIndicator.textContent = '';
    }
  }
  if (offlineMode && !offlineActive) {
    stopLoading(statusEl);
    statusEl.textContent = 'ローカルファイルを利用できません。トップに戻って読み込み直してください。';
    return;
  }
  const storeKey = params.get('store');

  const sheetIndex = parseInt(params.get('sheet'), 10) || 0;
  const sheetGidParam = params.get('gid');
  const store = getStore(storeKey);
  if (!store) {
    stopLoading(statusEl);
    return;
  }
  if (!offlineMode && typeof setOfflineWorkbookActive === 'function') {
    try {
      setOfflineWorkbookActive(false);
    } catch (e) {
      // Ignore failures when disabling offline mode.
    }
  }
  const titleEl = document.getElementById('store-name');
  if (titleEl) {
    if (offlineMode && offlineActive && offlineInfo && offlineInfo.fileName) {
      titleEl.textContent = offlineInfo.fileName;
      document.title = `${offlineInfo.fileName} - 給与計算`;
    } else {
      titleEl.textContent = store.name;
      document.title = `${store.name} - 給与計算`;
    }
  }
  const openSourceBtn = document.getElementById('open-source');
  if (openSourceBtn) {
    openSourceBtn.disabled = true;
  }
  try {
    const result = await fetchWorkbook(store.url, sheetIndex, { allowOffline: offlineMode });
    const data = result.data;
    const sheetId = result.sheetId;
    stopLoading(statusEl);
    const year = data[1] && data[1][2];
    const startMonthRaw = data[1] && data[1][4];
    const startMonth = parseInt(startMonthRaw, 10);
    const endMonth = (startMonth % 12) + 1;
    document.getElementById('period').textContent = `${year}年${startMonth}月16日～${endMonth}月15日`;
    const startDate = new Date(year, startMonth - 1, 16);
    const nameRow = data[36] || [];
    const storeName = nameRow[14] || store.name;
    const displayTitleEl = document.getElementById('store-name');
    if (displayTitleEl) {
      const displayName = (offlineMode && offlineActive && offlineInfo && offlineInfo.fileName) ? offlineInfo.fileName : storeName;
      displayTitleEl.textContent = displayName;
      document.title = `${displayName} - 給与計算`;
    }
    startLoading(statusEl, '計算中・・・');

    const { results, totalSalary, schedules } = calculatePayroll(data, store.baseWage, store.overtime, store.excludeWords || []);
    document.getElementById('total-salary').textContent = `合計支払い給与：${totalSalary.toLocaleString()}円`;
    const tbody = document.querySelector('#employees tbody');

    const detailOverlay = document.createElement('div');
    detailOverlay.id = 'employee-detail-overlay';
    detailOverlay.style.display = 'none';
    const detailPopup = document.createElement('div');
    detailPopup.id = 'employee-detail-popup';
    const detailTitle = document.createElement('h2');
    detailTitle.id = 'employee-detail-title';
    const detailSummary = document.createElement('div');
    detailSummary.id = 'employee-detail-summary';
    const detailTableWrapper = document.createElement('div');
    detailTableWrapper.id = 'employee-detail-table-wrapper';
    const detailTable = document.createElement('table');
    detailTable.id = 'employee-detail-table';
    const detailTableBody = document.createElement('tbody');
    detailTable.appendChild(detailTableBody);
    detailTableWrapper.appendChild(detailTable);
    const detailClose = document.createElement('button');
    detailClose.id = 'employee-detail-close';
    detailClose.textContent = '閉じる';

    detailPopup.appendChild(detailTitle);
    detailPopup.appendChild(detailSummary);
    detailPopup.appendChild(detailTableWrapper);
    detailPopup.appendChild(detailClose);
    detailOverlay.appendChild(detailPopup);
    document.body.appendChild(detailOverlay);

    function hideEmployeeDetail() {
      detailOverlay.style.display = 'none';
    }

    function formatTimeSegment(match) {
      const startHour = match[1].padStart(2, '0');
      const startMinute = match[2];
      const endHour = match[3].padStart(2, '0');
      const endMinute = match[4];
      let startText = `${startHour}時`;
      if (startMinute !== undefined) {
        startText += `${startMinute}分`;
      }
      let endText = `${endHour}時`;
      if (endMinute !== undefined) {
        endText += `${endMinute}分`;
      }
      return `${startText}～${endText}`;
    }

    function showEmployeeDetail(idx) {
      const employee = results[idx];
      if (!employee) {
        return;
      }

      detailTitle.textContent = employee.name;
      const summaryLines = [
        `基本時給：${Number(employee.baseWage).toLocaleString()}円`,
        `総勤務時間：${employee.hours.toFixed(2)}時間`,
        `出勤日数：${employee.days}日`,
        `交通費：${Number(employee.transport || 0).toLocaleString()}円`,
        `給与：${Number(employee.salary || 0).toLocaleString()}円`
      ];
      detailSummary.innerHTML = '';
      summaryLines.forEach(line => {
        const lineEl = document.createElement('div');
        lineEl.className = 'employee-detail-summary-line';
        lineEl.textContent = line;
        detailSummary.appendChild(lineEl);
      });

      detailTableBody.innerHTML = '';

      const entries = [];
      const schedule = schedules[idx] || [];
      schedule.forEach((cell, dayIdx) => {
        if (!cell) return;
        const segments = cell.toString().split(',')
          .map(s => s.trim())
          .map(seg => {
            const match = seg.match(TIME_RANGE_REGEX);
            return match ? formatTimeSegment(match) : null;
          })
          .filter(Boolean);
        if (segments.length === 0) return;
        const current = new Date(startDate);
        current.setDate(startDate.getDate() + dayIdx);
        entries.push({
          month: current.getMonth() + 1,
          day: current.getDate(),
          segments,
        });
      });

      if (entries.length === 0) {
        const emptyRow = document.createElement('tr');
        const emptyCell = document.createElement('td');
        emptyCell.colSpan = 2;
        emptyCell.textContent = '出勤記録がありません';
        emptyRow.appendChild(emptyCell);
        detailTableBody.appendChild(emptyRow);
      } else {
        let currentMonth = null;
        entries.forEach(entry => {
          if (entry.month !== currentMonth) {
            currentMonth = entry.month;
            const monthRow = document.createElement('tr');
            monthRow.className = 'month-row';
            const monthCell = document.createElement('th');
            monthCell.colSpan = 2;
            monthCell.textContent = `${entry.month}月`;
            monthRow.appendChild(monthCell);
            detailTableBody.appendChild(monthRow);
          }

          const row = document.createElement('tr');
          const dateCell = document.createElement('td');
          dateCell.className = 'date-cell';
          dateCell.textContent = `${entry.day}日`;
          const timeCell = document.createElement('td');
          timeCell.className = 'time-cell';
          entry.segments.forEach((segment, segIdx) => {
            if (segIdx > 0) {
              timeCell.appendChild(document.createElement('br'));
            }
            timeCell.appendChild(document.createTextNode(segment));
          });
          row.appendChild(dateCell);
          row.appendChild(timeCell);
          detailTableBody.appendChild(row);
        });
      }

      detailOverlay.style.display = 'flex';
    }

    detailClose.addEventListener('click', hideEmployeeDetail);
    detailOverlay.addEventListener('click', e => {
      if (e.target === detailOverlay) {
        hideEmployeeDetail();
      }
    });
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape' && detailOverlay.style.display === 'flex') {
        hideEmployeeDetail();
      }
    });

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
      nameTd.addEventListener('click', () => showEmployeeDetail(idx));

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
        const idx = Number(input.dataset.idx);
        if (!Number.isFinite(idx)) return;
        const employee = results[idx];
        if (!employee || employee.days === 0) {
          return;
        }
        input.value = transport;
      });
      recalc();
    });

    setupDownload(storeName, `${year}${startMonthRaw}`, results);
    const preferredSheetId = sheetGidParam !== null ? sheetGidParam : sheetId;
    setupSourceOpener(store.url, preferredSheetId);
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
  if (typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive()) {
    button.disabled = true;
    button.title = 'ローカルファイルを使用しているため開けません。';
    return;
  }
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
