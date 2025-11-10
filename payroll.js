document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  initializeHelp('help/payroll.txt');
  await ensureSettingsLoaded();
  const params = new URLSearchParams(location.search);
  const offlineMode = params.get('offline') === '1';
  const offlineInfo = typeof getOfflineWorkbookInfo === 'function' ? getOfflineWorkbookInfo() : null;
  const offlineActive = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();
  if (offlineMode && !offlineActive) {
    stopLoading(statusEl);
    statusEl.textContent = 'ローカルファイルを利用できません。トップに戻って読み込み直してください。';
    return;
  }
  const storeKey = params.get('store');

  const multiSheetsParam = params.get('sheets');
  let sheetIndices = multiSheetsParam
    ? multiSheetsParam.split(',').map(v => parseInt(v, 10))
    : [];
  const seenIndices = new Set();
  sheetIndices = sheetIndices.filter(idx => {
    if (!Number.isFinite(idx)) return false;
    if (seenIndices.has(idx)) return false;
    seenIndices.add(idx);
    return true;
  });
  if (sheetIndices.length === 0) {
    const fallbackSheet = parseInt(params.get('sheet'), 10);
    sheetIndices = [Number.isFinite(fallbackSheet) ? fallbackSheet : 0];
  }
  const isMultiSheetMode = sheetIndices.length > 1;
  const sheetGidParam = params.get('gid');
  const multiGidsParam = params.get('gids');
  const requestedGids = multiGidsParam
    ? multiGidsParam.split(',')
    : (sheetGidParam !== null ? [sheetGidParam] : []);
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
    const workbookResults = await Promise.all(
      sheetIndices.map(idx => fetchWorkbook(store.url, idx, { allowOffline: offlineMode }))
    );
    stopLoading(statusEl);

    const summaries = workbookResults.map(result => {
      const data = result.data;
      const rawYear = data[1] && data[1][2];
      const year = Number.parseInt(rawYear, 10);
      const startMonthRaw = data[1] && data[1][4];
      const startMonth = Number.parseInt(startMonthRaw, 10);
      const hasValidStart = Number.isFinite(year) && Number.isFinite(startMonth);
      const startDate = hasValidStart ? new Date(year, startMonth - 1, 16) : null;
      const endDate = startDate ? new Date(startDate.getFullYear(), startDate.getMonth() + 1, 15) : null;
      let periodLabel = '';
      if (startDate && endDate) {
        const endYear = endDate.getFullYear();
        const endMonth = endDate.getMonth() + 1;
        periodLabel = endYear !== year
          ? `${year}年${startMonth}月16日～${endYear}年${endMonth}月15日`
          : `${year}年${startMonth}月16日～${endMonth}月15日`;
      } else {
        periodLabel = result.sheetName || `シート${result.sheetIndex + 1}`;
      }
      const payrollResult = calculatePayroll(data, store.baseWage, store.overtime, store.excludeWords || []);
      const nameRow = data[36] || [];
      const sheetStoreName = nameRow[14];
      return {
        result,
        data,
        payrollResult,
        startDate,
        endDate,
        periodLabel,
        startMonthRaw: startMonthRaw !== undefined && startMonthRaw !== null ? String(startMonthRaw) : '',
        year: Number.isFinite(year) ? year : null,
        sheetStoreName,
        sheetId: result.sheetId,
      };
    });

    const resolvedStoreName = summaries.map(s => s.sheetStoreName).find(name => !!name) || store.name;
    const displayTitleEl = document.getElementById('store-name');
    if (displayTitleEl) {
      const displayName = (offlineMode && offlineActive && offlineInfo && offlineInfo.fileName)
        ? offlineInfo.fileName
        : resolvedStoreName;
      displayTitleEl.textContent = displayName;
      document.title = `${displayName} - 給与計算`;
    }

    const periodEl = document.getElementById('period');
    if (periodEl) {
      if (isMultiSheetMode) {
        const labels = summaries.map(s => s.periodLabel).filter(Boolean);
        periodEl.textContent = labels.length ? `選択した月：${labels.join(' ／ ')}` : '選択した月：複数シート';
      } else {
        periodEl.textContent = summaries[0] ? summaries[0].periodLabel : '';
      }
    }

    startLoading(statusEl, '計算中・・・');

    const resultMap = new Map();
    summaries.forEach(summary => {
      const schedulesList = summary.payrollResult.schedules || [];
      summary.payrollResult.results.forEach((employee, idx) => {
        const baseSalary = Number(employee.baseSalary ?? employee.salary ?? 0);
        const hours = Number(employee.hours) || 0;
        const days = Number(employee.days) || 0;
        const scheduleSource = Array.isArray(schedulesList[idx]) ? schedulesList[idx].slice() : [];
        const blockStartDate = summary.startDate ? new Date(summary.startDate.getTime()) : null;
        const existing = resultMap.get(employee.name);
        if (!existing) {
          resultMap.set(employee.name, {
            name: employee.name,
            baseWage: employee.baseWage,
            hours,
            days,
            baseSalary,
            transport: Number(employee.transport || 0),
            salary: baseSalary + Number(employee.transport || 0),
            flattenedSchedule: scheduleSource.slice(),
            scheduleBlocks: blockStartDate && scheduleSource.length > 0
              ? [{ startDate: blockStartDate, schedule: scheduleSource.slice() }]
              : [],
          });
        } else {
          existing.baseWage = employee.baseWage;
          existing.hours += hours;
          existing.days += days;
          existing.baseSalary += baseSalary;
          existing.salary = existing.baseSalary + Number(existing.transport || 0);
          if (scheduleSource.length > 0) {
            if (!Array.isArray(existing.flattenedSchedule)) {
              existing.flattenedSchedule = [];
            }
            existing.flattenedSchedule.push(...scheduleSource);
          }
          if (blockStartDate && scheduleSource.length > 0) {
            if (!Array.isArray(existing.scheduleBlocks)) {
              existing.scheduleBlocks = [];
            }
            existing.scheduleBlocks.push({ startDate: blockStartDate, schedule: scheduleSource.slice() });
          }
        }
      });
    });

    const results = Array.from(resultMap.values());
    results.forEach(employee => {
      if (!Array.isArray(employee.flattenedSchedule)) {
        employee.flattenedSchedule = [];
      }
      if (!Array.isArray(employee.scheduleBlocks)) {
        employee.scheduleBlocks = [];
      }
      employee.transport = Number(employee.transport || 0);
      employee.salary = employee.baseSalary + employee.transport;
    });

    const totalSalary = results.reduce((sum, r) => sum + (Number(r.salary) || 0), 0);
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
      (employee.scheduleBlocks || []).forEach(block => {
        if (!block || !block.schedule || block.schedule.length === 0) return;
        const baseDate = block.startDate ? new Date(block.startDate) : null;
        if (!baseDate || Number.isNaN(baseDate.getTime())) return;
        block.schedule.forEach((cell, dayIdx) => {
          if (!cell) return;
          const segments = cell.toString().split(',')
            .map(s => s.trim())
            .map(seg => {
              const match = seg.match(TIME_RANGE_REGEX);
              return match ? formatTimeSegment(match) : null;
            })
            .filter(Boolean);
          if (segments.length === 0) return;
          const current = new Date(baseDate);
          current.setDate(baseDate.getDate() + dayIdx);
          entries.push({
            date: current,
            segments,
          });
        });
      });

      entries.sort((a, b) => a.date - b.date);

      if (entries.length === 0) {
        const emptyRow = document.createElement('tr');
        const emptyCell = document.createElement('td');
        emptyCell.colSpan = 2;
        emptyCell.textContent = '出勤記録がありません';
        emptyRow.appendChild(emptyCell);
        detailTableBody.appendChild(emptyRow);
      } else {
        let currentMonthKey = null;
        entries.forEach(entry => {
          const monthKey = `${entry.date.getFullYear()}-${entry.date.getMonth()}`;
          if (monthKey !== currentMonthKey) {
            currentMonthKey = monthKey;
            const monthRow = document.createElement('tr');
            monthRow.className = 'month-row';
            const monthCell = document.createElement('th');
            monthCell.colSpan = 2;
            monthCell.textContent = `${entry.date.getFullYear()}年${entry.date.getMonth() + 1}月`;
            monthRow.appendChild(monthCell);
            detailTableBody.appendChild(monthRow);
          }

          const row = document.createElement('tr');
          const dateCell = document.createElement('td');
          dateCell.className = 'date-cell';
          dateCell.textContent = `${entry.date.getDate()}日`;
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
          const employee = results[idx];
          const schedule = employee && Array.isArray(employee.flattenedSchedule) ? employee.flattenedSchedule : [];
          const calcResult = calculateEmployee(schedule, wage, store.overtime);
          if (employee) {
            employee.baseWage = wage;
            employee.baseSalary = calcResult.salary;
          }
        }
        const baseSalary = results[idx] ? results[idx].baseSalary : 0;
        const transportInput = document.querySelector(`.transport-input[data-idx="${idx}"]`);
        const transportRaw = transportInput ? Number(transportInput.value) : 0;
        const transport = Number.isFinite(transportRaw) ? transportRaw : 0;
        if (results[idx]) {
          results[idx].transport = transport;
        }
        const salary = baseSalary + transport;
        if (results[idx]) {
          results[idx].salary = salary;
        }
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

    const sortedSummaries = summaries.slice().sort((a, b) => {
      if (a.startDate && b.startDate) return a.startDate - b.startDate;
      if (a.startDate) return -1;
      if (b.startDate) return 1;
      return (a.result.sheetIndex || 0) - (b.result.sheetIndex || 0);
    });
    let downloadPeriodId = 'result';
    if (isMultiSheetMode) {
      const datedSummaries = sortedSummaries.filter(s => s.startDate && s.endDate);
      if (datedSummaries.length > 0) {
        const first = datedSummaries[0];
        const last = datedSummaries[datedSummaries.length - 1];
        const startLabel = `${first.startDate.getFullYear()}${String(first.startDate.getMonth() + 1).padStart(2, '0')}16`;
        const endLabel = `${last.endDate.getFullYear()}${String(last.endDate.getMonth() + 1).padStart(2, '0')}15`;
        downloadPeriodId = `${startLabel}-${endLabel}`;
      } else {
        downloadPeriodId = 'multi-month';
      }
    } else if (sortedSummaries[0]) {
      const primary = sortedSummaries[0];
      if (primary.year !== null && primary.startMonthRaw) {
        const paddedMonth = primary.startMonthRaw.padStart(2, '0');
        downloadPeriodId = `${primary.year}${paddedMonth}`;
      } else if (primary.result && primary.result.sheetName) {
        downloadPeriodId = primary.result.sheetName;
      }
    }

    setupDownload(resolvedStoreName, downloadPeriodId, results);

    if (openSourceBtn) {
      if (isMultiSheetMode) {
        openSourceBtn.disabled = true;
        openSourceBtn.title = '月横断計算モードでは利用できません。';
      } else {
        const preferredSheetId = (requestedGids[0] && requestedGids[0] !== '')
          ? requestedGids[0]
          : (summaries[0] ? summaries[0].sheetId : undefined);
        setupSourceOpener(store.url, preferredSheetId);
      }
    }
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
