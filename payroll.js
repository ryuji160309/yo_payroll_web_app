const DAY_IN_MS = 24 * 60 * 60 * 1000;
const CROSS_STORE_LOADING_MESSAGE = [
  '店舗横断計算モードでは複数店舗のデータを読み込むため、通常の計算よりも時間がかかります。',
  'しばらくお待ち下さい。'
].join('\n');

function isValidDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

function formatPeriodRange(startDate, endDate) {
  if (!isValidDate(startDate) || !isValidDate(endDate)) {
    return '';
  }
  const startYear = startDate.getFullYear();
  const startMonth = startDate.getMonth() + 1;
  const startDay = startDate.getDate();
  const endYear = endDate.getFullYear();
  const endMonth = endDate.getMonth() + 1;
  const endDay = endDate.getDate();
  const startLabel = `${startYear}年${startMonth}月${startDay}日`;
  const endLabelYear = startYear === endYear ? '' : `${endYear}年`;
  const endLabel = `${endLabelYear}${endMonth}月${endDay}日`;
  return `${startLabel}～${endLabel}`;
}

document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('status');
  const params = new URLSearchParams(location.search);
  const selectionsParamRaw = params.get('selections');
  const isCrossStoreMode = selectionsParamRaw !== null && selectionsParamRaw.trim() !== '';
  startLoading(statusEl, isCrossStoreMode ? CROSS_STORE_LOADING_MESSAGE : '読込中・・・');
  initializeHelp('help/payroll.txt');
  await ensureSettingsLoaded();

  let downloadPeriodId = 'result';
  let currentDetailDownloadInfo = null;

  const offlineRequested = params.get('offline') === '1';
  const offlineMode = offlineRequested && !isCrossStoreMode;
  const offlineInfo = typeof getOfflineWorkbookInfo === 'function' ? getOfflineWorkbookInfo() : null;
  const offlineActive = typeof isOfflineWorkbookActive === 'function' && isOfflineWorkbookActive();
  if (offlineMode && !offlineActive) {
    stopLoading(statusEl);
    statusEl.textContent = 'ローカルファイルを利用できません。トップに戻って読み込み直してください。';
    return;
  }

  function normalizeStoreKey(value) {
    if (value === null || value === undefined) {
      return '';
    }
    const trimmed = value.trim();
    if (!trimmed) {
      return '';
    }
    try {
      return decodeURIComponent(trimmed);
    } catch (e) {
      return trimmed;
    }
  }

  const targetSelections = [];
  const requestedGids = [];
  const seenSelectionKeys = new Set();

  if (isCrossStoreMode) {
    const segments = selectionsParamRaw.split(',').map(seg => seg.trim()).filter(Boolean);
    segments.forEach(segment => {
      const parts = segment.split(':');
      if (parts.length < 2) {
        return;
      }
      const storeKey = normalizeStoreKey(parts[0]);
      const sheetIndex = Number.parseInt(parts[1], 10);
      if (!storeKey || !Number.isFinite(sheetIndex)) {
        return;
      }
      const identifier = `${storeKey}|${sheetIndex}`;
      if (seenSelectionKeys.has(identifier)) {
        return;
      }
      const sheetId = parts.length > 2 ? parts.slice(2).join(':') : '';
      const store = getStore(storeKey);
      if (!store) {
        return;
      }
      seenSelectionKeys.add(identifier);
      targetSelections.push({ storeKey, store, sheetIndex, sheetId: sheetId || '' });
    });
  } else {
    const storeKey = normalizeStoreKey(params.get('store'));
    const store = storeKey ? getStore(storeKey) : null;
    if (!store) {
      stopLoading(statusEl);
      return;
    }
    const multiSheetsParam = params.get('sheets');
    let sheetIndices = multiSheetsParam
      ? multiSheetsParam.split(',').map(v => parseInt(v, 10))
      : [];
    const seenIndices = new Set();
    sheetIndices = sheetIndices.filter(idx => {
      if (!Number.isFinite(idx) || seenIndices.has(idx)) {
        return false;
      }
      seenIndices.add(idx);
      return true;
    });
    if (sheetIndices.length === 0) {
      const fallbackSheet = parseInt(params.get('sheet'), 10);
      sheetIndices = [Number.isFinite(fallbackSheet) ? fallbackSheet : 0];
    }
    const sheetGidParam = params.get('gid');
    const multiGidsParam = params.get('gids');
    const rawGids = multiGidsParam
      ? multiGidsParam.split(',')
      : (sheetGidParam !== null ? [sheetGidParam] : []);
    const sanitizedGids = sheetIndices.map((_, idx) => rawGids[idx] !== undefined ? rawGids[idx] : '');
    sanitizedGids.forEach(gid => requestedGids.push(gid));
    sheetIndices.forEach((sheetIndex, idx) => {
      const identifier = `${storeKey}|${sheetIndex}`;
      if (seenSelectionKeys.has(identifier)) {
        return;
      }
      seenSelectionKeys.add(identifier);
      targetSelections.push({ storeKey, store, sheetIndex, sheetId: sanitizedGids[idx] || '' });
    });
  }

  if (targetSelections.length === 0) {
    stopLoading(statusEl);
    statusEl.textContent = '計算対象のシートが選択されていません。前の画面に戻って選び直してください。';
    return;
  }

  const store = targetSelections[0].store;
  if (!store) {
    stopLoading(statusEl);
    statusEl.textContent = '店舗情報を取得できませんでした。設定を確認してください。';
    return;
  }

  if (!offlineMode && typeof setOfflineWorkbookActive === 'function') {
    try {
      setOfflineWorkbookActive(false);
    } catch (e) {
      // Ignore failures when disabling offline mode.
    }
  }

  const isMultiSheetMode = targetSelections.length > 1;

  const titleEl = document.getElementById('store-name');
  if (titleEl) {
    if (isCrossStoreMode) {
      titleEl.textContent = '店舗横断計算モード';
      document.title = '店舗横断計算モード - 給与計算';
    } else if (offlineMode && offlineActive && offlineInfo && offlineInfo.fileName) {
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
    const fetchPromises = targetSelections.map(selection =>
      fetchWorkbook(selection.store.url, selection.sheetIndex, { allowOffline: !isCrossStoreMode && offlineMode })
    );
    const settledResults = await Promise.allSettled(fetchPromises);
    const workbookResults = [];
    const failedSheets = [];
    settledResults.forEach((res, idx) => {
      const selection = targetSelections[idx];
      if (!selection) {
        return;
      }
      if (res.status === 'fulfilled') {
        workbookResults.push({ selection, workbook: res.value });
      } else {
        failedSheets.push({ selection, reason: res.reason });
      }
    });
    if (workbookResults.length === 0) {
      throw new Error('シートの読み込みに失敗しました');
    }

    stopLoading(statusEl);

    const processingFailures = [];
    const summaries = [];
    workbookResults.forEach(({ selection, workbook }) => {
      const data = workbook.data;
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
        periodLabel = workbook.sheetName || `シート${workbook.sheetIndex + 1}`;
      }
      try {
        const payrollResult = calculatePayroll(
          data,
          selection.store.baseWage,
          selection.store.overtime,
          selection.store.excludeWords || []
        );
        const nameRow = data[36] || [];
        const sheetStoreName = nameRow[14];
        summaries.push({
          selection,
          result: workbook,
          data,
          payrollResult,
          startDate,
          endDate,
          periodLabel,
          startMonthRaw: startMonthRaw !== undefined && startMonthRaw !== null ? String(startMonthRaw) : '',
          year: Number.isFinite(year) ? year : null,
          sheetStoreName,
          sheetId: workbook.sheetId,
        });
      } catch (err) {
        processingFailures.push({ sheetIndex: workbook.sheetIndex, reason: err, selection });
      }
    });

    if (summaries.length === 0) {
      throw new Error('有効なシートがありません');
    }

    const resolvedStoreName = isCrossStoreMode
      ? '店舗横断計算モード'
      : (summaries.map(s => s.sheetStoreName).find(name => !!name) || store.name);
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
        const chronologicalSummaries = summaries
          .filter(summary => isValidDate(summary.startDate) && isValidDate(summary.endDate))
          .slice()
          .sort((a, b) => a.startDate - b.startDate);

        const mergedRanges = [];
        chronologicalSummaries.forEach(summary => {
          const lastRange = mergedRanges[mergedRanges.length - 1];
          if (!lastRange) {
            mergedRanges.push({ start: summary.startDate, end: summary.endDate });
            return;
          }
          const gap = summary.startDate.getTime() - lastRange.end.getTime();
          if (gap <= DAY_IN_MS) {
            if (summary.endDate.getTime() > lastRange.end.getTime()) {
              lastRange.end = summary.endDate;
            }
          } else {
            mergedRanges.push({ start: summary.startDate, end: summary.endDate });
          }
        });

        const mergedLabels = mergedRanges
          .map(range => formatPeriodRange(range.start, range.end))
          .filter(Boolean);

        if (mergedLabels.length > 0) {
          periodEl.textContent = `選択した月：${mergedLabels.join(' ／ ')}`;
        } else {
          const fallbackLabels = summaries.map(s => s.periodLabel).filter(Boolean);
          periodEl.textContent = fallbackLabels.length
            ? `選択した月：${fallbackLabels.join(' ／ ')}`
            : '選択した月：複数シート';
        }
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
        const scheduleStoreNameRaw = summary.sheetStoreName && String(summary.sheetStoreName).trim()
          ? String(summary.sheetStoreName).trim()
          : (summary.selection && summary.selection.store && summary.selection.store.name
            ? summary.selection.store.name
            : resolvedStoreName);
        const scheduleStoreName = scheduleStoreNameRaw || resolvedStoreName;
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
              ? [{ startDate: blockStartDate, schedule: scheduleSource.slice(), storeName: scheduleStoreName }]
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
            existing.scheduleBlocks.push({ startDate: blockStartDate, schedule: scheduleSource.slice(), storeName: scheduleStoreName });
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
    const detailDownloadBtn = document.createElement('button');
    detailDownloadBtn.id = 'employee-detail-download';
    detailDownloadBtn.textContent = '個別計算結果をダウンロード';
    detailDownloadBtn.disabled = true;
    const detailClose = document.createElement('button');
    detailClose.id = 'employee-detail-close';
    detailClose.textContent = '閉じる';

    detailPopup.appendChild(detailTitle);
    detailPopup.appendChild(detailSummary);
    detailPopup.appendChild(detailTableWrapper);
    detailPopup.appendChild(detailDownloadBtn);
    detailPopup.appendChild(detailClose);
    detailOverlay.appendChild(detailPopup);
    document.body.appendChild(detailOverlay);

    const detailDownloadOverlay = document.createElement('div');
    detailDownloadOverlay.id = 'employee-detail-download-overlay';
    detailDownloadOverlay.style.display = 'none';
    const detailDownloadPopup = document.createElement('div');
    detailDownloadPopup.id = 'employee-detail-download-popup';
    const detailDownloadOptions = document.createElement('div');
    detailDownloadOptions.id = 'employee-detail-download-options';
    const detailTxtBtn = document.createElement('button');
    detailTxtBtn.textContent = 'テキスト形式';
    const detailXlsxBtn = document.createElement('button');
    detailXlsxBtn.textContent = 'EXCEL形式';
    const detailCsvBtn = document.createElement('button');
    detailCsvBtn.textContent = 'CSV形式';
    detailDownloadOptions.appendChild(detailTxtBtn);
    detailDownloadOptions.appendChild(detailXlsxBtn);
    detailDownloadOptions.appendChild(detailCsvBtn);
    const detailDownloadClose = document.createElement('button');
    detailDownloadClose.id = 'employee-detail-download-close';
    detailDownloadClose.textContent = '閉じる';
    detailDownloadPopup.appendChild(detailDownloadOptions);
    detailDownloadPopup.appendChild(detailDownloadClose);
    detailDownloadOverlay.appendChild(detailDownloadPopup);
    document.body.appendChild(detailDownloadOverlay);

    function hideEmployeeDetail() {
      detailOverlay.style.display = 'none';
      detailDownloadOverlay.style.display = 'none';
      currentDetailDownloadInfo = null;
      detailDownloadBtn.disabled = true;
    }

    function hideDetailDownload() {
      detailDownloadOverlay.style.display = 'none';
    }

    function showDetailDownload() {
      if (!currentDetailDownloadInfo) {
        return;
      }
      detailDownloadOverlay.style.display = 'flex';
    }

    detailDownloadBtn.addEventListener('click', () => {
      if (detailDownloadBtn.disabled) {
        return;
      }
      showDetailDownload();
    });
    detailDownloadClose.addEventListener('click', hideDetailDownload);
    detailDownloadOverlay.addEventListener('click', e => {
      if (e.target === detailDownloadOverlay) {
        hideDetailDownload();
      }
    });

    detailTxtBtn.addEventListener('click', () => {
      if (!currentDetailDownloadInfo) return;
      downloadEmployeeDetail(resolvedStoreName, downloadPeriodId, currentDetailDownloadInfo, 'txt');
      hideDetailDownload();
    });
    detailXlsxBtn.addEventListener('click', () => {
      if (!currentDetailDownloadInfo) return;
      downloadEmployeeDetail(resolvedStoreName, downloadPeriodId, currentDetailDownloadInfo, 'xlsx');
      hideDetailDownload();
    });
    detailCsvBtn.addEventListener('click', () => {
      if (!currentDetailDownloadInfo) return;
      downloadEmployeeDetail(resolvedStoreName, downloadPeriodId, currentDetailDownloadInfo, 'csv');
      hideDetailDownload();
    });

    function parseTimeSegment(match) {
      const startHour = match[1].padStart(2, '0');
      const startMinuteRaw = match[2];
      const endHour = match[3].padStart(2, '0');
      const endMinuteRaw = match[4];
      const startMinute = startMinuteRaw !== undefined ? startMinuteRaw.padStart(2, '0') : '00';
      const endMinute = endMinuteRaw !== undefined ? endMinuteRaw.padStart(2, '0') : '00';
      let startText = `${startHour}時`;
      if (startMinuteRaw !== undefined) {
        startText += `${startMinuteRaw}分`;
      }
      let endText = `${endHour}時`;
      if (endMinuteRaw !== undefined) {
        endText += `${endMinuteRaw}分`;
      }
      return {
        display: `${startText}～${endText}`,
        canonical: `${startHour}:${startMinute}-${endHour}:${endMinute}`,
        sortKey: `${startHour}${startMinute}-${endHour}${endMinute}`
      };
    }

    function showEmployeeDetail(idx) {
      const employee = results[idx];
      if (!employee) {
        return;
      }

      hideDetailDownload();
      detailDownloadBtn.disabled = true;

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

      const entriesMap = new Map();
      (employee.scheduleBlocks || []).forEach(block => {
        if (!block || !block.schedule || block.schedule.length === 0) return;
        const baseDate = block.startDate ? new Date(block.startDate) : null;
        if (!baseDate || Number.isNaN(baseDate.getTime())) return;
        const blockStoreName = block.storeName && String(block.storeName).trim()
          ? String(block.storeName).trim()
          : resolvedStoreName;
        block.schedule.forEach((cell, dayIdx) => {
          if (!cell) return;
          const segments = cell.toString().split(',')
            .map(s => s.trim())
            .map(seg => {
              const match = seg.match(TIME_RANGE_REGEX);
              return match ? parseTimeSegment(match) : null;
            })
            .filter(Boolean);
          if (segments.length === 0) return;
          const current = new Date(baseDate);
          current.setDate(baseDate.getDate() + dayIdx);
          const dateKey = `${current.getFullYear()}-${current.getMonth()}-${current.getDate()}`;
          let entry = entriesMap.get(dateKey);
          if (!entry) {
            entry = { date: current, segments: new Map() };
            entriesMap.set(dateKey, entry);
          }
          segments.forEach(segmentInfo => {
            let segment = entry.segments.get(segmentInfo.canonical);
            if (!segment) {
              segment = {
                display: segmentInfo.display,
                sortKey: segmentInfo.sortKey,
                storeNames: new Set()
              };
              entry.segments.set(segmentInfo.canonical, segment);
            }
            if (blockStoreName) {
              segment.storeNames.add(blockStoreName);
            }
          });
        });
      });

      const entries = Array.from(entriesMap.values()).sort((a, b) => a.date - b.date);
      const detailRowsForDownload = [];

      if (entries.length === 0) {
        const emptyRow = document.createElement('tr');
        const emptyCell = document.createElement('td');
        emptyCell.colSpan = 3;
        emptyCell.textContent = '出勤記録がありません';
        emptyRow.appendChild(emptyCell);
        detailTableBody.appendChild(emptyRow);
        currentDetailDownloadInfo = {
          employeeName: employee.name,
          summaryLines: summaryLines.slice(),
          rows: []
        };
        detailDownloadBtn.disabled = false;
      } else {
        let currentMonthKey = null;
        entries.forEach(entry => {
          const monthKey = `${entry.date.getFullYear()}-${entry.date.getMonth()}`;
          if (monthKey !== currentMonthKey) {
            currentMonthKey = monthKey;
            const monthRow = document.createElement('tr');
            monthRow.className = 'month-row';
            const monthCell = document.createElement('th');
            monthCell.colSpan = 3;
            monthCell.textContent = `${entry.date.getFullYear()}年${entry.date.getMonth() + 1}月`;
            monthRow.appendChild(monthCell);
            detailTableBody.appendChild(monthRow);
            detailRowsForDownload.push({ type: 'month', label: monthCell.textContent });
          }

          const row = document.createElement('tr');
          const dateCell = document.createElement('td');
          dateCell.className = 'date-cell';
          dateCell.textContent = `${entry.date.getDate()}日`;
          const timeCell = document.createElement('td');
          timeCell.className = 'time-cell';
          const storeCell = document.createElement('td');
          storeCell.className = 'store-cell';
          const sortedSegments = Array.from(entry.segments.values())
            .sort((a, b) => a.sortKey.localeCompare(b.sortKey));
          const segmentDetails = sortedSegments.map(segment => {
            const storeNames = Array.from(segment.storeNames)
              .filter(Boolean)
              .sort((a, b) => a.localeCompare(b, 'ja'));
            const storeLabel = storeNames.length > 0 ? storeNames.join('、') : '';
            return {
              timeLabel: segment.display,
              storeLabel
            };
          });
          segmentDetails.forEach((detail, segIdx) => {
            if (segIdx > 0) {
              timeCell.appendChild(document.createElement('br'));
              storeCell.appendChild(document.createElement('br'));
            }
            timeCell.appendChild(document.createTextNode(detail.timeLabel));
            storeCell.appendChild(document.createTextNode(detail.storeLabel));
          });
          row.appendChild(dateCell);
          row.appendChild(timeCell);
          row.appendChild(storeCell);
          detailTableBody.appendChild(row);
          detailRowsForDownload.push({
            type: 'day',
            dateLabel: dateCell.textContent,
            times: segmentDetails.map(detail => detail.timeLabel),
            stores: segmentDetails.map(detail => detail.storeLabel)
          });
        });
        currentDetailDownloadInfo = {
          employeeName: employee.name,
          summaryLines: summaryLines.slice(),
          rows: detailRowsForDownload
        };
        detailDownloadBtn.disabled = false;
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
    if (failedSheets.length > 0 || processingFailures.length > 0) {
      const messages = [];
      if (failedSheets.length > 0) {
        messages.push('一部のシートが読み込めなかったため除外しました。');
      }
      if (processingFailures.length > 0) {
        messages.push('一部のシートで計算エラーが発生したため除外しました。');
      }
      statusEl.textContent = messages.join('\n');
    }

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
    downloadPeriodId = 'result';
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
      if (isCrossStoreMode) {
        openSourceBtn.disabled = true;
        openSourceBtn.title = '店舗横断計算モードでは利用できません。';
      } else if (isMultiSheetMode) {
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

function downloadBlob(content, fileName, mimeType, options = {}) {
  const { addBom = false } = options;
  const parts = [];
  if (addBom) {
    parts.push('﻿');
  }
  parts.push(content);
  const blob = new Blob(parts, { type: mimeType });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}

function sanitizeFileNameComponent(text) {
  if (typeof text !== 'string') {
    return '';
  }
  return text.replace(/[\\/:*?"<>|]/g, '_').trim();
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
    downloadBlob(csv, `${period}_${storeName}.csv`, 'text/csv;charset=utf-8', { addBom: true });
  } else if (format === 'txt') {
    const text = aoa.map(row => row.join('\t')).join('\n');
    downloadBlob(text, `${period}_${storeName}.txt`, 'text/plain;charset=utf-8', { addBom: true });
  } else {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '結果');
    XLSX.writeFile(wb, `${period}_${storeName}.xlsx`);
  }
}

function downloadEmployeeDetail(storeName, period, detailInfo, format) {
  if (!detailInfo) {
    return;
  }

  const summaryLines = Array.isArray(detailInfo.summaryLines) ? detailInfo.summaryLines : [];
  const aoa = [];
  aoa.push(['従業員名', detailInfo.employeeName || '']);
  if (summaryLines.length > 0) {
    aoa.push([]);
    summaryLines.forEach(line => {
      aoa.push([line]);
    });
  }
  aoa.push([]);

  const rows = Array.isArray(detailInfo.rows) ? detailInfo.rows : [];
  if (rows.length === 0) {
    aoa.push(['出勤記録がありません']);
  } else {
    rows.forEach(row => {
      if (!row) return;
      if (row.type === 'month') {
        aoa.push([row.label || '', '', '']);
      } else if (row.type === 'day') {
        const times = Array.isArray(row.times) ? row.times : [];
        const stores = Array.isArray(row.stores) ? row.stores : [];
        aoa.push([
          row.dateLabel || '',
          times.join('\n'),
          stores.join('\n')
        ]);
      }
    });
  }

  const baseParts = [];
  const periodPart = sanitizeFileNameComponent(period || '');
  if (periodPart) baseParts.push(periodPart);
  const storePart = sanitizeFileNameComponent(storeName || '');
  if (storePart) baseParts.push(storePart);
  const employeePart = sanitizeFileNameComponent(detailInfo.employeeName || '');
  if (employeePart) baseParts.push(employeePart);
  const baseName = `${baseParts.join('_') || 'employee'}_detail`;

  if (format === 'csv') {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const csv = XLSX.utils.sheet_to_csv(ws);
    downloadBlob(csv, `${baseName}.csv`, 'text/csv;charset=utf-8', { addBom: true });
  } else if (format === 'txt') {
    const text = aoa.map(row => row.join('\t')).join('\n');
    downloadBlob(text, `${baseName}.txt`, 'text/plain;charset=utf-8', { addBom: true });
  } else {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '詳細');
    XLSX.writeFile(wb, `${baseName}.xlsx`);
  }
}
