const DAY_IN_MS = 24 * 60 * 60 * 1000;
const CROSS_STORE_LOADING_MESSAGE = [
  '店舗横断計算モードでは複数店舗のデータを読み込むため、通常の計算よりも時間がかかります。',
  'しばらくお待ち下さい。'
].join('\n');

function showToastWithNativeNotice(message, options) {
  if (!message) {
    return null;
  }
  if (typeof window === 'undefined') {
    return null;
  }
  if (typeof window.showToastWithFeedback === 'function') {
    return window.showToastWithFeedback(message, options);
  }
  let toastHandle = null;
  if (typeof window.showToast === 'function') {
    toastHandle = window.showToast(message, options);
  }
  if (typeof window.notifyPlatformFeedback === 'function') {
    window.notifyPlatformFeedback(message, options);
  }
  return toastHandle;
}

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
  startLoading(
    statusEl,
    isCrossStoreMode ? CROSS_STORE_LOADING_MESSAGE : '読込中・・・',
    { disableSlowNote: isCrossStoreMode }
  );
  const ensureDownloadOverlayOpen = () => {
    const overlay = document.getElementById('download-overlay');
    if (!overlay) {
      return;
    }
    const style = window.getComputedStyle(overlay);
    if (style.display !== 'flex') {
      const button = document.getElementById('download');
      if (button) {
        button.click();
      }
    }
  };

  const closeDownloadOverlay = () => {
    const overlay = document.getElementById('download-overlay');
    if (!overlay) {
      return;
    }
    const style = window.getComputedStyle(overlay);
    if (style.display === 'flex') {
      const closeBtn = document.getElementById('download-close');
      if (closeBtn) {
        closeBtn.click();
      } else {
        overlay.style.display = 'none';
      }
    }
  };

  initializeHelp('help/payroll.txt', {
    pageKey: 'payroll',
    showPrompt: false,
    autoStartIf: ({ hasAutoStartFlag }) => hasAutoStartFlag,
    onFinish: () => {
      closeDownloadOverlay();
    },
    steps: {
      back: '#payroll-back',
      restart: '#payroll-home',
      setBase: '#set-base-wage',
      setTransport: '#set-transport',
      firstEmployee: {
        getElement: () => document.querySelector('#employees tbody tr') || document.getElementById('employees')
      },
      download: '#download',
      downloadOptions: {
        selector: '#download-result-xlsx',
        onEnter: ensureDownloadOverlayOpen,
        onExit: closeDownloadOverlay
      },
      source: '#open-source',
      help: () => document.getElementById('help-button')
    }
  });
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

    const downloadedStoreNames = workbookResults
      .map(result => (result && result.selection && result.selection.store && result.selection.store.name
        ? result.selection.store.name
        : ''))
      .filter(name => !!name);
    let downloadMessage = '';
    if (offlineMode && offlineActive && offlineInfo && offlineInfo.fileName) {
      downloadMessage = `${offlineInfo.fileName} のシートを読み込みました。`;
    } else if (downloadedStoreNames.length === 0) {
      downloadMessage = 'シートの読み込みが完了しました。';
    } else if (downloadedStoreNames.length <= 3) {
      downloadMessage = `${downloadedStoreNames.join('・')} のシートを読み込みました。`;
    } else {
      downloadMessage = `${downloadedStoreNames.length}件のシートを読み込みました。`;
    }
    if (failedSheets.length > 0) {
      downloadMessage += '（一部のシートは取得できませんでした）';
    }
    showToastWithNativeNotice(downloadMessage, { duration: 3200 });

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
        const rawBreakdown = employee.breakdown && typeof employee.breakdown === 'object'
          ? employee.breakdown
          : {
              regularHours: employee.regularHours,
              overtimeHours: employee.overtimeHours
            };
        const breakdown = {
          regularHours: Number(rawBreakdown && rawBreakdown.regularHours) || 0,
          overtimeHours: Number(rawBreakdown && rawBreakdown.overtimeHours) || 0
        };
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
            breakdown: { ...breakdown },
            regularHours: breakdown.regularHours,
            overtimeHours: breakdown.overtimeHours,
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
          if (!existing.breakdown || typeof existing.breakdown !== 'object') {
            existing.breakdown = { regularHours: 0, overtimeHours: 0 };
          }
          existing.breakdown.regularHours += breakdown.regularHours;
          existing.breakdown.overtimeHours += breakdown.overtimeHours;
          existing.regularHours = existing.breakdown.regularHours;
          existing.overtimeHours = existing.breakdown.overtimeHours;
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
      if (!employee.breakdown || typeof employee.breakdown !== 'object') {
        employee.breakdown = {
          regularHours: Number(employee.regularHours) || 0,
          overtimeHours: Number(employee.overtimeHours) || 0
        };
      } else {
        employee.breakdown = {
          regularHours: Number(employee.breakdown.regularHours) || 0,
          overtimeHours: Number(employee.breakdown.overtimeHours) || 0
        };
      }
      employee.regularHours = employee.breakdown.regularHours;
      employee.overtimeHours = employee.breakdown.overtimeHours;
      employee.transport = Number(employee.transport || 0);
      employee.salary = employee.baseSalary + employee.transport;
    });

    const totalSalary = results.reduce((sum, r) => sum + (Number(r.salary) || 0), 0);
    const totalSalaryEl = document.getElementById('total-salary');
    totalSalaryEl.textContent = `合計支払い給与：${totalSalary.toLocaleString()}円`;
    let totalSalaryValue = totalSalary;
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
    detailTxtBtn.textContent = 'テキスト形式（.txt）';
    const detailXlsxBtn = document.createElement('button');
    detailXlsxBtn.textContent = 'EXCEL形式（.xlsx）';
    const detailCsvBtn = document.createElement('button');
    detailCsvBtn.textContent = 'CSV形式（.csv）';
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

    function showEmployeeDetail(idx) {
      const employee = results[idx];
      if (!employee) {
        return;
      }

      hideDetailDownload();
      detailDownloadBtn.disabled = true;

      const detailInfo = buildEmployeeDetailDownloadInfo(employee, resolvedStoreName);
      if (!detailInfo) {
        return;
      }

      detailTitle.textContent = detailInfo.employeeName || '';
      detailSummary.innerHTML = '';
      const summaryLines = Array.isArray(detailInfo.summaryLines) ? detailInfo.summaryLines : [];
      summaryLines.forEach(line => {
        const lineEl = document.createElement('div');
        lineEl.className = 'employee-detail-summary-line';
        lineEl.textContent = line;
        detailSummary.appendChild(lineEl);
      });

      detailTableBody.innerHTML = '';
      const rows = Array.isArray(detailInfo.rows) ? detailInfo.rows : [];
      if (rows.length === 0) {
        const emptyRow = document.createElement('tr');
        const emptyCell = document.createElement('td');
        emptyCell.colSpan = 3;
        emptyCell.textContent = '出勤記録がありません';
        emptyRow.appendChild(emptyCell);
        detailTableBody.appendChild(emptyRow);
      } else {
        let currentMonthKey = null;
        rows.forEach(row => {
          if (!row) {
            return;
          }
          if (row.type === 'month') {
            const monthKey = row.label;
            if (monthKey !== currentMonthKey) {
              currentMonthKey = monthKey;
              const monthRow = document.createElement('tr');
              monthRow.className = 'month-row';
              const monthCell = document.createElement('th');
              monthCell.colSpan = 3;
              monthCell.textContent = row.label || '';
              monthRow.appendChild(monthCell);
              detailTableBody.appendChild(monthRow);
            }
            return;
          }
          if (row.type === 'day') {
            const dayRow = document.createElement('tr');
            const dateCell = document.createElement('td');
            dateCell.className = 'date-cell';
            dateCell.textContent = row.dateLabel || '';
            const timeCell = document.createElement('td');
            timeCell.className = 'time-cell';
            const storeCell = document.createElement('td');
            storeCell.className = 'store-cell';
            const times = Array.isArray(row.times) ? row.times : [];
            const stores = Array.isArray(row.stores) ? row.stores : [];
            const maxSegments = Math.max(times.length, stores.length);
            for (let i = 0; i < maxSegments; i += 1) {
              if (i > 0) {
                timeCell.appendChild(document.createElement('br'));
                storeCell.appendChild(document.createElement('br'));
              }
              if (times[i]) {
                timeCell.appendChild(document.createTextNode(times[i]));
              }
              if (stores[i]) {
                storeCell.appendChild(document.createTextNode(stores[i]));
              }
            }
            dayRow.appendChild(dateCell);
            dayRow.appendChild(timeCell);
            dayRow.appendChild(storeCell);
            detailTableBody.appendChild(dayRow);
          }
        });
      }

      currentDetailDownloadInfo = detailInfo;
      detailDownloadBtn.disabled = false;
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

    const scheduleFrame = (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function')
      ? window.requestAnimationFrame.bind(window)
      : (typeof requestAnimationFrame === 'function'
        ? requestAnimationFrame
        : (cb => setTimeout(cb, 16)));
    const dirtyRows = new Set();
    let rafScheduled = false;
    let totalDirty = false;

    function flushUpdates() {
      rafScheduled = false;
      dirtyRows.forEach(idx => {
        const row = tbody.querySelector(`tr[data-idx="${idx}"]`);
        if (!row) return;
        const salaryCell = row.querySelector('.salary-cell');
        if (!salaryCell) return;
        const salaryValue = Number(results[idx] && results[idx].salary) || 0;
        salaryCell.textContent = salaryValue.toLocaleString();
      });
      dirtyRows.clear();
      if (totalDirty) {
        totalSalaryEl.textContent = `合計支払い給与：${totalSalaryValue.toLocaleString()}円`;
        totalDirty = false;
      }
    }

    function queueFlush() {
      if (rafScheduled) return;
      rafScheduled = true;
      scheduleFrame(flushUpdates);
    }

    function recalc(idx, overrides = {}) {
      if (!Number.isInteger(idx) || !results[idx]) {
        totalSalaryValue = results.reduce((sum, r) => sum + (Number(r.salary) || 0), 0);
        totalDirty = true;
        queueFlush();
        return;
      }
      const employee = results[idx];
      const previousSalary = Number(employee.salary) || 0;
      const updates = overrides || {};
      if (Object.prototype.hasOwnProperty.call(updates, 'wage')) {
        const wage = updates.wage;
        if (Number.isFinite(wage)) {
          employee.baseWage = wage;
          const breakdown = employee.breakdown && typeof employee.breakdown === 'object'
            ? employee.breakdown
            : {
                regularHours: Number(employee.regularHours) || 0,
                overtimeHours: Number(employee.overtimeHours) || 0
              };
          const regular = Number(breakdown.regularHours) || 0;
          const overtimeHours = Number(breakdown.overtimeHours) || 0;
          const overtimeRate = Number.isFinite(store.overtime) ? Number(store.overtime) : 1;
          const newBaseSalary = Math.floor((regular * wage) + (overtimeHours * wage * overtimeRate));
          employee.baseSalary = newBaseSalary;
          employee.salary = newBaseSalary + (Number(employee.transport) || 0);
          employee.breakdown = {
            regularHours: regular,
            overtimeHours
          };
          employee.regularHours = regular;
          employee.overtimeHours = overtimeHours;
        }
      }
      if (Object.prototype.hasOwnProperty.call(updates, 'transport')) {
        const transport = Number(updates.transport) || 0;
        employee.transport = transport;
        employee.salary = (Number(employee.baseSalary) || 0) + transport;
      }
      const newSalary = Number(employee.salary) || 0;
      totalSalaryValue += newSalary - previousSalary;
      dirtyRows.add(idx);
      totalDirty = true;
      queueFlush();
    }

    results.forEach((r, idx) => {
      const tr = document.createElement('tr');
      tr.dataset.idx = idx;

      const nameTd = document.createElement('td');
      nameTd.textContent = r.name;
      nameTd.addEventListener('click', () => showEmployeeDetail(idx));

      const wageTd = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'number';
      input.value = r.baseWage;
      input.className = 'wage-input';
      input.dataset.idx = idx;
      input.addEventListener('input', () => {
        const wage = Number(input.value);
        if (!Number.isFinite(wage)) return;
        recalc(idx, { wage });
      });
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
      transportInput.addEventListener('input', () => {
        const value = Number(transportInput.value);
        const transport = Number.isFinite(value) ? value : 0;
        recalc(idx, { transport });
      });
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
    const toastOptions = { duration: 3200 };
    showToastWithNativeNotice('計算が完了しました。', toastOptions);
    if (failedSheets.length > 0) {
      showToastWithNativeNotice(
        `${failedSheets.length}件のシートを読み込めなかったため除外しました。`,
        toastOptions,
      );
    }
    if (processingFailures.length > 0) {
      showToastWithNativeNotice(
        `${processingFailures.length}件のシートにエラーがあったため除外されました。`,
        toastOptions,
      );
    }
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
      if (!Number.isFinite(wage)) {
        return;
      }
      document.querySelectorAll('.wage-input').forEach(input => {
        input.value = wage;
        const idx = Number(input.dataset.idx);
        if (!Number.isInteger(idx) || !results[idx]) return;
        recalc(idx, { wage });
      });
    });

    const transportAllInput = document.getElementById('transport-input');
    transportAllInput.value = 0;
    document.getElementById('set-transport').addEventListener('click', () => {
      const value = Number(transportAllInput.value);
      const transport = Number.isFinite(value) ? value : 0;
      document.querySelectorAll('.transport-input').forEach(input => {
        const idx = Number(input.dataset.idx);
        if (!Number.isInteger(idx)) return;
        const employee = results[idx];
        if (!employee || employee.days === 0) {
          return;
        }
        input.value = transport;
        recalc(idx, { transport });
      });
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
  if (!button) {
    return;
  }
  const overlay = document.createElement('div');
  overlay.id = 'download-overlay';
  overlay.style.display = 'none';
  const popup = document.createElement('div');
  popup.id = 'download-popup';
  const options = document.createElement('div');
  options.id = 'download-options';

  const includeDetail = document.createElement('label');
  includeDetail.id = 'download-include-detail';
  const includeCheckbox = document.createElement('input');
  includeCheckbox.type = 'checkbox';
  includeCheckbox.id = 'download-include-detail-checkbox';
  const includeLabel = document.createElement('span');
  includeLabel.textContent = '詳細を含める';
  includeDetail.appendChild(includeCheckbox);
  includeDetail.appendChild(includeLabel);

  const txtBtn = document.createElement('button');
  txtBtn.id = 'download-result-txt';
  txtBtn.textContent = 'テキスト形式（.txt）';
  const xlsxBtn = document.createElement('button');
  xlsxBtn.id = 'download-result-xlsx';
  xlsxBtn.textContent = 'EXCEL形式（.xlsx）';
  const csvBtn = document.createElement('button');
  csvBtn.id = 'download-result-csv';
  csvBtn.textContent = 'CSV形式（.csv）';

  options.appendChild(includeDetail);
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
    updateFormatButtons();
    overlay.style.display = 'flex';
  });
  closeBtn.addEventListener('click', hide);
  overlay.addEventListener('click', e => { if (e.target === overlay) hide(); });

  function updateFormatButtons() {
    const includeDetails = includeCheckbox.checked;
    txtBtn.disabled = includeDetails;
    csvBtn.disabled = includeDetails;
  }

  includeCheckbox.addEventListener('change', updateFormatButtons);

  txtBtn.addEventListener('click', () => {
    downloadResults(storeName, period, results, 'txt');
    hide();
  });
  xlsxBtn.addEventListener('click', () => {
    downloadResults(storeName, period, results, {
      format: 'xlsx',
      includeDetails: includeCheckbox.checked
    });
    hide();
  });
  csvBtn.addEventListener('click', () => {
    downloadResults(storeName, period, results, 'csv');
    hide();
  });
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

function parseTimeSegment(match) {
  if (!Array.isArray(match)) {
    return null;
  }
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

function buildEmployeeDetailDownloadInfo(employee, defaultStoreName) {
  if (!employee) {
    return null;
  }

  const baseWageValue = Number(employee.baseWage || 0);
  const hoursValueRaw = Number(employee.hours);
  const hoursValue = Number.isFinite(hoursValueRaw) ? hoursValueRaw : 0;
  const daysValueRaw = Number(employee.days);
  const daysValue = Number.isFinite(daysValueRaw) ? daysValueRaw : (employee.days || 0);
  const transportValue = Number(employee.transport || 0);
  const salaryValue = Number(employee.salary || 0);

  const summaryLines = [
    `基本時給：${baseWageValue.toLocaleString()}円`,
    `総勤務時間：${hoursValue.toFixed(2)}時間`,
    `出勤日数：${daysValue}日`,
    `交通費：${transportValue.toLocaleString()}円`,
    `給与：${salaryValue.toLocaleString()}円`
  ];

  const entriesMap = new Map();
  const scheduleBlocks = Array.isArray(employee.scheduleBlocks) ? employee.scheduleBlocks : [];
  scheduleBlocks.forEach(block => {
    if (!block || !Array.isArray(block.schedule) || block.schedule.length === 0) {
      return;
    }
    const baseDate = block.startDate ? new Date(block.startDate) : null;
    if (!baseDate || Number.isNaN(baseDate.getTime())) {
      return;
    }
    const blockStoreName = block.storeName && String(block.storeName).trim()
      ? String(block.storeName).trim()
      : defaultStoreName;
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
  const rows = [];
  let currentMonthLabel = null;
  entries.forEach(entry => {
    const monthLabel = `${entry.date.getFullYear()}年${entry.date.getMonth() + 1}月`;
    if (monthLabel !== currentMonthLabel) {
      currentMonthLabel = monthLabel;
      rows.push({ type: 'month', label: monthLabel });
    }
    const sortedSegments = Array.from(entry.segments.values())
      .sort((a, b) => a.sortKey.localeCompare(b.sortKey));
    const segmentDetails = sortedSegments.map(segment => {
      const storeNames = Array.from(segment.storeNames)
        .filter(Boolean)
        .sort((a, b) => a.localeCompare(b, 'ja'));
      return {
        timeLabel: segment.display,
        storeLabel: storeNames.length > 0 ? storeNames.join('、') : ''
      };
    });
    rows.push({
      type: 'day',
      dateLabel: `${entry.date.getDate()}日`,
      times: segmentDetails.map(detail => detail.timeLabel),
      stores: segmentDetails.map(detail => detail.storeLabel)
    });
  });

  return {
    employeeName: employee.name || '',
    summaryLines,
    rows
  };
}

function detailInfoToAoa(detailInfo) {
  const aoa = [];
  if (!detailInfo) {
    return aoa;
  }

  aoa.push(['従業員名', detailInfo.employeeName || '']);
  const summaryLines = Array.isArray(detailInfo.summaryLines) ? detailInfo.summaryLines : [];
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
      if (!row) {
        return;
      }
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

  return aoa;
}

function sanitizeSheetName(name) {
  if (typeof name !== 'string') {
    return '';
  }
  return name.replace(/[\\/?*\[\]:]/g, '_').trim();
}

function makeUniqueSheetName(baseName, usedNames) {
  const MAX_LENGTH = 31;
  const base = sanitizeSheetName(baseName) || '詳細';
  let candidate = base.slice(0, MAX_LENGTH) || '詳細';
  let counter = 1;
  while (usedNames.has(candidate)) {
    counter += 1;
    const suffix = `_${counter}`;
    const trimmedBase = base.slice(0, Math.max(0, MAX_LENGTH - suffix.length)) || '詳細';
    candidate = `${trimmedBase}${suffix}`;
  }
  usedNames.add(candidate);
  return candidate;
}

async function downloadResults(storeName, period, results, options = {}) {
  let format = 'xlsx';
  let includeDetails = false;
  if (typeof options === 'string') {
    format = options;
  } else if (options && typeof options === 'object') {
    if (options.format) {
      format = options.format;
    }
    includeDetails = Boolean(options.includeDetails);
  }

  const normalizedFormat = typeof format === 'string' ? format.toLowerCase() : 'xlsx';
  const formatLabelMap = { xlsx: 'EXCEL', csv: 'CSV', txt: 'テキスト' };
  const formatLabel = formatLabelMap[normalizedFormat] || normalizedFormat.toUpperCase();
  const subjectParts = [];
  if (period) {
    subjectParts.push(String(period));
  }
  if (storeName) {
    subjectParts.push(String(storeName));
  }
  const subjectPrefix = subjectParts.length > 0 ? `${subjectParts.join('・')}の` : '';
  const detailSuffix = includeDetails ? '（詳細付き）' : '';
  showToastWithNativeNotice(
    `${subjectPrefix}計算結果${detailSuffix}の${formatLabel}ダウンロードを開始しました。`,
    { duration: 3200 },
  );
  const aoa = [
    ['従業員名', '基本時給', '勤務時間', '出勤日数', '交通費', '給与'],
    ...results.map(r => [r.name, r.baseWage, r.hours, r.days, r.transport, r.salary])
  ];
  const total = results.reduce((sum, r) => sum + (Number(r.salary) || 0), 0);
  aoa.push(['合計支払い給与', '', '', '', '', total]);

  if (normalizedFormat === 'csv') {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const csv = XLSX.utils.sheet_to_csv(ws);
    downloadBlob(csv, `${period}_${storeName}.csv`, 'text/csv;charset=utf-8', { addBom: true });
    return;
  }

  if (normalizedFormat === 'txt') {
    const text = aoa.map(row => row.join('\t')).join('\n');
    downloadBlob(text, `${period}_${storeName}.txt`, 'text/plain;charset=utf-8', { addBom: true });
    return;
  }

  const wb = XLSX.utils.book_new();
  const summarySheet = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, summarySheet, '結果');
  const usedSheetNames = new Set(['結果']);

  if (includeDetails) {
    results.forEach(employee => {
      const detailInfo = buildEmployeeDetailDownloadInfo(employee, storeName);
      if (!detailInfo) {
        return;
      }
      const detailAoa = detailInfoToAoa(detailInfo);
      const detailSheet = XLSX.utils.aoa_to_sheet(detailAoa);
      const baseSheetName = detailInfo.employeeName
        ? `詳細_${detailInfo.employeeName}`
        : '詳細';
      const sheetName = makeUniqueSheetName(baseSheetName, usedSheetNames);
      XLSX.utils.book_append_sheet(wb, detailSheet, sheetName);
    });
  }

  const parts = [];
  const periodPart = sanitizeFileNameComponent(period || '');
  if (periodPart) parts.push(periodPart);
  const storePart = sanitizeFileNameComponent(storeName || '');
  if (storePart) parts.push(storePart);
  const fileName = `${parts.join('_') || 'result'}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

function downloadEmployeeDetail(storeName, period, detailInfo, format) {
  if (!detailInfo) {
    return;
  }

  const aoa = detailInfoToAoa(detailInfo);
  const normalizedFormat = typeof format === 'string' ? format.toLowerCase() : 'xlsx';
  const formatLabelMap = { txt: 'テキスト', csv: 'CSV', xlsx: 'EXCEL' };
  const formatLabel = formatLabelMap[normalizedFormat] || normalizedFormat.toUpperCase();
  const subjectParts = [];
  if (period) {
    subjectParts.push(String(period));
  }
  if (storeName) {
    subjectParts.push(String(storeName));
  }
  const subjectPrefix = subjectParts.length > 0 ? `${subjectParts.join('・')}の` : '';
  const employeePrefix = detailInfo.employeeName ? `${detailInfo.employeeName}さんの` : '';
  showToastWithNativeNotice(
    `${subjectPrefix}${employeePrefix}勤務詳細の${formatLabel}ダウンロードを開始しました。`,
    { duration: 3200 },
  );

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
