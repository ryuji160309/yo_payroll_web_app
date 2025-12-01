let reportChart = null;
const periodCache = new Map();

function parseWorkbookPeriod(data) {
  const year = data[1] && Number.parseInt(data[1][2], 10);
  const startMonth = data[1] && Number.parseInt(data[1][4], 10);
  if (!Number.isFinite(year) || !Number.isFinite(startMonth)) {
    return null;
  }
  const startDate = new Date(year, startMonth - 1, 16);
  const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 15, 23, 59, 59, 999);
  return { startDate, endDate };
}

function formatPeriodLabel(period) {
  if (!period || !(period.startDate instanceof Date) || !(period.endDate instanceof Date)) {
    return '';
  }
  const startYear = period.startDate.getFullYear();
  const startMonth = period.startDate.getMonth() + 1;
  const endYear = period.endDate.getFullYear();
  const endMonth = period.endDate.getMonth() + 1;
  const startLabel = `${startYear}年${startMonth}月`;
  const endLabel = endYear !== startYear ? `${endYear}年${endMonth}月` : `${endMonth}月`;
  return `${startLabel}度（${startMonth}月16日〜${endLabel}15日）`;
}

function populateStoreOptions(storeSelect, stores) {
  const entries = Object.entries(stores || {});
  storeSelect.innerHTML = '';
  entries.forEach(([key, store]) => {
    const option = document.createElement('option');
    option.value = key;
    option.textContent = store.name;
    storeSelect.appendChild(option);
  });
}

function getSelectedStore(stores, selectEl) {
  const key = selectEl.value;
  return key && stores[key] ? { key, store: stores[key] } : null;
}

function setStatus(message) {
  const statusEl = document.getElementById('report-status');
  if (!statusEl) return;
  statusEl.textContent = message || '';
}

function formatNumber(value) {
  return Number.isFinite(value) ? value.toLocaleString('ja-JP') : '--';
}

function updateSummary({ totalHours, overtimeHours, totalSalary, workdays }) {
  const hoursEl = document.getElementById('summary-total-hours');
  const overtimeEl = document.getElementById('summary-overtime');
  const salaryEl = document.getElementById('summary-salary');
  const daysEl = document.getElementById('summary-days');

  if (hoursEl) hoursEl.textContent = formatNumber(totalHours);
  if (overtimeEl) overtimeEl.textContent = formatNumber(overtimeHours);
  if (salaryEl) salaryEl.textContent = `${formatNumber(totalSalary)} 円`;
  if (daysEl) daysEl.textContent = formatNumber(workdays);
}

function renderChart(summary) {
  const canvas = document.getElementById('report-chart');
  if (!canvas || typeof Chart === 'undefined') return;

  const ctx = canvas.getContext('2d');
  const labels = ['総労働時間', '残業時間', '出勤日数', '概算給与'];
  const data = [summary.totalHours, summary.overtimeHours, summary.workdays, summary.totalSalary];
  if (reportChart) {
    reportChart.destroy();
  }
  reportChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        {
          label: '月次サマリー',
          data,
          backgroundColor: ['#0f4a73', '#1f7aa3', '#4a90e2', '#9bc6ff'],
        }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        legend: {
          display: false
        },
        tooltip: {
          callbacks: {
            label: context => `${context.label}: ${formatNumber(context.parsed.y)}`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            callback: value => formatNumber(value)
          }
        }
      }
    }
  });
}

async function loadPeriodsForStore(storeKey, store) {
  const monthSelect = document.getElementById('report-month');
  if (!monthSelect) return [];
  setStatus('シートを読み込んでいます…');
  monthSelect.innerHTML = '';
  const options = [];
  try {
    const sheetList = await fetchSheetList(store.url, { allowOffline: true });
    const list = Array.isArray(sheetList) ? sheetList : [];
    for (const meta of list) {
      try {
        const workbook = await fetchWorkbook(store.url, meta.index, { allowOffline: true });
        const period = parseWorkbookPeriod(workbook.data || []);
        if (!period) {
          continue;
        }
        const label = formatPeriodLabel(period);
        const value = `${storeKey}|${meta.index}`;
        periodCache.set(value, { workbook, period });
        options.push({ value, label, meta });
      } catch (error) {
        // Skip this sheet but continue loading others.
      }
    }
  } catch (error) {
    setStatus('シート情報を取得できませんでした。オンライン接続を確認してください。');
  }

  options.sort((a, b) => b.label.localeCompare(a.label, 'ja'));
  options.forEach(entry => {
    const option = document.createElement('option');
    option.value = entry.value;
    option.textContent = entry.label;
    monthSelect.appendChild(option);
  });
  if (!options.length) {
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = '対象シートが見つかりません';
    monthSelect.appendChild(placeholder);
  }
  setStatus(options.length ? '月を選択して「集計」を押してください。' : '対象月を選択できません。');
  return options;
}

async function runReport(stores) {
  const storeSelect = document.getElementById('report-store');
  const monthSelect = document.getElementById('report-month');
  const periodLabel = document.getElementById('report-period-label');
  const selection = getSelectedStore(stores, storeSelect);
  if (!selection) {
    setStatus('店舗を選択してください。');
    return;
  }
  const store = selection.store;
  const monthValue = monthSelect.value;
  const cacheEntry = periodCache.get(monthValue);
  if (!monthValue || !cacheEntry) {
    setStatus('対象月を選択してください。');
    return;
  }
  setStatus('集計中です…');
  try {
    const { workbook, period } = cacheEntry;
    const payroll = calculatePayroll(
      workbook.data,
      store.baseWage,
      store.overtime,
      store.excludeWords || []
    );
    const summary = payroll.results.reduce((acc, result) => ({
      totalHours: acc.totalHours + (Number(result.hours) || 0),
      overtimeHours: acc.overtimeHours + (Number(result.overtimeHours) || 0),
      totalSalary: acc.totalSalary + (Number(result.salary) || 0),
      workdays: acc.workdays + (Number(result.days) || 0)
    }), { totalHours: 0, overtimeHours: 0, totalSalary: 0, workdays: 0 });

    if (periodLabel) {
      periodLabel.textContent = `${store.name} / ${formatPeriodLabel(period)}`;
    }
    updateSummary(summary);
    renderChart(summary);
    setStatus('集計が完了しました。');
  } catch (error) {
    console.error(error);
    setStatus('集計処理中にエラーが発生しました。');
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  await ensureSettingsLoaded();
  const stores = loadStores();
  const storeSelect = document.getElementById('report-store');
  const runButton = document.getElementById('run-report');
  populateStoreOptions(storeSelect, stores);

  if (storeSelect.options.length > 0) {
    storeSelect.selectedIndex = 0;
    const { key, store } = getSelectedStore(stores, storeSelect) || {};
    if (key && store) {
      await loadPeriodsForStore(key, store);
    }
  }

  storeSelect.addEventListener('change', async () => {
    const selection = getSelectedStore(stores, storeSelect);
    if (selection) {
      await loadPeriodsForStore(selection.key, selection.store);
    }
  });

  runButton.addEventListener('click', () => runReport(stores));
});
