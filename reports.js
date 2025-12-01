let storeChart = null;
let employeeChart = null;
const storeCache = new Map();
let activeStoreKey = null;
let currentLoadId = 0;

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

function formatMonthLabel(period) {
  if (!period || !(period.startDate instanceof Date)) return '';
  const year = period.startDate.getFullYear();
  const month = period.startDate.getMonth() + 1;
  return `${year}/${String(month).padStart(2, '0')}`;
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

function populateEmployeeOptions(report) {
  const datalist = document.getElementById('employee-options');
  if (!datalist) return;
  datalist.innerHTML = '';
  const names = Array.from(report?.employees?.keys?.() || []).sort((a, b) => a.localeCompare(b, 'ja'));
  names.forEach(name => {
    const option = document.createElement('option');
    option.value = name;
    datalist.appendChild(option);
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

function setEmployeeStatus(message) {
  const statusEl = document.getElementById('employee-status');
  if (!statusEl) return;
  statusEl.textContent = message || '';
}

function formatNumber(value) {
  return Number.isFinite(value) ? value.toLocaleString('ja-JP') : '--';
}

function summarizePayroll(payroll) {
  const summary = payroll.results.reduce((acc, result) => ({
    totalHours: acc.totalHours + (Number(result.hours) || 0),
    overtimeHours: acc.overtimeHours + (Number(result.overtimeHours) || 0),
    totalSalary: acc.totalSalary + (Number(result.salary) || 0),
    workdays: acc.workdays + (Number(result.days) || 0)
  }), { totalHours: 0, overtimeHours: 0, totalSalary: 0, workdays: 0 });
  return { ...summary, headcount: payroll.results.length };
}

function updateSummary({ totalHours, overtimeHours, totalSalary, workdays, headcount }) {
  const hoursEl = document.getElementById('summary-total-hours');
  const overtimeEl = document.getElementById('summary-overtime');
  const salaryEl = document.getElementById('summary-salary');
  const daysEl = document.getElementById('summary-days');
  const headcountEl = document.getElementById('summary-headcount');

  if (hoursEl) hoursEl.textContent = formatNumber(totalHours);
  if (overtimeEl) overtimeEl.textContent = formatNumber(overtimeHours);
  if (salaryEl) salaryEl.textContent = `${formatNumber(totalSalary)} 円`;
  if (daysEl) daysEl.textContent = formatNumber(workdays);
  if (headcountEl) headcountEl.textContent = formatNumber(headcount);
}

function renderStoreTable(entries) {
  const tbody = document.getElementById('store-trend-body');
  if (!tbody) return;
  tbody.innerHTML = '';
  if (!entries || !entries.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = 6;
    cell.textContent = 'シートが見つかりませんでした。';
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }
  const sorted = [...entries].sort((a, b) => b.period.startDate - a.period.startDate);
  sorted.forEach(entry => {
    const row = document.createElement('tr');
    const cols = [
      entry.monthLabel,
      entry.summary.headcount,
      entry.summary.totalHours,
      entry.summary.overtimeHours,
      entry.summary.totalSalary,
      entry.summary.workdays
    ];
    cols.forEach(value => {
      const cell = document.createElement('td');
      cell.textContent = typeof value === 'number' ? formatNumber(value) : value;
      row.appendChild(cell);
    });
    tbody.appendChild(row);
  });
}

function buildStoreRangeLabel(entries) {
  if (!entries || !entries.length) return '';
  const first = entries[0];
  const last = entries[entries.length - 1];
  return `${first.monthLabel} 〜 ${last.monthLabel}（${entries.length}件）`;
}

function renderStoreChart(entries) {
  const canvas = document.getElementById('store-trend-chart');
  if (!canvas || typeof Chart === 'undefined') return;

  if (storeChart) {
    storeChart.destroy();
  }
  if (!entries || !entries.length) {
    storeChart = null;
    return;
  }

  const labels = entries.map(entry => entry.monthLabel);
  const salaryData = entries.map(entry => entry.summary.totalSalary);
  const overtimeData = entries.map(entry => entry.summary.overtimeHours);
  const headcountData = entries.map(entry => entry.summary.headcount);

  const ctx = canvas.getContext('2d');
  storeChart = new Chart(ctx, {
    data: {
      labels,
      datasets: [
        {
          type: 'bar',
          label: '総支払給与',
          data: salaryData,
          backgroundColor: '#0f4a73',
          yAxisID: 'ySalary'
        },
        {
          type: 'line',
          label: '残業時間',
          data: overtimeData,
          borderColor: '#ff9933',
          backgroundColor: 'rgba(255, 153, 51, 0.25)',
          tension: 0.25,
          yAxisID: 'yHours'
        },
        {
          type: 'line',
          label: '従業員数',
          data: headcountData,
          borderColor: '#4a90e2',
          borderDash: [6, 4],
          tension: 0.25,
          yAxisID: 'yHeadcount'
        }
      ]
    },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        tooltip: {
          callbacks: {
            label: context => `${context.dataset.label}: ${formatNumber(context.parsed.y)}`
          }
        },
        legend: {
          labels: {
            usePointStyle: true
          }
        }
      },
      scales: {
        ySalary: {
          position: 'left',
          title: { display: true, text: '概算給与 (円)' },
          ticks: { callback: value => formatNumber(value) }
        },
        yHours: {
          position: 'right',
          grid: { drawOnChartArea: false },
          title: { display: true, text: '時間' },
          ticks: { callback: value => formatNumber(value) }
        },
        yHeadcount: {
          position: 'right',
          grid: { drawOnChartArea: false },
          title: { display: true, text: '従業員数' },
          offset: true,
          ticks: { callback: value => formatNumber(value) }
        }
      }
    }
  });
}

function renderEmployeeChart(name, timeline) {
  const canvas = document.getElementById('employee-trend-chart');
  if (!canvas || typeof Chart === 'undefined') return;

  if (employeeChart) {
    employeeChart.destroy();
  }
  if (!timeline || !timeline.length) {
    employeeChart = null;
    return;
  }

  const labels = timeline.map(entry => entry.monthLabel);
  const hours = timeline.map(entry => entry.hours);
  const overtime = timeline.map(entry => entry.overtimeHours);
  const salary = timeline.map(entry => entry.salary);

  const ctx = canvas.getContext('2d');
  employeeChart = new Chart(ctx, {
    data: {
      labels,
      datasets: [
        {
          type: 'bar',
          label: '概算給与',
          data: salary,
          backgroundColor: '#1f7aa3',
          yAxisID: 'ySalary'
        },
        {
          type: 'line',
          label: '総労働時間',
          data: hours,
          borderColor: '#0f4a73',
          tension: 0.25,
          yAxisID: 'yHours'
        },
        {
          type: 'line',
          label: '残業時間',
          data: overtime,
          borderColor: '#ff9933',
          borderDash: [6, 4],
          tension: 0.25,
          yAxisID: 'yHours'
        }
      ]
    },
    options: {
      responsive: true,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { labels: { usePointStyle: true } },
        tooltip: {
          callbacks: {
            label: context => `${context.dataset.label}: ${formatNumber(context.parsed.y)}`
          }
        },
        title: {
          display: true,
          text: `${name} の推移`
        }
      },
      scales: {
        ySalary: {
          position: 'left',
          title: { display: true, text: '概算給与 (円)' },
          ticks: { callback: value => formatNumber(value) }
        },
        yHours: {
          position: 'right',
          grid: { drawOnChartArea: false },
          title: { display: true, text: '時間' },
          ticks: { callback: value => formatNumber(value) }
        }
      }
    }
  });
}

function renderLatestSummary(report, storeName) {
  const periodLabel = document.getElementById('report-period-label');
  if (!report || !report.summaries || !report.summaries.length) {
    updateSummary({});
    if (periodLabel) {
      periodLabel.textContent = `${storeName} / データなし`;
    }
    return;
  }
  const latest = report.summaries[report.summaries.length - 1];
  if (periodLabel) {
    periodLabel.textContent = `${storeName} / ${latest.label}`;
  }
  updateSummary(latest.summary);
}

async function buildStoreReport(storeKey, store) {
  setStatus('全てのシートを読み込んでいます…');
  const sheetList = await fetchSheetList(store.url, { allowOffline: true });
  const entries = [];
  const employeeMap = new Map();

  for (const meta of sheetList || []) {
    try {
      const workbook = await fetchWorkbook(store.url, meta.index, { allowOffline: true });
      const period = parseWorkbookPeriod(workbook.data || []);
      if (!period) continue;
      const payroll = calculatePayroll(
        workbook.data,
        store.baseWage,
        store.overtime,
        store.excludeWords || []
      );
      const summary = summarizePayroll(payroll);
      const monthLabel = formatMonthLabel(period);
      const label = formatPeriodLabel(period);
      entries.push({
        key: `${storeKey}|${meta.index}`,
        monthLabel,
        label,
        period,
        summary,
        payroll
      });

      payroll.results.forEach(result => {
        const timeline = employeeMap.get(result.name) || [];
        timeline.push({
          name: result.name,
          monthLabel,
          label,
          period,
          salary: Number(result.salary) || 0,
          hours: Number(result.hours) || 0,
          overtimeHours: Number(result.overtimeHours) || 0,
          days: Number(result.days) || 0
        });
        employeeMap.set(result.name, timeline);
      });
    } catch (error) {
      console.error('Failed to load sheet', meta, error);
    }
  }

  entries.sort((a, b) => a.period.startDate - b.period.startDate);
  for (const timeline of employeeMap.values()) {
    timeline.sort((a, b) => a.period.startDate - b.period.startDate);
  }

  return { storeKey, storeName: store.name, summaries: entries, employees: employeeMap };
}

function renderReport(report, storeName) {
  renderLatestSummary(report, storeName);
  populateEmployeeOptions(report);
  const rangeLabel = document.getElementById('store-range-label');
  if (rangeLabel) {
    rangeLabel.textContent = buildStoreRangeLabel(report?.summaries);
  }
  renderStoreChart(report?.summaries || []);
  renderStoreTable(report?.summaries || []);
  setStatus(report?.summaries?.length ? '最新データを表示しています。' : '対象のシートがありません。');
}

function handleEmployeeFilter(report) {
  const input = document.getElementById('employee-filter');
  const name = input ? input.value.trim() : '';
  if (!input) return;

  if (!name) {
    setEmployeeStatus('従業員名を入力すると推移を表示します。');
    renderEmployeeChart('', []);
    return;
  }
  const timeline = report && report.employees ? report.employees.get(name) : null;
  if (!timeline || !timeline.length) {
    setEmployeeStatus(`「${name}」のデータが見つかりませんでした。`);
    renderEmployeeChart(name, []);
    return;
  }
  setEmployeeStatus(`「${name}」の${timeline.length}件の推移を表示しています。`);
  renderEmployeeChart(name, timeline);
}

async function loadAndRenderStore(stores, forceReload = false) {
  const storeSelect = document.getElementById('report-store');
  const selection = getSelectedStore(stores, storeSelect);
  if (!selection) {
    setStatus('店舗を選択してください。');
    return;
  }

  const loadId = ++currentLoadId;
  activeStoreKey = selection.key;
  setStatus('データを取得しています…');

  if (forceReload) {
    storeCache.delete(selection.key);
  }

  let report = storeCache.get(selection.key);
  try {
    if (!report) {
      report = await buildStoreReport(selection.key, selection.store);
      storeCache.set(selection.key, report);
    }
    if (loadId !== currentLoadId) {
      return; // stale
    }
    renderReport(report, selection.store.name);
    handleEmployeeFilter(report);
  } catch (error) {
    console.error(error);
    setStatus('レポートを取得できませんでした。オンライン接続を確認してください。');
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  await ensureSettingsLoaded();
  const stores = loadStores();
  const storeSelect = document.getElementById('report-store');
  const runButton = document.getElementById('run-report');
  const employeeFilter = document.getElementById('employee-filter');

  populateStoreOptions(storeSelect, stores);
  if (storeSelect.options.length > 0) {
    storeSelect.selectedIndex = 0;
    await loadAndRenderStore(stores);
  }

  storeSelect.addEventListener('change', async () => {
    const input = document.getElementById('employee-filter');
    if (input) input.value = '';
    setEmployeeStatus('');
    await loadAndRenderStore(stores);
  });

  runButton.addEventListener('click', async () => loadAndRenderStore(stores, true));
  employeeFilter.addEventListener('change', () => handleEmployeeFilter(storeCache.get(activeStoreKey)));
  employeeFilter.addEventListener('keyup', () => handleEmployeeFilter(storeCache.get(activeStoreKey)));
});
