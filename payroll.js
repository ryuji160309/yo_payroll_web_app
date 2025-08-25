document.addEventListener('DOMContentLoaded', async () => {
  document.getElementById('version').textContent = `ver.${APP_VERSION}`;
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');
  const sheetName = params.get('sheet');
  const store = getStore(storeKey);
  if (!store) return;
  try {
    const wb = await fetchWorkbook(store.url);
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
    const year = data[1] && data[1][0];
    const startMonth = data[3] && data[3][14];
    const endMonth = ('0' + (((parseInt(startMonth, 10) || 0) % 12) + 1)).slice(-2);
    document.getElementById('period').textContent = `${year}年${startMonth}月16日～${endMonth}月15日`;
    const nameRow = data[36] || [];
    const storeName = nameRow.slice(14, 25).find(v => v) || store.name;
    document.getElementById('store-name').textContent = storeName;

    const { results, totalSalary } = calculatePayroll(data, store.baseWage, store.overtime);
    document.getElementById('total-salary').textContent = `合計支払い給与：${totalSalary.toLocaleString()}円`;
    const tbody = document.querySelector('#employees tbody');
    results.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.name}</td><td>${store.baseWage}</td><td>${r.hours.toFixed(2)}</td><td>${r.days}</td><td>${r.salary.toLocaleString()}</td>`;
      tr.addEventListener('click', () => alert('日毎の勤務時間表示は未実装です'));
      tbody.appendChild(tr);
    });

    document.getElementById('download').addEventListener('click', () => downloadResults(storeName, `${year}${startMonth}`, store, results));
  } catch (e) {
    document.getElementById('error').textContent = 'URLが変更された可能性があります。設定からURL変更をお試しください。';
  }
});

function downloadResults(storeName, period, store, results) {
  const aoa = [['従業員名', '基本時給', '勤務時間', '出勤日数', '給与'], ...results.map(r => [r.name, store.baseWage, r.hours, r.days, r.salary])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '結果');
  const fileName = `${period}_${storeName}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
