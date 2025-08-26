document.addEventListener('DOMContentLoaded', async () => {
  const params = new URLSearchParams(location.search);
  const storeKey = params.get('store');

  const sheetIndex = parseInt(params.get('sheet'), 10) || 0;
  const store = getStore(storeKey);
  if (!store) return;
  const statusEl = document.getElementById('status');
  startLoading(statusEl, '読込中・・・');
  try {
    const { data } = await fetchWorkbook(store.url, sheetIndex);
    stopLoading(statusEl);
    const year = data[1] && data[1][2];
    const startMonth = data[1] && data[1][4];
    const endMonth = ('0' + (((parseInt(startMonth, 10) || 0) % 12) + 1)).slice(-2);
    document.getElementById('period').textContent = `${year}年${startMonth}月16日～${endMonth}月15日`;
    const nameRow = data[36] || [];
    const storeName = nameRow[14] || store.name;
    document.getElementById('store-name').textContent = storeName;
    startLoading(statusEl, '計算中・・・');

    const { results, totalSalary, schedules } = calculatePayroll(data, store.baseWage, store.overtime, store.excludeWords || []);
    document.getElementById('total-salary').textContent = `合計支払い給与：${totalSalary.toLocaleString()}円`;
    const tbody = document.querySelector('#employees tbody');
    results.forEach((r, idx) => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.name}</td><td><input type="number" value="${r.baseWage}" class="wage-input" data-idx="${idx}"></td><td>${r.hours.toFixed(2)}</td><td>${r.days}</td><td class="salary-cell">${r.salary.toLocaleString()}</td>`;
      tr.addEventListener('click', () => alert('日毎の勤務時間表示は未実装です'));
      tbody.appendChild(tr);
    });
    stopLoading(statusEl);

    document.getElementById('recalc').addEventListener('click', () => {
      const inputs = document.querySelectorAll('.wage-input');
      let total = 0;
      inputs.forEach(input => {
        const idx = parseInt(input.dataset.idx, 10);
        const wage = Number(input.value);
        if (!Number.isFinite(wage)) {
          total += results[idx].salary;
          return;
        }
        const r = calculateEmployee(schedules[idx], wage, store.overtime);
        results[idx].baseWage = wage;
        results[idx].salary = r.salary;
        total += r.salary;
        input.closest('tr').querySelector('.salary-cell').textContent = r.salary.toLocaleString();
      });
      document.getElementById('total-salary').textContent = `合計支払い給与：${total.toLocaleString()}円`;
    });

    document.getElementById('download').addEventListener('click', () => downloadResults(storeName, `${year}${startMonth}`, results));
  } catch (e) {
    stopLoading(statusEl);
    document.getElementById('error').textContent = 'URLが変更された可能性があります。設定からURL変更をお試しください。';
  }
});

function downloadResults(storeName, period, results) {
  const aoa = [['従業員名', '基本時給', '勤務時間', '出勤日数', '給与'], ...results.map(r => [r.name, r.baseWage, r.hours, r.days, r.salary])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '結果');
  const fileName = `${period}_${storeName}.xlsx`;
  XLSX.writeFile(wb, fileName);
}
