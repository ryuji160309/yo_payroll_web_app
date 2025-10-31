const { calculateEmployee, calculatePayroll } = require('../calc');

describe('calculateEmployee', () => {
  test('computes hours, days, and salary including overtime', () => {
    const schedule = ['9-17', '9-18', '9-19'];
    const baseWage = 1000;
    const overtime = 1.25;
    const result = calculateEmployee(schedule, baseWage, overtime);
    expect(result).toEqual({ hours: 24, days: 3, salary: 24250 });
  });
});

describe('calculatePayroll', () => {
  test('aggregates payroll for multiple employees', () => {
    const data = Array.from({ length: 34 }, () => []);
    // Header row with employee names starting from column index 3
    data[2][3] = 'Alice';
    data[2][4] = 'Bob';
    // Alice schedules
    data[3][3] = '9-17'; // 8h -> 7h after deduction
    data[4][3] = '10-15'; // 5h
    // Bob schedules
    data[3][4] = '9-19'; // 10h -> 9h after deduction

    const baseWage = 1000;
    const overtime = 1.25;
    const { results, totalSalary } = calculatePayroll(data, baseWage, overtime);

    expect(results).toEqual([
      { name: 'Alice', baseWage, hours: 12, days: 2, baseSalary: 12000, transport: 0, salary: 12000 },
      { name: 'Bob', baseWage, hours: 9, days: 1, baseSalary: 9250, transport: 0, salary: 9250 }
    ]);
    expect(totalSalary).toBe(21250);
  });
});
