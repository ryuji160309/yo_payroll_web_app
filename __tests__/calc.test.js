const {
  calculateEmployee,
  calculatePayroll,
  calculateSalaryFromBreakdown
} = require('../calc');

describe('calculateEmployee', () => {
  test('computes hours, days, and salary including overtime', () => {
    const schedule = ['9-17', '9-18', '9-19'];
    const baseWage = 1000;
    const overtime = 1.25;
    const result = calculateEmployee(schedule, baseWage, overtime);
    expect(result).toEqual({
      hours: 24,
      days: 3,
      absentDays: 0,
      salary: 24250,
      breakdown: { regularHours: 23, overtimeHours: 1 },
      regularHours: 23,
      overtimeHours: 1
    });
  });

  test('counts absence entries without affecting hours', () => {
    const schedule = ['欠勤', '9-18'];
    const baseWage = 1000;
    const overtime = 1.25;
    const result = calculateEmployee(schedule, baseWage, overtime);
    expect(result).toEqual({
      hours: 8,
      days: 1,
      absentDays: 1,
      salary: 8000,
      breakdown: { regularHours: 8, overtimeHours: 0 },
      regularHours: 8,
      overtimeHours: 0
    });
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
      {
        name: 'Alice',
        baseWage,
        hours: 12,
        days: 2,
        absentDays: 0,
        baseSalary: 12000,
        transport: 0,
        salary: 12000,
        breakdown: { regularHours: 12, overtimeHours: 0 },
        regularHours: 12,
        overtimeHours: 0
      },
      {
        name: 'Bob',
        baseWage,
        hours: 9,
        days: 1,
        absentDays: 0,
        baseSalary: 9250,
        transport: 0,
        salary: 9250,
        breakdown: { regularHours: 8, overtimeHours: 1 },
        regularHours: 8,
        overtimeHours: 1
      }
    ]);
    expect(totalSalary).toBe(21250);
  });
});

describe('calculateSalaryFromBreakdown', () => {
  test('recomputes salary using stored hour breakdown', () => {
    const schedule = ['9-17', '9-19'];
    const baseWage = 1000;
    const overtime = 1.25;
    const result = calculateEmployee(schedule, baseWage, overtime);
    const newBaseWage = 1200;
    const recomputed = calculateSalaryFromBreakdown(result.breakdown, newBaseWage, overtime);
    const expected = Math.floor(
      result.breakdown.regularHours * newBaseWage +
      result.breakdown.overtimeHours * newBaseWage * overtime
    );
    expect(recomputed).toBe(expected);
  });
});
