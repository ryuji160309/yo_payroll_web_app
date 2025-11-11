const TIME_RANGE_REGEX = /^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/;
const BREAK_DEDUCTIONS = [
  { minHours: 8, deduct: 1 },
  { minHours: 7, deduct: 0.75 },
  { minHours: 6, deduct: 0.5 }
];

function calculateSalaryFromBreakdown(breakdown, baseWage, overtime) {
  if (!breakdown) return 0;
  const regular = Number(breakdown.regularHours ?? breakdown.regular ?? 0) || 0;
  const overtimeHours = Number(breakdown.overtimeHours ?? breakdown.overtime ?? 0) || 0;
  const overtimeRate = Number.isFinite(overtime) ? Number(overtime) : 1;
  const base = regular * baseWage;
  const extra = overtimeHours * baseWage * overtimeRate;
  return Math.floor(base + extra);
}

function calculateEmployee(schedule, baseWage, overtime) {
  let total = 0;
  let workdays = 0;
  let regularTotal = 0;
  let overtimeTotal = 0;
  schedule.forEach(cell => {
    if (!cell) return;
    const segments = cell.toString().split(',');
    let dayHours = 0;
    let hasValid = false;
    segments.forEach(seg => {
      const m = seg.trim().match(TIME_RANGE_REGEX);
      if (!m) return;
      const sh = parseInt(m[1], 10);
      const sm = m[2] ? parseInt(m[2], 10) : 0;
      const eh = parseInt(m[3], 10);
      const em = m[4] ? parseInt(m[4], 10) : 0;
      if (
        sh < 0 || sh > 24 || eh < 0 || eh > 24 ||
        sm < 0 || sm >= 60 || em < 0 || em >= 60
      ) return;
      hasValid = true;
      const start = sh * 60 + sm;
      const end = eh * 60 + em;
      const diff = end >= start ? end - start : 24 * 60 - start + end;
      dayHours += diff / 60;
    });
    if (!hasValid || dayHours <= 0) return;
    workdays++;
    for (const rule of BREAK_DEDUCTIONS) {
      if (dayHours >= rule.minHours) {
        dayHours -= rule.deduct;
        break;
      }
    }
    total += dayHours;
    const regular = Math.min(dayHours, 8);
    const over = Math.max(dayHours - 8, 0);
    regularTotal += regular;
    overtimeTotal += over;
  });
  const breakdown = { regularHours: regularTotal, overtimeHours: overtimeTotal };
  const salary = calculateSalaryFromBreakdown(breakdown, baseWage, overtime);
  return {
    hours: total,
    days: workdays,
    salary,
    breakdown,
    regularHours: regularTotal,
    overtimeHours: overtimeTotal
  };
}

function calculatePayroll(data, baseWage, overtime, excludeWords = []) {
  const header = data[2];
  const names = [];
  const schedules = [];
  for (let col = 3; col < header.length; col++) {
    const name = header[col];
    if (name && !excludeWords.some(word => name.includes(word))) {
      names.push(name);
      // rows 4-34 contain daily schedules
      schedules.push(data.slice(3, 34).map(row => row[col]));
    }
  }

  const results = names.map((name, idx) => {
    const r = calculateEmployee(schedules[idx], baseWage, overtime);
    const baseSalary = r.salary;
    return {
      name,
      baseWage,
      hours: r.hours,
      days: r.days,
      baseSalary,
      transport: 0,
      salary: baseSalary,
      breakdown: r.breakdown,
      regularHours: r.regularHours,
      overtimeHours: r.overtimeHours
    };
  });

  const totalSalary = results.reduce((sum, r) => sum + r.salary, 0);
  return { results, totalSalary, schedules };
}

module.exports = {
  calculateEmployee,
  calculatePayroll,
  calculateSalaryFromBreakdown,
  TIME_RANGE_REGEX,
  BREAK_DEDUCTIONS
};
