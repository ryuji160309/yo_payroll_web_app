const MINUTES_IN_DAY = 24 * 60;
const DAY_IN_MS = 24 * 60 * 60 * 1000;

function formatDateLabel(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const weekday = days[date.getDay()] || '';
  return `${year}年${month}月${day}日 (${weekday})`;
}

function normalizeDate(date) {
  const normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}

function minutesToPercent(min) {
  const clamped = Math.max(0, Math.min(min, MINUTES_IN_DAY));
  return `${(clamped / MINUTES_IN_DAY) * 100}%`;
}

function parseRangeText(text) {
  const trimmed = typeof text === 'string' ? text.trim() : '';
  if (!trimmed) return [];
  const segments = trimmed.split(',');
  const results = [];
  segments.forEach(seg => {
    const match = seg.trim().match(typeof TIME_RANGE_REGEX !== 'undefined' ? TIME_RANGE_REGEX : /^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/);
    if (!match) return;
    const startHour = parseInt(match[1], 10);
    const startMinute = match[2] ? parseInt(match[2], 10) : 0;
    const endHour = parseInt(match[3], 10);
    const endMinute = match[4] ? parseInt(match[4], 10) : 0;
    if (
      Number.isNaN(startHour) || Number.isNaN(startMinute) ||
      Number.isNaN(endHour) || Number.isNaN(endMinute) ||
      startHour < 0 || startHour > 24 || endHour < 0 || endHour > 24 ||
      startMinute < 0 || startMinute >= 60 || endMinute < 0 || endMinute >= 60
    ) {
      return;
    }
    const start = startHour * 60 + startMinute;
    let end = endHour * 60 + endMinute;
    const crossesMidnight = end <= start;
    if (crossesMidnight) {
      end += MINUTES_IN_DAY;
    }
    const startLabel = `${String(startHour).padStart(2, '0')}:${String(startMinute).padStart(2, '0')}`;
    const endLabel = `${String(endHour).padStart(2, '0')}:${String(endMinute).padStart(2, '0')}`;
    results.push({ start, end, display: `${startLabel}～${endLabel}` });
  });
  return results;
}

function overlapsWithDayRange(start, end) {
  const overlapStart = Math.max(start, 0);
  const overlapEnd = Math.min(end, MINUTES_IN_DAY);
  return overlapEnd > overlapStart ? { start: overlapStart, end: overlapEnd } : null;
}

function isSameDate(a, b) {
  return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

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

async function findWorkbookForDate(store, targetDate) {
  const sheetList = await fetchSheetList(store.url, { allowOffline: true });
  const candidates = sheetList && sheetList.length > 0 ? sheetList.slice().reverse() : [{ index: 0 }];
  for (const meta of candidates) {
    try {
      const index = typeof meta.index === 'number' ? meta.index : 0;
      const workbook = await fetchWorkbook(store.url, index, { allowOffline: true });
      const period = parseWorkbookPeriod(workbook.data || []);
      if (period && targetDate >= period.startDate && targetDate <= period.endDate) {
        return { workbook, period };
      }
    } catch (error) {
      // Try next candidate
    }
  }
  throw new Error('対象の日付を含むシートがありません。');
}

function collectShiftsForStore(store, workbook, period, targetDate) {
  const data = workbook.data || [];
  const header = data[2] || [];
  const dayIndex = Math.floor((targetDate - period.startDate) / DAY_IN_MS);
  const scheduleRows = data.slice(3, 34);
  const entries = new Map();
  const excludeWords = Array.isArray(store.excludeWords) ? store.excludeWords : [];

  const addSegment = (name, segment, label) => {
    const entry = entries.get(name) || { name, storeName: store.name, segments: [] };
    entry.segments.push({
      start: segment.start,
      end: segment.end,
      label,
      absoluteStart: segment.absoluteStart,
      absoluteEnd: segment.absoluteEnd,
    });
    entries.set(name, entry);
  };

  const processRow = (row, offset) => {
    if (!Array.isArray(row)) return;
    header.forEach((name, colIndex) => {
      if (colIndex < 3) return;
      const label = typeof name === 'string' ? name.trim() : '';
      if (!label) return;
      if (excludeWords.some(word => label.includes(word))) return;
      const cell = row[colIndex];
      const ranges = parseRangeText(cell);
      ranges.forEach(range => {
        const absoluteStart = range.start + offset * MINUTES_IN_DAY;
        const absoluteEnd = range.end + offset * MINUTES_IN_DAY;
        const overlap = overlapsWithDayRange(absoluteStart, absoluteEnd);
        if (!overlap) return;
        addSegment(label, {
          start: overlap.start,
          end: overlap.end,
          absoluteStart,
          absoluteEnd,
        }, range.display);
      });
    });
  };

  if (dayIndex >= 0 && dayIndex < scheduleRows.length) {
    processRow(scheduleRows[dayIndex], 0);
  }
  if (dayIndex - 1 >= 0 && dayIndex - 1 < scheduleRows.length) {
    processRow(scheduleRows[dayIndex - 1], -1);
  }

  const employees = Array.from(entries.values())
    .map(entry => ({
      ...entry,
      segments: entry.segments.sort((a, b) => a.start - b.start)
    }))
    .sort((a, b) => {
      const aStart = a.segments[0] ? a.segments[0].start : Number.POSITIVE_INFINITY;
      const bStart = b.segments[0] ? b.segments[0].start : Number.POSITIVE_INFINITY;
      if (aStart !== bStart) return aStart - bStart;
      return a.name.localeCompare(b.name, 'ja');
    });

  return employees;
}

function createTimeGrid(nowMinutes) {
  const grid = document.createElement('div');
  grid.className = 'time-grid';
  const fragment = document.createDocumentFragment();
  for (let hour = 0; hour <= 24; hour += 1) {
    const marker = document.createElement('div');
    marker.className = 'time-marker time-marker--major';
    const left = minutesToPercent(hour * 60);
    marker.style.left = left;
    const label = document.createElement('span');
    label.className = 'time-marker-label';
    label.textContent = `${String(hour).padStart(2, '0')}:00`;
    marker.appendChild(label);
    fragment.appendChild(marker);
    if (hour < 24) {
      for (let quarter = 1; quarter < 4; quarter += 1) {
        const minor = document.createElement('div');
        minor.className = 'time-marker time-marker--minor';
        const minute = hour * 60 + quarter * 15;
        minor.style.left = minutesToPercent(minute);
        fragment.appendChild(minor);
      }
    }
  }
  if (nowMinutes !== null && nowMinutes >= 0 && nowMinutes <= MINUTES_IN_DAY) {
    const nowLine = document.createElement('div');
    nowLine.className = 'time-marker time-marker--now';
    nowLine.style.left = minutesToPercent(nowMinutes);
    fragment.appendChild(nowLine);
  }
  grid.appendChild(fragment);
  return grid;
}

function renderTimeline(sections, selectedDate, nowMinutes) {
  const container = document.getElementById('attendance-content');
  if (!container) return;
  container.textContent = '';

  if (!Array.isArray(sections) || sections.length === 0) {
    const empty = document.createElement('p');
    empty.className = 'attendance-status';
    empty.textContent = '表示できるシフトがありません。';
    container.appendChild(empty);
    return;
  }

  const timeline = document.createElement('div');
  timeline.className = 'attendance-timeline';
  const grid = createTimeGrid(nowMinutes);
  timeline.appendChild(grid);

  sections.forEach(section => {
    const storeSection = document.createElement('section');
    storeSection.className = 'store-section';

    const header = document.createElement('div');
    header.className = 'store-section__header';
    const title = document.createElement('h2');
    title.textContent = section.storeName;
    header.appendChild(title);
    storeSection.appendChild(header);

    const body = document.createElement('div');
    body.className = 'store-section__body';

    if (!section.employees || section.employees.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'store-empty';
      empty.textContent = 'この日のシフトが見つかりません。';
      body.appendChild(empty);
    } else {
      section.employees.forEach(employee => {
        const row = document.createElement('div');
        row.className = 'shift-row';

        const badge = document.createElement('div');
        const isActive = employee.segments.some(seg => nowMinutes !== null && nowMinutes >= seg.start && nowMinutes < seg.end);
        badge.className = `shift-badge${isActive ? ' shift-badge--active' : ''}`;
        badge.textContent = employee.name;
        row.appendChild(badge);

        const track = document.createElement('div');
        track.className = 'shift-track';

        employee.segments.forEach(seg => {
          const bar = document.createElement('div');
          bar.className = 'shift-bar';
          bar.style.left = minutesToPercent(seg.start);
          bar.style.width = `calc(${minutesToPercent(seg.end)} - ${minutesToPercent(seg.start)})`;
          const label = document.createElement('span');
          label.className = 'shift-bar__label';
          label.textContent = seg.label;
          bar.appendChild(label);
          track.appendChild(bar);
        });

        row.appendChild(track);
        body.appendChild(row);
      });
    }

    storeSection.appendChild(body);
    timeline.appendChild(storeSection);
  });

  container.appendChild(timeline);
}

async function loadAttendance(targetDate) {
  const status = document.getElementById('attendance-status');
  const normalizedDate = normalizeDate(targetDate);
  if (status) {
    status.textContent = '読み込み中…';
  }
  const stores = loadStores();
  const tasks = Object.values(stores).map(store => (
    findWorkbookForDate(store, normalizedDate)
      .then(({ workbook, period }) => ({
        storeName: store.name,
        employees: collectShiftsForStore(store, workbook, period, normalizedDate)
      }))
      .catch(error => ({ error, storeName: store.name }))
  ));

  const results = await Promise.all(tasks);
  const sections = results.filter(res => !res.error);
  const errors = results.filter(res => res.error);

  const now = new Date();
  const nowMinutes = isSameDate(normalizedDate, now)
    ? now.getHours() * 60 + now.getMinutes() + now.getSeconds() / 60
    : null;

  renderTimeline(sections, normalizedDate, nowMinutes);
  if (status) {
    if (errors.length > 0) {
      const failedStores = errors.map(e => e.storeName).filter(Boolean).join('、');
      status.textContent = `${failedStores} のシートを読み込めませんでした。`;
      status.classList.add('attendance-status--error');
    } else {
      status.textContent = `${sections.length}件の店舗からシフトを読み込みました。`;
      status.classList.remove('attendance-status--error');
    }
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  await ensureSettingsLoaded();
  const dateLabel = document.getElementById('selected-date');
  const prevBtn = document.getElementById('prev-day');
  const nextBtn = document.getElementById('next-day');
  const todayBtn = document.getElementById('today-button');

  let currentDate = normalizeDate(new Date());

  const updateDateLabel = () => {
    if (dateLabel) {
      dateLabel.textContent = formatDateLabel(currentDate);
    }
  };

  const load = () => {
    updateDateLabel();
    loadAttendance(currentDate);
  };

  if (prevBtn) {
    prevBtn.addEventListener('click', () => {
      currentDate = new Date(currentDate.getTime() - DAY_IN_MS);
      load();
    });
  }

  if (nextBtn) {
    nextBtn.addEventListener('click', () => {
      currentDate = new Date(currentDate.getTime() + DAY_IN_MS);
      load();
    });
  }

  if (todayBtn) {
    todayBtn.addEventListener('click', () => {
      currentDate = normalizeDate(new Date());
      load();
    });
  }

  load();
});
