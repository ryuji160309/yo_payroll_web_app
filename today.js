const DAY_IN_MS = 24 * 60 * 60 * 1000;
const FALLBACK_TIME_REGEX = /^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/;

function createEmptySlots() {
  return Array.from({ length: 24 }, () => []);
}

function hasAttendance(slots) {
  if (!Array.isArray(slots)) {
    return false;
  }
  return slots.some(hour => Array.isArray(hour) && hour.length > 0);
}

function isWithinPeriod(date, period) {
  if (!(date instanceof Date) || !period || !period.startDate || !period.endDate) {
    return false;
  }
  const normalizedDate = normalizeDate(date);
  return normalizedDate >= normalizeDate(period.startDate)
    && normalizedDate <= normalizeDate(period.endDate);
}

function normalizeDate(value) {
  const date = new Date(value);
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatDateLabel(date) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日（${days[date.getDay()]}）`;
}

function formatPeriodLabel(startDate, endDate, fallback = '') {
  const startTime = startDate instanceof Date ? startDate.getTime() : NaN;
  const endTime = endDate instanceof Date ? endDate.getTime() : NaN;
  if (Number.isNaN(startTime) || Number.isNaN(endTime)) {
    return fallback;
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

function parsePeriodFromSheet(data, fallbackLabel = '') {
  const rawYear = data?.[1]?.[2];
  const rawMonth = data?.[1]?.[4];
  const year = Number.parseInt(rawYear, 10);
  const startMonth = Number.parseInt(rawMonth, 10);
  if (!Number.isFinite(year) || !Number.isFinite(startMonth)) {
    return { startDate: null, endDate: null, label: fallbackLabel };
  }
  const startDate = new Date(year, startMonth - 1, 16);
  startDate.setHours(0, 0, 0, 0);
  const endDate = new Date(startDate.getFullYear(), startDate.getMonth() + 1, 15);
  endDate.setHours(23, 59, 59, 999);
  return { startDate, endDate, label: formatPeriodLabel(startDate, endDate, fallbackLabel) };
}

function parseTimeRanges(cellValue) {
  if (cellValue === null || cellValue === undefined) {
    return [];
  }
  const raw = String(cellValue).trim();
  if (!raw) {
    return [];
  }
  const regex = (typeof TIME_RANGE_REGEX !== 'undefined') ? TIME_RANGE_REGEX : FALLBACK_TIME_REGEX;
  const segments = raw.split(/[,、\s]+/).map(part => part.trim()).filter(Boolean);
  const ranges = [];
  segments.forEach(seg => {
    const match = seg.match(regex);
    if (!match) {
      return;
    }
    const sh = parseInt(match[1], 10);
    const sm = match[2] ? parseInt(match[2], 10) : 0;
    const eh = parseInt(match[3], 10);
    const em = match[4] ? parseInt(match[4], 10) : 0;
    if (sh < 0 || sh > 24 || eh < 0 || eh > 24 || sm < 0 || sm >= 60 || em < 0 || em >= 60) {
      return;
    }
    ranges.push({ start: sh * 60 + sm, end: eh * 60 + em });
  });
  return ranges;
}

function fillSlots(slots, startMinutes, endMinutes, name) {
  if (!Array.isArray(slots) || !name) {
    return;
  }
  const start = Number.isFinite(startMinutes) ? startMinutes : NaN;
  const end = Number.isFinite(endMinutes) ? endMinutes : NaN;
  if (Number.isNaN(start) || Number.isNaN(end)) {
    return;
  }
  const clampedStart = Math.max(0, start);
  const clampedEnd = Math.min(24 * 60, end);
  if (clampedEnd <= clampedStart) {
    return;
  }
  for (let minutes = Math.floor(clampedStart / 60) * 60; minutes < clampedEnd; minutes += 60) {
    const hourIndex = Math.floor(minutes / 60);
    const slotStart = minutes;
    const slotEnd = minutes + 60;
    const overlap = Math.min(clampedEnd, slotEnd) - Math.max(clampedStart, slotStart);
    if (overlap > 0 && slots[hourIndex]) {
      slots[hourIndex].add(name);
    }
  }
}

function buildSlotsForDay(data, startDate, targetDate, store) {
  const startTime = startDate instanceof Date ? startDate.getTime() : NaN;
  const targetTime = targetDate instanceof Date ? targetDate.getTime() : NaN;
  if (!Array.isArray(data) || Number.isNaN(startTime) || Number.isNaN(targetTime)) {
    return null;
  }
  const normalizedStart = normalizeDate(startDate);
  const normalizedTarget = normalizeDate(targetDate);
  const offset = Math.floor((normalizedTarget.getTime() - normalizedStart.getTime()) / DAY_IN_MS);
  const scheduleRowIndex = 3 + offset;
  if (offset < 0 || scheduleRowIndex < 0 || scheduleRowIndex >= data.length) {
    return null;
  }
  const row = data[scheduleRowIndex];
  const header = data[2] || [];
  if (!row || !Array.isArray(row) || !Array.isArray(header)) {
    return null;
  }
  const slots = Array.from({ length: 24 }, () => new Set());
  const excludeWords = Array.isArray(store.excludeWords) ? store.excludeWords : [];

  const appendRowToSlots = (rowData, { overnightHeadOnly = false } = {}) => {
    if (!rowData || !Array.isArray(rowData)) {
      return;
    }
    for (let col = 3; col < header.length; col += 1) {
      const name = header[col];
      if (!name || excludeWords.some(word => String(name).includes(word))) {
        continue;
      }
      const ranges = parseTimeRanges(rowData[col]);
      ranges.forEach(range => {
        const start = range.start;
        const end = range.end;
        if (overnightHeadOnly) {
          if (end <= start) {
            fillSlots(slots, 0, end, String(name));
          }
          return;
        }
        if (end <= start) {
          fillSlots(slots, start, 24 * 60, String(name));
          return;
        }
        fillSlots(slots, start, end, String(name));
      });
    }
  };

  appendRowToSlots(row);
  if (scheduleRowIndex - 1 >= 3 && offset > 0) {
    const previousRow = data[scheduleRowIndex - 1];
    appendRowToSlots(previousRow, { overnightHeadOnly: true });
  }
  return slots.map(set => Array.from(set));
}

async function resolveSheetForToday(store, targetDate) {
  const warnings = [];
  let sheetList = [];
  try {
    sheetList = await fetchSheetList(store.url, { allowOffline: true });
  } catch (error) {
    warnings.push(`${store.name}：シート一覧の取得に失敗しました`);
  }
  const candidates = Array.isArray(sheetList) && sheetList.length ? sheetList : [{ index: 0, name: '' }];
  const normalizedTarget = normalizeDate(targetDate);
  for (let i = 0; i < candidates.length; i += 1) {
    const meta = candidates[i] || {};
    const sheetIndex = Number.isFinite(meta.index) ? meta.index : i;
    try {
      const workbook = await fetchWorkbook(store.url, sheetIndex, { allowOffline: true });
      const labelSeed = workbook.sheetName || meta.name || `シート${sheetIndex + 1}`;
      const period = parsePeriodFromSheet(workbook.data, labelSeed);
      if (!period.startDate || !period.endDate) {
        continue;
      }
      if (normalizedTarget >= normalizeDate(period.startDate) && normalizedTarget <= normalizeDate(period.endDate)) {
        return { workbook, period, warnings };
      }
    } catch (error) {
      warnings.push(`${store.name}：シート${sheetIndex + 1}の読み込みに失敗しました`);
    }
  }
  return { workbook: null, period: null, warnings };
}

function renderWarnings(messages) {
  const warningBox = document.getElementById('today-warnings');
  if (!warningBox) {
    return;
  }
  const validMessages = messages.filter(Boolean);
  if (!validMessages.length) {
    warningBox.hidden = true;
    warningBox.textContent = '';
    return;
  }
  warningBox.hidden = false;
  warningBox.innerHTML = validMessages.map(msg => `<div>${msg}</div>`).join('');
}

function buildSheetViewUrl(storeUrl, workbook) {
  if (typeof extractFileId !== 'function') {
    return null;
  }
  const fileId = extractFileId(storeUrl || '');
  if (!fileId) {
    return null;
  }
  const sheetId = Number.isFinite(workbook?.sheetId) ? Number(workbook.sheetId) : null;
  const fallbackGidMatch = storeUrl ? storeUrl.match(/[#&]gid=(\d+)/) : null;
  const gid = sheetId !== null
    ? sheetId
    : (fallbackGidMatch ? Number(fallbackGidMatch[1]) : null);
  const gidFragment = Number.isFinite(gid) ? `#gid=${gid}` : '';
  return `https://docs.google.com/spreadsheets/d/${fileId}/edit${gidFragment}`;
}

function renderAttendanceTable(stores, options = {}) {
  const headerRow = document.getElementById('today-header-row');
  const body = document.getElementById('today-table-body');
  if (!headerRow || !body) {
    return;
  }
  headerRow.textContent = '';
  body.textContent = '';

  if (!stores.length) {
    const row = document.createElement('tr');
    const cell = document.createElement('td');
    cell.colSpan = Math.max(1, stores.length + 1);
    cell.textContent = '表示できる出勤予定がありません。';
    cell.className = 'today-empty';
    row.appendChild(cell);
    body.appendChild(row);
    return;
  }

  const timeHeader = document.createElement('th');
  timeHeader.textContent = '時間帯';
  timeHeader.className = 'today-time-cell today-time-header';
  headerRow.appendChild(timeHeader);

  stores.forEach(store => {
    const sheetUrl = buildSheetViewUrl(store.sourceStore?.url, store.workbook);
    const th = document.createElement('th');
    th.className = 'today-store-header';
    if (sheetUrl) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'today-store-link';
      button.textContent = store.storeName || '';
      button.addEventListener('click', () => {
        window.open(sheetUrl, '_blank', 'noopener');
      });
      button.title = `${store.storeName || '店舗'}のシートを開く`;
      th.appendChild(button);
    } else {
      th.textContent = store.storeName || '';
    }
    headerRow.appendChild(th);
  });

  for (let hour = 0; hour < 24; hour += 1) {
    const row = document.createElement('tr');
    row.dataset.hour = String(hour);
    const timeCell = document.createElement('th');
    timeCell.className = 'today-time-cell';
    timeCell.scope = 'row';
    timeCell.textContent = `${hour}時`;
    if (options.currentHour === hour) {
      row.classList.add('today-current-hour-row');
      timeCell.classList.add('today-current-hour-cell');
    }
    row.appendChild(timeCell);

    stores.forEach(store => {
      const cell = document.createElement('td');
      const badges = store.slots && store.slots[hour] ? store.slots[hour] : [];
      if (badges.length === 0) {
        const empty = document.createElement('span');
        empty.className = 'today-empty';
        empty.textContent = '―';
        cell.appendChild(empty);
      } else {
        const list = document.createElement('div');
        list.className = 'today-badges';
        badges.forEach(name => {
          const badge = document.createElement('span');
          badge.className = 'today-badge';
          badge.textContent = name;
          list.appendChild(badge);
        });
        cell.appendChild(list);
      }
      if (options.currentHour === hour) {
        cell.classList.add('today-current-hour-slot');
      }
      row.appendChild(cell);
    });
    body.appendChild(row);
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('today-status');
  const dateEl = document.getElementById('today-date');
  const periodEl = document.getElementById('today-period-note');
  const prevButton = document.getElementById('today-prev-day');
  const nextButton = document.getElementById('today-next-day');
  const today = normalizeDate(new Date());
  const currentHour = new Date().getHours();

  let displayDate = today;
  let rangeStart = null;
  let rangeEnd = null;
  let initialScrollDone = false;
  const warnings = [];
  const loadedStores = [];

  if (periodEl) {
    periodEl.textContent = '全店舗のシートから当日の出勤予定を読み込みます。';
  }
  if (statusEl) {
    startLoading(statusEl, '今日の出勤者を読み込み中…', { disableSlowNote: true });
  }

  const updateDateLabel = date => {
    if (dateEl) {
      dateEl.textContent = formatDateLabel(date);
    }
  };

  const updatePeriodNote = storesForDate => {
    if (!periodEl) {
      return;
    }
    const labels = storesForDate
      .filter(store => store.periodActive && store.periodLabel)
      .map(store => store.periodLabel);
    const fallbackLabels = loadedStores
      .map(store => store.periodLabel)
      .filter(label => !!label);
    const uniqueLabels = Array.from(new Set(labels.length ? labels : fallbackLabels));
    periodEl.textContent = uniqueLabels.length
      ? `対象期間：${uniqueLabels.join(' ／ ')}`
      : '各店舗の最新シートから出勤予定を表示しています。';
  };

  const updateStatus = ({ hasPeriodData, hasAttendance }) => {
    if (!statusEl) {
      return;
    }
    if (!loadedStores.length) {
      statusEl.textContent = '出勤予定を取得できませんでした。';
      return;
    }
    if (!hasPeriodData) {
      statusEl.textContent = '選択した日は対象期間外です。';
      return;
    }
    if (!hasAttendance) {
      statusEl.textContent = '選択した日の出勤予定はありません。';
      return;
    }
    statusEl.textContent = '';
  };

  const scrollToHour = hour => {
    if (typeof hour !== 'number' || Number.isNaN(hour)) {
      return false;
    }
    const wrapper = document.querySelector('.today-table-wrapper');
    if (!wrapper) {
      return false;
    }
    const targetRow = wrapper.querySelector(`tr[data-hour="${hour}"]`);
    if (!targetRow) {
      return false;
    }
    requestAnimationFrame(() => {
      const offset = targetRow.offsetTop - wrapper.offsetTop;
      const scrollTarget = Math.max(0, offset - Math.max(0, (wrapper.clientHeight - targetRow.clientHeight) / 2));
      wrapper.scrollTop = scrollTarget;
    });
    return true;
  };

  const updateNavButtons = () => {
    const normalizedDate = normalizeDate(displayDate);
    if (prevButton) {
      prevButton.disabled = !rangeStart || normalizedDate <= normalizeDate(rangeStart);
    }
    if (nextButton) {
      nextButton.disabled = !rangeEnd || normalizedDate >= normalizeDate(rangeEnd);
    }
  };

  const renderForDate = (targetDate, { shouldScrollToCurrent = false } = {}) => {
    const normalizedDate = normalizeDate(targetDate);
    displayDate = normalizedDate;
    updateDateLabel(normalizedDate);

    const storesForDate = loadedStores.map(store => {
      const periodActive = isWithinPeriod(normalizedDate, store.period);
      let slots = null;
      if (periodActive && store.workbook?.data) {
        slots = buildSlotsForDay(store.workbook.data, store.period.startDate, normalizedDate, store.sourceStore)
          || createEmptySlots();
      }
      return {
        ...store,
        slots,
        periodActive
      };
    });

    const storesWithinPeriod = storesForDate.filter(store => store.periodActive);
    const storesWithAttendance = storesWithinPeriod.filter(store => hasAttendance(store.slots));

    renderAttendanceTable(storesWithAttendance, {
      currentHour: normalizedDate.getTime() === today.getTime() ? currentHour : null
    });
    renderWarnings(warnings);
    updatePeriodNote(storesForDate);
    updateStatus({
      hasPeriodData: storesWithinPeriod.length > 0,
      hasAttendance: storesWithAttendance.length > 0
    });
    updateNavButtons();

    if (!initialScrollDone && shouldScrollToCurrent && normalizedDate.getTime() === today.getTime()) {
      const scrolled = scrollToHour(currentHour);
      if (scrolled) {
        initialScrollDone = true;
      } else {
        setTimeout(() => {
          if (!initialScrollDone) {
            initialScrollDone = scrollToHour(currentHour);
          }
        }, 200);
      }
    }
  };

  const clampDate = date => {
    let nextDate = date instanceof Date ? new Date(date.getTime()) : new Date();
    nextDate = normalizeDate(nextDate);
    if (rangeStart && nextDate < normalizeDate(rangeStart)) {
      return normalizeDate(rangeStart);
    }
    if (rangeEnd && nextDate > normalizeDate(rangeEnd)) {
      return normalizeDate(rangeEnd);
    }
    return nextDate;
  };

  const shiftDate = delta => {
    const nextDate = clampDate(new Date(displayDate.getTime() + delta * DAY_IN_MS));
    if (nextDate.getTime() === displayDate.getTime()) {
      return;
    }
    renderForDate(nextDate, { shouldScrollToCurrent: true });
  };

  try {
    if (window.settingsLoadPromise && typeof window.settingsLoadPromise.then === 'function') {
      await window.settingsLoadPromise;
    }
  } catch (error) {
    // Keep going with cached settings if available.
  }

  const stores = DEFAULT_STORES || {};
  const storeEntries = Object.keys(stores)
    .map(key => ({ key, ...stores[key] }))
    .sort((a, b) => (a.name || '').localeCompare(b.name || '', 'ja'));

  for (let i = 0; i < storeEntries.length; i += 1) {
    const store = storeEntries[i];
    if (!store || !store.url) {
      continue;
    }
    const { workbook, period, warnings: storeWarnings } = await resolveSheetForToday(store, today);
    if (Array.isArray(storeWarnings)) {
      warnings.push(...storeWarnings);
    }
    if (!workbook || !period) {
      warnings.push(`${store.name}：今日を含むシートが見つかりません`);
    } else {
      if (period.startDate instanceof Date) {
        rangeStart = rangeStart && rangeStart instanceof Date
          ? new Date(Math.min(rangeStart.getTime(), period.startDate.getTime()))
          : new Date(period.startDate.getTime());
      }
      if (period.endDate instanceof Date) {
        rangeEnd = rangeEnd && rangeEnd instanceof Date
          ? new Date(Math.max(rangeEnd.getTime(), period.endDate.getTime()))
          : new Date(period.endDate.getTime());
      }
    }
    loadedStores.push({
      storeKey: store.key,
      storeName: store.name,
      workbook: workbook || null,
      period: period || null,
      periodLabel: period?.label || '',
      sourceStore: store
    });
  }

  if (statusEl) {
    stopLoading(statusEl);
  }

  renderForDate(displayDate, { shouldScrollToCurrent: true });

  if (prevButton) {
    prevButton.addEventListener('click', () => shiftDate(-1));
  }
  if (nextButton) {
    nextButton.addEventListener('click', () => shiftDate(1));
  }

  initializeHelp('help/today.txt', {
    pageKey: 'today',
    steps: {
      date: '#today-date',
      prev: '#today-prev-day',
      next: '#today-next-day',
      table: '.today-table-wrapper',
      warnings: '#today-warnings'
    },
    waitForReady: () => document.getElementById('today-table-body')?.children.length > 0
  });
});
