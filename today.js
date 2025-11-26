const DAY_IN_MS = 24 * 60 * 60 * 1000;
const FALLBACK_TIME_REGEX = /^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/;

function createEmptySlots() {
  return Array.from({ length: 24 }, () => []);
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
  let start = startMinutes;
  let end = endMinutes;
  if (!Number.isFinite(start) || !Number.isFinite(end)) {
    return;
  }
  if (end <= start) {
    end += 24 * 60;
  }
  for (let minutes = Math.floor(start / 60) * 60; minutes < end; minutes += 60) {
    const hourIndex = Math.floor(minutes / 60) % 24;
    const slotStart = minutes;
    const slotEnd = minutes + 60;
    const overlap = Math.min(end, slotEnd) - Math.max(start, slotStart);
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
  for (let col = 3; col < header.length; col += 1) {
    const name = header[col];
    if (!name || excludeWords.some(word => String(name).includes(word))) {
      continue;
    }
    const ranges = parseTimeRanges(row[col]);
    ranges.forEach(range => fillSlots(slots, range.start, range.end, String(name)));
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

function renderAttendanceTable(stores) {
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
  headerRow.appendChild(timeHeader);

  stores.forEach(store => {
    const th = document.createElement('th');
    th.textContent = store.storeName || '';
    if (store.periodLabel) {
      const note = document.createElement('span');
      note.className = 'today-header-note';
      note.textContent = store.periodLabel;
      th.appendChild(document.createElement('br'));
      th.appendChild(note);
    }
    headerRow.appendChild(th);
  });

  for (let hour = 0; hour < 24; hour += 1) {
    const row = document.createElement('tr');
    const timeCell = document.createElement('th');
    timeCell.className = 'today-time-cell';
    timeCell.scope = 'row';
    timeCell.textContent = `${hour}時`;
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
      row.appendChild(cell);
    });
    body.appendChild(row);
  }
}

document.addEventListener('DOMContentLoaded', async () => {
  const statusEl = document.getElementById('today-status');
  const dateEl = document.getElementById('today-date');
  const periodEl = document.getElementById('today-period-note');
  const targetDate = normalizeDate(new Date());

  if (dateEl) {
    dateEl.textContent = formatDateLabel(targetDate);
  }
  if (periodEl) {
    periodEl.textContent = '全店舗のシートから当日の出勤予定を読み込みます。';
  }
  if (statusEl) {
    startLoading(statusEl, '今日の出勤者を読み込み中…', { disableSlowNote: true });
  }

  try {
    if (window.settingsLoadPromise && typeof window.settingsLoadPromise.then === 'function') {
      await window.settingsLoadPromise;
    }
  } catch (error) {
    // Keep going with cached settings if available.
  }

  const stores = DEFAULT_STORES || {};
  const storeEntries = Object.keys(stores).map(key => ({ key, ...stores[key] }));
  const warnings = [];
  const tableStores = [];

  for (let i = 0; i < storeEntries.length; i += 1) {
    const store = storeEntries[i];
    if (!store || !store.url) {
      continue;
    }
    const { workbook, period, warnings: storeWarnings } = await resolveSheetForToday(store, targetDate);
    if (Array.isArray(storeWarnings)) {
      warnings.push(...storeWarnings);
    }
    if (!workbook || !period) {
      warnings.push(`${store.name}：今日を含むシートが見つかりません`);
      tableStores.push({
        storeKey: store.key,
        storeName: store.name,
        slots: createEmptySlots(),
        periodLabel: ''
      });
      continue;
    }
    const slots = buildSlotsForDay(workbook.data, period.startDate, targetDate, store);
    if (!slots) {
      warnings.push(`${store.name}：当日の行を取得できませんでした`);
      tableStores.push({
        storeKey: store.key,
        storeName: store.name,
        slots: createEmptySlots(),
        periodLabel: period.label || ''
      });
      continue;
    }
    tableStores.push({
      storeKey: store.key,
      storeName: store.name,
      slots,
      periodLabel: period.label || ''
    });
  }

  tableStores.sort((a, b) => (a.storeName || '').localeCompare(b.storeName || '', 'ja'));

  if (statusEl) {
    stopLoading(statusEl);
    statusEl.textContent = tableStores.length ? '' : '出勤予定を取得できませんでした。';
  }

  renderAttendanceTable(tableStores);
  renderWarnings(warnings);

  if (periodEl && tableStores.length) {
    const matchedSheets = tableStores
      .map(store => store.periodLabel)
      .filter(label => !!label);
    const uniqueLabels = Array.from(new Set(matchedSheets));
    periodEl.textContent = uniqueLabels.length
      ? `対象期間：${uniqueLabels.join(' ／ ')}`
      : '各店舗の最新シートから出勤予定を表示しています。';
  }
});
