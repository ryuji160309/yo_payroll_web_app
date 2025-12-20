const MINUTES_IN_DAY = 24 * 60;
const DAY_IN_MS = 24 * 60 * 60 * 1000;
const SHIFT_LANE_WIDTH = 54;
const SHIFT_LANE_GAP = 8;
const TODAY_WARNING_ACK_KEY = 'todayWarningAcknowledgedAt';
const TODAY_WARNING_INTERVAL_MS = 7 * DAY_IN_MS;
const NON_CALCULABLE_EXCLUDE_WORDS = ['✕', 'x', 'X', '×', 'ｘ', '✕‬']; // 追加したい除外ワードがあればここへ
const UNPARSED_CELL_EXCLUDE_WORDS = ['✕', 'x', 'X', '×', 'ｘ', '✕‬'];
const UNPARSED_TITLE_HEIGHT = 20;
const UNPARSED_LINE_HEIGHT = 22;
const UNPARSED_SECTION_PADDING = 8;

function createDeferred() {
  let resolved = false;
  let resolveFn = null;
  const promise = new Promise(resolve => {
    resolveFn = () => {
      if (resolved) return;
      resolved = true;
      resolve();
    };
  });
  return {
    resolve: () => {
      if (typeof resolveFn === 'function') {
        resolveFn();
      }
    },
    isResolved: () => resolved,
    promise
  };
}

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

function computeUnparsedDisplayHeight(lineCount) {
  if (!lineCount) return 0;
  return (UNPARSED_SECTION_PADDING * 2) + UNPARSED_TITLE_HEIGHT + (lineCount * UNPARSED_LINE_HEIGHT);
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

function daysInMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
}

function getScheduleRowIndex(targetDate, period) {
  const normalizedDate = normalizeDate(targetDate);
  const rawIndex = Math.floor((normalizedDate - period.startDate) / DAY_IN_MS);

  const startMonthLength = daysInMonth(period.startDate);
  const missingDays = Math.max(0, 31 - startMonthLength);
  const nextMonthStart = new Date(period.startDate.getFullYear(), period.startDate.getMonth() + 1, 1);

  if (normalizedDate >= nextMonthStart) {
    return rawIndex + missingDays;
  }

  return rawIndex;
}

function collectShiftsForStore(store, workbook, period, targetDate) {
  const data = workbook.data || [];
  const header = data[2] || [];
  const dayIndex = getScheduleRowIndex(targetDate, period);
  const scheduleRows = data.slice(3, 34);
  const entries = new Map();
  const excludeWords = [
    ...NON_CALCULABLE_EXCLUDE_WORDS,
    ...(Array.isArray(store.excludeWords) ? store.excludeWords : [])
  ];
  const unparsedCells = [];

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

  const addUnparsedCell = (name, raw) => {
    const value = typeof raw === 'string' ? raw.trim() : String(raw || '').trim();
    if (!value) return;
    const shouldExclude = UNPARSED_CELL_EXCLUDE_WORDS.some(word => value.includes(word));
    if (shouldExclude) return;
    unparsedCells.push({ name, value });
  };

  const processRow = (row, offset) => {
    if (!Array.isArray(row)) return;
    header.forEach((name, colIndex) => {
      if (colIndex < 3) return;
      const label = typeof name === 'string' ? name.trim() : '';
      if (!label) return;
      if (excludeWords.some(word => label.includes(word))) return;
      const cell = row[colIndex];
      const cellText = typeof cell === 'string' || typeof cell === 'number'
        ? String(cell).trim()
        : '';
      if (!cellText) return;
      const ranges = parseRangeText(cellText);
      if (!ranges.length) {
        addUnparsedCell(label, cellText);
        return;
      }
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

  const previousDayIndex = getScheduleRowIndex(new Date(targetDate.getTime() - DAY_IN_MS), period);

  if (dayIndex >= 0 && dayIndex < scheduleRows.length) {
    processRow(scheduleRows[dayIndex], 0);
  }
  if (previousDayIndex >= 0 && previousDayIndex < scheduleRows.length) {
    processRow(scheduleRows[previousDayIndex], -1);
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

  return { employees, unparsedCells };
}

function appendTimeGrid(container, includeLabels, nowMinutes) {
  const fragment = document.createDocumentFragment();
  for (let hour = 0; hour <= 24; hour += 1) {
    const major = document.createElement('div');
    major.className = 'time-row-line time-row-line--major';
    major.style.top = minutesToPercent(hour * 60);
    fragment.appendChild(major);

    if (includeLabels && hour < 24) {
      const label = document.createElement('div');
      label.className = 'time-label';
      label.style.top = minutesToPercent(hour * 60);
      label.textContent = `${String(hour).padStart(2, '0')}:00`;
      fragment.appendChild(label);
    }

    if (hour < 24) {
      const minute = hour * 60 + 30;
      const minor = document.createElement('div');
      minor.className = 'time-row-line time-row-line--minor';
      minor.style.top = minutesToPercent(minute);
      fragment.appendChild(minor);
    }
  }

  if (nowMinutes !== null && nowMinutes >= 0 && nowMinutes <= MINUTES_IN_DAY) {
    const nowLine = document.createElement('div');
    nowLine.className = 'time-row-line time-row-line--now';
    nowLine.style.top = minutesToPercent(nowMinutes);
    fragment.appendChild(nowLine);
  }

  container.appendChild(fragment);
}

function layoutSegments(employees) {
  const items = [];
  employees.forEach(employee => {
    employee.segments.forEach(segment => {
      items.push({ employee, segment });
    });
  });

  items.sort((a, b) => {
    if (a.segment.start !== b.segment.start) return a.segment.start - b.segment.start;
    return a.segment.end - b.segment.end;
  });

  const laneEnds = [];
  items.forEach(item => {
    let laneIndex = laneEnds.findIndex(end => end <= item.segment.start);
    if (laneIndex === -1) {
      laneIndex = laneEnds.length;
      laneEnds.push(item.segment.end);
    } else {
      laneEnds[laneIndex] = item.segment.end;
    }
    item.laneIndex = laneIndex;
  });

  return { items, laneCount: laneEnds.length };
}

function renderTimeline(sections, selectedDate, nowMinutes, { alignUnparsedHeight = false } = {}) {
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

  const grid = document.createElement('div');
  grid.className = 'attendance-grid';
  grid.style.setProperty('--store-count', sections.length);

  const unparsedSections = [];
  const layoutInfo = sections.map(section => {
    const layout = section.employees && section.employees.length > 0
      ? layoutSegments(section.employees)
      : { items: [], laneCount: 0 };
    const laneSpace = layout.laneCount > 0
      ? (layout.laneCount * SHIFT_LANE_WIDTH) + ((layout.laneCount - 1) * SHIFT_LANE_GAP)
      : 0;
    const requiredWidth = Math.max(80, laneSpace + 24);
    return { section, layout, requiredWidth };
  });

  const maxRequiredWidth = layoutInfo.reduce((acc, info) => Math.max(acc, info.requiredWidth), 200);
  grid.style.setProperty('--store-min-width', `${maxRequiredWidth}px`);

  const hasUnparsed = sections.some(sec => (sec.unparsedCells || []).length > 0);
  const shouldNormalizeUnparsed = alignUnparsedHeight || hasUnparsed;
  const maxUnparsedLines = shouldNormalizeUnparsed
    ? Math.max(0, ...sections.map(sec => (sec.unparsedCells || []).length))
    : 0;
  const normalizedUnparsedHeight = shouldNormalizeUnparsed
    ? computeUnparsedDisplayHeight(maxUnparsedLines)
    : null;

  const timeColumn = document.createElement('div');
  timeColumn.className = 'time-column';

  const timeHeader = document.createElement('div');
  timeHeader.className = 'time-column__header';
  timeHeader.textContent = '時間';
  timeColumn.appendChild(timeHeader);

  let spacer = null;
  if (normalizedUnparsedHeight !== null) {
    spacer = document.createElement('div');
    spacer.className = 'time-column__spacer';
    spacer.style.minHeight = `${normalizedUnparsedHeight}px`;
    timeColumn.appendChild(spacer);
  }

  const timeBody = document.createElement('div');
  timeBody.className = 'time-column__body';
  appendTimeGrid(timeBody, true, nowMinutes);
  timeColumn.appendChild(timeBody);
  grid.appendChild(timeColumn);

  layoutInfo.forEach(info => {
    const { section, layout } = info;
    const storeColumn = document.createElement('div');
    storeColumn.className = 'store-column';

    const header = document.createElement('div');
    header.className = 'store-column__header';
    const title = document.createElement(section.storeUrl ? 'a' : 'h2');
    title.textContent = section.storeName;
    if (section.storeUrl) {
      title.href = section.storeUrl;
      title.target = '_blank';
      title.rel = 'noreferrer noopener';
      title.className = 'store-column__link';
    }
    header.appendChild(title);
    storeColumn.appendChild(header);

    const hasUnparsedCells = section.unparsedCells && section.unparsedCells.length > 0;
    const shouldRenderUnparsed = hasUnparsedCells || normalizedUnparsedHeight !== null;

    if (shouldRenderUnparsed) {
      const unparsed = document.createElement('div');
      unparsed.className = 'store-column__unparsed';
      unparsed.style.setProperty('--unparsed-section-padding', `${UNPARSED_SECTION_PADDING}px`);

      if (normalizedUnparsedHeight !== null) {
        unparsed.style.minHeight = `${normalizedUnparsedHeight}px`;
      }

      if (hasUnparsedCells) {
        const unparsedTitle = document.createElement('div');
        unparsedTitle.className = 'store-column__unparsed-title';
        unparsedTitle.textContent = '反映されていない項目';
        unparsed.appendChild(unparsedTitle);

        section.unparsedCells.forEach(entry => {
          const line = document.createElement('div');
          line.className = 'store-column__unparsed-line';
          line.textContent = `${entry.name}：${entry.value}`;
          unparsed.appendChild(line);
        });
      }

      storeColumn.appendChild(unparsed);
      unparsedSections.push(unparsed);
    }

    const body = document.createElement('div');
    body.className = 'store-column__body';
    appendTimeGrid(body, false, nowMinutes);

    if (!section.employees || section.employees.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'store-empty';
      empty.textContent = 'この日のシフトが見つかりません。';
      body.appendChild(empty);
    } else {
      const { items: laidOut, laneCount } = layout;
      const laneSpace = laneCount > 0
        ? (laneCount * SHIFT_LANE_WIDTH) + ((laneCount - 1) * SHIFT_LANE_GAP)
        : 0;
      const requiredWidth = Math.max(80, laneSpace + 24);
      storeColumn.style.minWidth = `${requiredWidth}px`;
      storeColumn.style.setProperty('--shift-lane-width', `${SHIFT_LANE_WIDTH}px`);
      storeColumn.style.setProperty('--shift-lane-gap', `${SHIFT_LANE_GAP}px`);

      laidOut.forEach(item => {
        const { employee, segment, laneIndex } = item;
        const block = document.createElement('div');
        const isActive = nowMinutes !== null && nowMinutes >= segment.start && nowMinutes < segment.end;
        block.className = `shift-block${isActive ? ' shift-block--active' : ''}`;
        block.style.top = minutesToPercent(segment.start);
        block.style.height = `calc(${minutesToPercent(segment.end)} - ${minutesToPercent(segment.start)})`;
        block.style.setProperty('--lane-index', laneIndex);

        const name = document.createElement('div');
        name.className = 'shift-block__name';
        name.textContent = employee.name;
        block.appendChild(name);

        body.appendChild(block);
      });
    }

    storeColumn.appendChild(body);
    grid.appendChild(storeColumn);
  });

  container.appendChild(grid);

  if (unparsedSections.length > 0 && spacer) {
    requestAnimationFrame(() => {
      const maxUnparsedHeight = Math.max(...unparsedSections.map(section => section.offsetHeight));
      spacer.style.minHeight = `${maxUnparsedHeight}px`;
      unparsedSections.forEach(section => {
        section.style.minHeight = `${maxUnparsedHeight}px`;
      });
    });
  }
}

async function loadStoreWorkbooks(store) {
  const result = { store, storeName: store.name, workbooks: [] };
  try {
    const sheetList = await fetchSheetList(store.url, { allowOffline: true });
    const candidates = sheetList && sheetList.length > 0 ? sheetList.slice().reverse() : [{ index: 0 }];
    for (const meta of candidates) {
      try {
        const index = typeof meta.index === 'number' ? meta.index : 0;
        const workbook = await fetchWorkbook(store.url, index, { allowOffline: true });
        const period = parseWorkbookPeriod(workbook.data || []);
        if (period) {
          result.workbooks.push({ period, workbook });
        }
      } catch (innerError) {
        // Continue trying other workbooks.
        if (!result.error) {
          result.error = innerError;
        }
      }
    }

    if (result.workbooks.length === 0 && !result.error) {
      result.error = new Error('対象期間のシートが見つかりませんでした。');
    }
  } catch (error) {
    result.error = error;
  }

  return result;
}

function buildSectionsForDate(targetDate, storeDataList) {
  const sections = [];
  const errors = [];
  const normalizedDate = normalizeDate(targetDate);

  storeDataList.forEach(entry => {
    if (entry.error) {
      errors.push(`${entry.storeName} のシートを読み込めませんでした。`);
      return;
    }
    const matched = entry.workbooks.find(w => normalizedDate >= w.period.startDate && normalizedDate <= w.period.endDate);
    if (!matched) {
      errors.push(`${entry.storeName} の対象日が含まれるシートがありませんでした。`);
      return;
    }
    try {
      const { employees, unparsedCells } = collectShiftsForStore(entry.store, matched.workbook, matched.period, normalizedDate);
      sections.push({
        storeName: entry.storeName,
        storeUrl: entry.store && entry.store.url,
        employees,
        unparsedCells
      });
    } catch (error) {
      errors.push(`${entry.storeName} のシフトを表示できませんでした。`);
    }
  });

  return { sections, errors };
}

document.addEventListener('DOMContentLoaded', async () => {
  await ensureSettingsLoaded();
  const dateLabel = document.getElementById('selected-date');
  const prevBtn = document.getElementById('prev-day');
  const nextBtn = document.getElementById('next-day');
  const todayBtn = document.getElementById('today-button');
  const status = document.getElementById('attendance-status');
  const content = document.getElementById('attendance-content');
  const storeList = document.getElementById('today-store-list');
  const scheduleSection = document.getElementById('attendance-schedule-section');
  const backButton = document.getElementById('today-back');
  const homeButton = document.getElementById('today-home');
  const dateControls = document.querySelector('.attendance-controls');

  const shouldShowNavigationButtons = () => {
    try {
      if (!document.referrer) {
        return false;
      }
      const referrerUrl = new URL(document.referrer);
      return referrerUrl.origin === window.location.origin;
    } catch (error) {
      return false;
    }
  };

  const applyNavigationVisibility = () => {
    const visible = shouldShowNavigationButtons();
    [backButton, homeButton].forEach(btn => {
      if (!btn) return;
      btn.style.display = visible ? '' : 'none';
      btn.setAttribute('aria-hidden', visible ? 'false' : 'true');
    });
  };

  const shouldShowWarningOverlay = () => {
    try {
      const raw = localStorage.getItem(TODAY_WARNING_ACK_KEY);
      const lastShown = Number(raw);
      if (!Number.isFinite(lastShown)) {
        return true;
      }
      return Date.now() - lastShown >= TODAY_WARNING_INTERVAL_MS;
    } catch (error) {
      return true;
    }
  };

  const recordWarningAcknowledged = () => {
    try {
      localStorage.setItem(TODAY_WARNING_ACK_KEY, String(Date.now()));
    } catch (error) {
      // Ignore failures
    }
  };

  const createWarningOverlay = () => {
    const overlay = document.createElement('div');
    overlay.className = 'today-warning-overlay';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');

    const contentWrapper = document.createElement('div');
    contentWrapper.className = 'today-warning-overlay__content';

    const title = document.createElement('p');
    title.className = 'today-warning-overlay__title';
    title.textContent = '⚠️必ず確認してください⚠️';
    contentWrapper.appendChild(title);

    const list = document.createElement('ul');
    list.className = 'today-warning-overlay__list';
    [
      'シフト表が正しく入力されていない',
      '同じ時刻に出勤するためにシフト表に入力されていない',
      '他店舗からのヘルプや派遣の入力'
    ].forEach(text => {
      const li = document.createElement('li');
      li.textContent = `・${text}`;
      list.appendChild(li);
    });
    contentWrapper.appendChild(list);

    const note = document.createElement('p');
    note.className = 'today-warning-overlay__note';
    note.textContent = '以上の場合はシフト予定に表示されません。';
    contentWrapper.appendChild(note);

    const reminder = document.createElement('p');
    reminder.className = 'today-warning-overlay__note';
    reminder.innerHTML = 'このページは<strong>あくまで</strong>簡単に当日のシフト者を確認するために作っているため、正確なシフトはスプレッドシートを確認してください。';
    contentWrapper.appendChild(reminder);

    const responsibility = document.createElement('p');
    responsibility.className = 'today-warning-overlay__note';
    responsibility.innerHTML = 'このページを確認して間違った時間に出勤してしまった場合/出勤を忘れた場合の<strong>責任は一切取れません</strong>。<strong>絶対</strong>に元のスプレッドシートを確認してください。';
    contentWrapper.appendChild(responsibility);

    const testNotice = document.createElement('p');
    testNotice.className = 'today-warning-overlay__note';
    testNotice.innerHTML = '表上の店舗名部分をタップすると<strong>元のスプレッドシートが開きます</strong>。';
    contentWrapper.appendChild(testNotice);

    const actions = document.createElement('div');
    actions.className = 'today-warning-overlay__actions';
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.className = 'today-warning-overlay__button';
    closeBtn.textContent = '確認';
    closeBtn.addEventListener('click', () => {
      recordWarningAcknowledged();
      overlay.remove();
    });
    actions.appendChild(closeBtn);
    contentWrapper.appendChild(actions);

    overlay.appendChild(contentWrapper);
    return overlay;
  };

  const maybeShowWarningOverlay = () => {
    if (!shouldShowWarningOverlay()) {
      return;
    }
    const overlay = createWarningOverlay();
    document.body.appendChild(overlay);
  };

  applyNavigationVisibility();
  maybeShowWarningOverlay();

  const tutorialReady = createDeferred();

  const searchParams = new URLSearchParams(window.location.search);
  const requestedStoreKey = searchParams.get('store');

  let stores = {};
  let storeKeys = [];
  let currentDate = normalizeDate(new Date());
  let storeDataList = [];
  const selectedStoreKeys = [];
  let liveNowTimer = null;
  let shouldAutoScrollToControls = false;

  const setNavigationDisabled = disabled => {
    [prevBtn, nextBtn, todayBtn].forEach(btn => {
      if (btn) {
        btn.disabled = disabled;
      }
    });
  };

  const setScheduleVisible = visible => {
    if (scheduleSection) {
      scheduleSection.style.display = visible ? '' : 'none';
    }
  };

  const updateDateLabel = () => {
    if (dateLabel) {
      dateLabel.textContent = formatDateLabel(currentDate);
    }
  };

  const getSelectionLabel = () => {
    if (storeKeys.length === selectedStoreKeys.length) {
      return '全店舗';
    }
    const names = selectedStoreKeys
      .map(key => stores[key])
      .filter(Boolean)
      .map(entry => entry.name);
    return names.length > 0 ? names.join(' / ') : '選択した店舗';
  };

  const renderForDate = () => {
    if (selectedStoreKeys.length === 0) return;
    const { sections, errors } = buildSectionsForDate(currentDate, storeDataList);
    const now = new Date();
    const nowMinutes = isSameDate(currentDate, now)
      ? now.getHours() * 60 + now.getMinutes() + now.getSeconds() / 60
      : null;

    renderTimeline(sections, currentDate, nowMinutes, {
      alignUnparsedHeight: selectedStoreKeys.length === storeKeys.length
    });

    if (status) {
      if (errors.length > 0) {
        status.textContent = errors.join('\n');
        status.classList.add('attendance-status--error');
      } else {
        const selectionLabel = getSelectionLabel();
        status.textContent = `${selectionLabel}のシフトを読み込みました。`;
        status.classList.remove('attendance-status--error');
      }
    }

    if (typeof window.requestTutorialReposition === 'function') {
      window.requestTutorialReposition();
    }

    if (shouldAutoScrollToControls && sections.length > 0) {
      shouldAutoScrollToControls = false;
      if (dateControls) {
        dateControls.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }
  };

  const renderForDateWithLiveSync = () => {
    renderForDate();
    startLiveNowTimer();
  };

  const renderStoreButtons = () => {
    if (!storeList) return;

    storeList.textContent = '';

    const addButton = (label, key) => {
      if (!label || !key) return;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'store-button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        const params = new URLSearchParams();
        params.set('store', key);
        window.location.href = `today.html?${params.toString()}`;
      });
      storeList.appendChild(btn);
    };

    addButton('全店舗', 'all');
    storeKeys.forEach(key => {
      const store = stores[key];
      addButton(store.name, key);
    });
  };

  const initializeData = async () => {
    if (selectedStoreKeys.length === 0) {
      tutorialReady.resolve();
      return;
    }

    if (content) {
      startLoading(content, '読み込み中です…');
    }
    setNavigationDisabled(true);

    try {
      const targetStores = selectedStoreKeys.map(key => stores[key]).filter(Boolean);
      const tasks = targetStores.map(store => loadStoreWorkbooks(store));
      storeDataList = await Promise.all(tasks);
    } finally {
      if (content) {
        stopLoading(content);
      }
      setNavigationDisabled(false);
      shouldAutoScrollToControls = true;
      renderForDateWithLiveSync();
      tutorialReady.resolve();
    }
  };

  const stopLiveNowTimer = () => {
    if (liveNowTimer) {
      clearInterval(liveNowTimer);
      liveNowTimer = null;
    }
  };

  const startLiveNowTimer = () => {
    stopLiveNowTimer();
    const now = new Date();
    if (!isSameDate(currentDate, now)) {
      return;
    }
    liveNowTimer = window.setInterval(() => {
      const currentNow = new Date();
      if (!isSameDate(currentDate, currentNow)) {
        stopLiveNowTimer();
        renderForDate();
        return;
      }
      renderForDate();
    }, 30000);
  };

  initializeHelp('help/today.txt', {
    pageKey: 'today',
    showPrompt: false,
    waitForReady: () => (tutorialReady.isResolved() ? true : tutorialReady.promise),
    steps: {
      date: '#selected-date',
      prev: '#prev-day',
      next: '#next-day',
      table: () => document.querySelector('.attendance-grid') || content,
      warnings: '#attendance-status',
      help: () => document.getElementById('help-button')
    }
  });

  stores = loadStores();
  storeKeys = Object.keys(stores);
  selectedStoreKeys.push(
    ...(() => {
      if (requestedStoreKey === 'all') {
        return storeKeys;
      }
      if (requestedStoreKey && storeKeys.includes(requestedStoreKey)) {
        return [requestedStoreKey];
      }
      return [];
    })()
  );

  renderStoreButtons();

  if (selectedStoreKeys.length === 0) {
    setScheduleVisible(false);
    setNavigationDisabled(true);
    if (status) {
      if (requestedStoreKey) {
        status.textContent = '指定された店舗が見つかりません。店舗一覧から選択してください。';
      } else {
        status.textContent = '表示する店舗を選択してください。';
      }
      status.classList.remove('attendance-status--error');
    }
    tutorialReady.resolve();
  } else {
    setScheduleVisible(true);
  }

  if (prevBtn) {
    prevBtn.addEventListener('click', () => {
      currentDate = new Date(currentDate.getTime() - DAY_IN_MS);
      updateDateLabel();
      stopLiveNowTimer();
      renderForDateWithLiveSync();
    });
  }

  if (nextBtn) {
    nextBtn.addEventListener('click', () => {
      currentDate = new Date(currentDate.getTime() + DAY_IN_MS);
      updateDateLabel();
      stopLiveNowTimer();
      renderForDateWithLiveSync();
    });
  }

  if (todayBtn) {
    todayBtn.addEventListener('click', () => {
      currentDate = normalizeDate(new Date());
      updateDateLabel();
      stopLiveNowTimer();
      renderForDateWithLiveSync();
    });
  }

  updateDateLabel();
  if (selectedStoreKeys.length > 0) {
    initializeData();
  }
});
