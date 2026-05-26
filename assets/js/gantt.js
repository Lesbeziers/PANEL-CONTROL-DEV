(function initCalendarColumnDebugModule() {
  const RANGE_CELL_CLASS = "ganttBarCell";
  const RANGE_START_CLASS = "ganttBarStart";
  const RANGE_END_CLASS = "ganttBarEnd";
  const RANGE_START_FLAT_CLASS = "ganttBarStartFlat";
  const RANGE_END_FLAT_CLASS = "ganttBarEndFlat";
  const RANGE_MARKER_CLASS = "ganttBarRangeMarker";
  const RANGE_MARKER_START_CLASS = "ganttBarRangeMarkerStart";
  const RANGE_MARKER_END_CLASS = "ganttBarRangeMarkerEnd";
  const DAY_ATTR = "data-day";
  const OBSERVER_TARGET_SELECTOR = ".month-block #right-body";
  const DATE_COLUMN_SELECTOR = '.left-row > div[data-column-key="startDate"], .left-row > div[data-column-key="endDate"]';
  const GANTT_BODY_SELECTOR = "#left-body, #right-body, .month-block__body-grid";
  const BLOCK_HEADER_GREEN_CLASS = "blockHeader--green";
  const BLOCK_HEADER_YELLOW_CLASS = "blockHeader--yellow";
  const BLOCK_HEADER_PURPLE_CLASS = "blockHeader--purple";
  const GANTT_BLOCK_HEADER_CLASS = "ganttBlockHeader";
  const GANTT_BLOCK_GREEN_CLASS = "ganttBlockGreen";
  const GANTT_BLOCK_YELLOW_CLASS = "ganttBlockYellow";
  const GANTT_BLOCK_PURPLE_CLASS = "ganttBlockPurple";
  const BLOCK_DAY_COUNT_CLASS = "ganttBlockDayCount";
  const BLOCK_OVER_MAX_CLASS = "ganttOverMax";
  const FOCUS_DIM_CLASS = "ganttFocusDim";
  const FOCUS_ACTIVE_CLASS = "ganttFocusActive";
  const GLOBAL_DAY_DIM_CLASS = "ganttGlobalDayDim";
  const GLOBAL_DAY_BAND_CLASS = "ganttGlobalDayBandVisible";
  const BLOCK_DAY_DIM_CLASS = "ganttBlockDayDim";
  const BLOCK_DAY_HOVER_CELL_CLASS = "ganttBlockDayHoverCell";
  const BLOCK_DAY_BAND_CLASS = "ganttBlockDayBandVisible";
  const WEEKEND_CELL_CLASS = "ganttWeekendCell";
  const OUT_OF_MONTH_ROW_CLASS = "ganttOutOfMonthRow";
  const TOOLTIP_CLASS = "ganttBarTooltip";
  const BAR_HOVER_FOCUS_DELAY_MS = 140;
  const BAND_COLOR_GREEN = "#b5d2b3";
  const BAND_COLOR_YELLOW = "#fbdd9a";
  const BAND_COLOR_PURPLE = "#dcc4f0";
  const HEADER_COLOR_GREEN = "#8fb596";
  const HEADER_COLOR_YELLOW = "#e8cd8e";
  const HEADER_COLOR_PURPLE_DARK = "#aa87c6";
  const HEADER_COLOR_PURPLE_LIGHT = "#c7a8e5";
  const MONTH_LABEL_TO_NUMBER = {
    enero: 1,
    febrero: 2,
    marzo: 3,
    abril: 4,
    mayo: 5,
    junio: 6,
    julio: 7,
    agosto: 8,
    septiembre: 9,
    setiembre: 9,
    octubre: 10,
    noviembre: 11,
    diciembre: 12,
  };
  
  let observer = null;
  let rafId = null;
  let activeFocusRow = null;
  let activeFocusBlockRows = [];
  let focusTooltip = null;
  let hoverFocusTimerId = null;
  let pendingHoverFocusRow = null;
  let isRangePressActive = false;
  let headerHoverTimerId = null;
  let pendingHeaderHoverDay = null;
  let activeHeaderHoverDay = null;
  let blockDayHoverTimerId = null;
  let pendingBlockDayHover = null;
  let activeBlockDayHover = null;
  
  let activeHeaderHoverCell = null;

  function getGlobalBandHeight(rightBody) {
    const dayRows = [...rightBody.querySelectorAll(".day-row")];
    if (!dayRows.length) {
      return rightBody.scrollHeight;
    }

    const lastRow = dayRows[dayRows.length - 1];
    const lastRowBottom = lastRow.offsetTop + lastRow.offsetHeight;
    return Math.max(lastRowBottom, rightBody.scrollHeight);
  }

  function updateGlobalHeaderDayBandPosition(root) {
    const headerTrack = root.querySelector("#right-header-track");
    const rightBody = root.querySelector("#right-body");
    if (!headerTrack || !rightBody || !activeHeaderHoverCell || !headerTrack.contains(activeHeaderHoverCell)) {
      return;
    }

    const trackRect = headerTrack.getBoundingClientRect();
    const bodyRect = rightBody.getBoundingClientRect();
    const cellRect = activeHeaderHoverCell.getBoundingClientRect();
    const headerBandLeft = cellRect.left - trackRect.left + headerTrack.scrollLeft;
    const bandLeft = cellRect.left - bodyRect.left + rightBody.scrollLeft;
    const bandWidth = cellRect.width;
    const bandHeight = getGlobalBandHeight(rightBody);

    headerTrack.style.setProperty("--ganttGlobalBandLeft", `${headerBandLeft}px`);
    headerTrack.style.setProperty("--ganttGlobalBandWidth", `${bandWidth}px`);
    
    rightBody.style.setProperty("--ganttGlobalBandLeft", `${bandLeft}px`);
    rightBody.style.setProperty("--ganttGlobalBandWidth", `${bandWidth}px`);
    rightBody.style.setProperty("--ganttGlobalBandHeight", `${bandHeight}px`);
  }

  function clearGlobalHeaderDayBand(root) {
    const headerTrack = root.querySelector("#right-header-track");
    const rightBody = root.querySelector("#right-body");
    if (!headerTrack || !rightBody) {
      return;
    }

    headerTrack.classList.remove(GLOBAL_DAY_BAND_CLASS);
    headerTrack.style.removeProperty("--ganttGlobalBandLeft");
    headerTrack.style.removeProperty("--ganttGlobalBandWidth");

    rightBody.classList.remove(GLOBAL_DAY_BAND_CLASS);
    rightBody.style.removeProperty("--ganttGlobalBandLeft");
    rightBody.style.removeProperty("--ganttGlobalBandWidth");
    rightBody.style.removeProperty("--ganttGlobalBandHeight");
    activeHeaderHoverCell = null;
  }

  function updateBlockDayBandPosition(root) {
    const rightBody = root.querySelector("#right-body");
    if (!rightBody || !activeBlockDayHover) {
      return;
    }

    const { headerCell, headerRow, blockRows } = activeBlockDayHover;
    if (!headerCell || !headerRow || !rightBody.contains(headerCell) || !rightBody.contains(headerRow)) {
      return;
    }

    const bodyRect = rightBody.getBoundingClientRect();
    const cellRect = headerCell.getBoundingClientRect();
    const bandLeft = cellRect.left - bodyRect.left + rightBody.scrollLeft;
    const bandWidth = cellRect.width;
    const bandTop = headerRow.offsetTop;
    const lastBlockRow = blockRows[blockRows.length - 1] || headerRow;
    const blockBottom = lastBlockRow.offsetTop + lastBlockRow.offsetHeight;
    const bandHeight = Math.max(headerRow.offsetHeight, blockBottom - bandTop);

    rightBody.style.setProperty("--ganttBlockBandTop", `${bandTop}px`);
    rightBody.style.setProperty("--ganttBlockBandHeight", `${bandHeight}px`);
    rightBody.style.setProperty("--ganttBlockBandLeft", `${bandLeft}px`);
    rightBody.style.setProperty("--ganttBlockBandWidth", `${bandWidth}px`);
  }

  function clearBlockDayBand(root) {
    const rightBody = root.querySelector("#right-body");
    if (!rightBody) {
      return;
    }

    rightBody.classList.remove(BLOCK_DAY_BAND_CLASS);
    rightBody.style.removeProperty("--ganttBlockBandTop");
    rightBody.style.removeProperty("--ganttBlockBandHeight");
    rightBody.style.removeProperty("--ganttBlockBandLeft");
    rightBody.style.removeProperty("--ganttBlockBandWidth");
  }

  function applyBlockDayBand(root) {
    const rightBody = root.querySelector("#right-body");
    if (!rightBody || !activeBlockDayHover) {
      return;
    }

    updateBlockDayBandPosition(root);
    rightBody.classList.add(BLOCK_DAY_BAND_CLASS);
  }

  function applyGlobalHeaderDayBand(root, headerCell) {
    const headerTrack = root.querySelector("#right-header-track");
    const rightBody = root.querySelector("#right-body");
    if (!headerTrack || !rightBody || !headerCell) {
      return;
    }

    activeHeaderHoverCell = headerCell;
    updateGlobalHeaderDayBandPosition(root);
    headerTrack.classList.add(GLOBAL_DAY_BAND_CLASS);
    rightBody.classList.add(GLOBAL_DAY_BAND_CLASS);
  }
  
  let repaintRafId = null;
  let repaintRafId2 = null;
  let repaintTimeoutId = null;
  let stableRepaintRunId = 0;
  let pendingFullRepaint = false;
  const pendingRowsToRepaint = new Set();

  function waitPostRenderTick() {
    return new Promise((resolve) => {
      window.requestAnimationFrame(() => {
        window.setTimeout(resolve, 16);
      });
    });
  }

  function getStartEndHash(root) {
    const leftRows = [...root.querySelectorAll("#left-body .left-row")];
    const dayRows = [...root.querySelectorAll("#right-body .day-row")];
    const pairs = [];

    leftRows.forEach((leftRow, rowIndex) => {
      if (!isDataRow(dayRows[rowIndex], leftRow)) {
        return;
      }

      const startText = leftRow.querySelector('[data-column-key="startDate"]')?.textContent?.trim() || "";
      const endText = leftRow.querySelector('[data-column-key="endDate"]')?.textContent?.trim() || "";
      pairs.push(`${startText}→${endText}`);
    });

    return pairs.join("||");
  }

  function repaintAll(root) {
    markCalendarCells(root);
  }

  async function repaintUntilStable(root, { maxMs = 600 } = {}) {
    const runId = ++stableRepaintRunId;
    const startedAt = performance.now();
    let previousHash = null;
    let stableIterations = 0;
    let iteration = 0;

    while (performance.now() - startedAt <= maxMs) {
      if (runId !== stableRepaintRunId) {
        return;
      }

      iteration += 1;
      repaintAll(root);
      await waitPostRenderTick();

      const currentHash = getStartEndHash(root);
      const elapsedMs = Math.round(performance.now() - startedAt);


      if (currentHash === previousHash) {
        stableIterations += 1;
      } else {
        stableIterations = 0;
      }

      if (stableIterations >= 2) {
        return;
      }

      previousHash = currentHash;
    }

  }
  function parseDayLabel(value) {
    const normalized = (value || "").trim();
    if (!/^\d{1,2}$/.test(normalized)) {
      return null;
    }

    const day = Number.parseInt(normalized, 10);
    if (!Number.isInteger(day) || day < 1 || day > 31) {
      return null;
    }

    return day;
  }

  function daysInMonth(month, year) {
    return new Date(year, month, 0).getDate();
  }

  function normalizeMonthLabel(value) {
    return `${value ?? ""}`
      .trim()
      .toLocaleLowerCase("es-ES")
      .normalize("NFD")
      .replace(/\p{Diacritic}/gu, "");
  }

  function getCalendarContext(root) {
    const now = new Date();
    const fallback = {
      month: now.getMonth() + 1,
      year: now.getFullYear(),
      daysInMonth: daysInMonth(now.getMonth() + 1, now.getFullYear()),
    };

    const ownerDocument = root?.ownerDocument || document;
    const titleElement = root?.querySelector?.(".panel-layout__month-title")
      || ownerDocument.querySelector(".panel-layout__month-title");
    const titleText = titleElement?.textContent?.trim() || "";
    const match = titleText.match(/^([A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+)\s+(\d{4})$/u);
    if (!match) {
      return fallback;
    }

    const month = MONTH_LABEL_TO_NUMBER[normalizeMonthLabel(match[1])];
    const year = Number.parseInt(match[2], 10);
    if (!Number.isInteger(month) || !Number.isInteger(year)) {
      return fallback;
    }

    return {
      month,
      year,
      daysInMonth: daysInMonth(month, year),
    };
  }

  function parseDateValue(value, calendarContext) {
    const parser = window.PanelDateUtils?.parseDatePartsFromText;
    if (typeof parser !== "function") {
      return null;
    }

    const parsed = parser(value, {
      defaultMonth: calendarContext?.month,
      defaultYear: calendarContext?.year,
    });

    if (!parsed || !parsed.ok || parsed.empty) {
      return null;
    }

    return {
      date: parsed.date,
      month: parsed.month,
      hasExplicitYear: parsed.hasExplicitYear,
    };
  }
  
  function getDateRangeForRow(leftRow, calendarContext) {
    if (!leftRow) {
      return null;
    }

    const startText = leftRow.querySelector('[data-column-key="startDate"]')?.textContent?.trim() || "";
    const endText = leftRow.querySelector('[data-column-key="endDate"]')?.textContent?.trim() || "";
    const start = parseDateValue(startText, calendarContext);
    const end = parseDateValue(endText, calendarContext);
    if (!start || !end) {
      return null;
    }

    const startDate = start.date;
    let endDate = end.date;

    if (endDate < startDate && !end.hasExplicitYear && end.month < start.month) {
      endDate = new Date(endDate.getFullYear() + 1, endDate.getMonth(), endDate.getDate());
    }

    if (startDate > endDate) {
      return null;
    }

    return { startDate, endDate };
  }

  function getVisibleMonthRange(calendarContext) {
    return {
      startDate: new Date(calendarContext.year, calendarContext.month - 1, 1),
      endDate: new Date(calendarContext.year, calendarContext.month - 1, calendarContext.daysInMonth),
    };
  }

  function getVisibleDayIntervalForRow(leftRow, calendarContext) {
    const rowRange = getDateRangeForRow(leftRow, calendarContext);
    if (!rowRange) {
      return null;
    }

    const monthRange = getVisibleMonthRange(calendarContext);
    const visibleStart = rowRange.startDate > monthRange.startDate ? rowRange.startDate : monthRange.startDate;
    const visibleEnd = rowRange.endDate < monthRange.endDate ? rowRange.endDate : monthRange.endDate;
    if (visibleStart > visibleEnd) {
      return null;
    }

    return {
      startDay: visibleStart.getDate(),
      endDay: visibleEnd.getDate(),
      startsBeforeMonth: rowRange.startDate < monthRange.startDate,
      endsAfterMonth: rowRange.endDate > monthRange.endDate,
    };
  }

    function rowIsOutsideVisibleMonth(leftRow, calendarContext) {
    const rowRange = getDateRangeForRow(leftRow, calendarContext);
    if (!rowRange) {
      return false;
    }

    const monthRange = getVisibleMonthRange(calendarContext);
    return rowRange.endDate < monthRange.startDate || rowRange.startDate > monthRange.endDate;
  }

  function syncOutOfMonthRowState(dayRow, leftRow, calendarContext) {
    if (!dayRow || !leftRow || !isDataRow(dayRow, leftRow)) {
      dayRow?.classList.remove(OUT_OF_MONTH_ROW_CLASS);
      leftRow?.classList.remove(OUT_OF_MONTH_ROW_CLASS);
      return;
    }

    const isOutOfMonth = rowIsOutsideVisibleMonth(leftRow, calendarContext);
    dayRow.classList.toggle(OUT_OF_MONTH_ROW_CLASS, isOutOfMonth);
    leftRow.classList.toggle(OUT_OF_MONTH_ROW_CLASS, isOutOfMonth);
  }

  function detectCalendarColumns(headerTrack) {
    const headerCells = [...headerTrack.children];
    return headerCells.reduce((acc, headerCell, columnIndex) => {
      const dayFromAttr = parseDayLabel(headerCell.getAttribute(DAY_ATTR));
      const day = dayFromAttr ?? parseDayLabel(headerCell.textContent);
      if (day !== null) {
        acc.push({ columnIndex, day });
      }
      return acc;
    }, []);
  }

    function isWeekendDay(day, calendarContext) {
    if (!Number.isInteger(day) || !calendarContext) {
      return false;
    }

    const weekDay = new Date(calendarContext.year, calendarContext.month - 1, day).getDay();
    return weekDay === 0 || weekDay === 6;
  }

  function clearRangeCells(dayRow) {
    if (!dayRow) {
      return;
    }

    dayRow.querySelectorAll(".day-cell").forEach((cell) => {
      cell.classList.remove(RANGE_CELL_CLASS);
      cell.classList.remove(RANGE_START_CLASS);
      cell.classList.remove(RANGE_END_CLASS);
      cell.classList.remove(RANGE_START_FLAT_CLASS);
      cell.classList.remove(RANGE_END_FLAT_CLASS);
      cell.querySelectorAll(`.${RANGE_MARKER_CLASS}`).forEach((marker) => marker.remove());
    });
  }

  function appendRangeMarker(cell, markerClass) {
    if (!cell || !markerClass) {
      return;
    }

    if (cell.querySelector(`.${markerClass}`)) {
      return;
    }

    const marker = document.createElement("span");
    marker.className = `${RANGE_MARKER_CLASS} ${markerClass}`;
    marker.setAttribute("aria-hidden", "true");
    cell.appendChild(marker);
  }

  function normalizeColor(value) {
    const normalized = (value || "").trim().toLowerCase();
    if (!normalized) {
      return "";
    }

    if (normalized.startsWith("#")) {
      if (normalized.length === 4) {
        const [r, g, b] = normalized.slice(1);
        return `#${r}${r}${g}${g}${b}${b}`;
      }

      return normalized;
    }

    const rgbMatch = normalized.match(/rgba?\((\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (rgbMatch) {
      const toHex = (part) => Number.parseInt(part, 10).toString(16).padStart(2, "0");
      return `#${toHex(rgbMatch[1])}${toHex(rgbMatch[2])}${toHex(rgbMatch[3])}`;
    }

    return normalized;
  }

  function matchesColor(value, targetHex) {
    return normalizeColor(value) === normalizeColor(targetHex);
  }

  function findLastMatchingDescendant(root, selector) {
    if (!root) {
      return null;
    }

    for (let child = root.lastElementChild; child; child = child.previousElementSibling) {
      const matchInChild = findLastMatchingDescendant(child, selector);
      if (matchInChild) {
        return matchInChild;
      }
    }

    if (root.matches?.(selector)) {
      return root;
    }

    return null;
  }

  function findPreviousMatchingElementAcrossContainers(element, selector, stopContainer) {
    if (!element) {
      return null;
    }

    let current = element;
    while (current && current !== stopContainer) {
      for (let sibling = current.previousElementSibling; sibling; sibling = sibling.previousElementSibling) {
        const candidate = findLastMatchingDescendant(sibling, selector);
        if (candidate) {
          return candidate;
        }
      }

      const parent = current.parentElement;
      if (!parent || parent === stopContainer) {
        return null;
      }

      current = parent;
    }

    return null;
  }

  function ensureBlockHeaderClass(headerRow) {
    if (!headerRow || !headerRow.classList.contains("group")) {
      return null;
    }

        headerRow.classList.add(GANTT_BLOCK_HEADER_CLASS);

    if (headerRow.classList.contains(GANTT_BLOCK_GREEN_CLASS)) {
      return BLOCK_HEADER_GREEN_CLASS;
    }

    if (headerRow.classList.contains(GANTT_BLOCK_YELLOW_CLASS)) {
      return BLOCK_HEADER_YELLOW_CLASS;
    }

    if (headerRow.classList.contains(GANTT_BLOCK_PURPLE_CLASS)) {
      return BLOCK_HEADER_PURPLE_CLASS;
    }

    if (headerRow.classList.contains(BLOCK_HEADER_GREEN_CLASS)) {
      headerRow.classList.add(GANTT_BLOCK_GREEN_CLASS);
      return BLOCK_HEADER_GREEN_CLASS;
    }

    if (headerRow.classList.contains(BLOCK_HEADER_YELLOW_CLASS)) {
      headerRow.classList.add(GANTT_BLOCK_YELLOW_CLASS);
      return BLOCK_HEADER_YELLOW_CLASS;
    }

    if (headerRow.classList.contains(BLOCK_HEADER_PURPLE_CLASS)) {
      headerRow.classList.add(GANTT_BLOCK_PURPLE_CLASS);
      return BLOCK_HEADER_PURPLE_CLASS;
    }

    const inlineGroupBg = headerRow.style.getPropertyValue("--group-bg");
    if (matchesColor(inlineGroupBg, HEADER_COLOR_GREEN)) {
      headerRow.classList.add(BLOCK_HEADER_GREEN_CLASS);
      headerRow.classList.add(GANTT_BLOCK_GREEN_CLASS);
      return BLOCK_HEADER_GREEN_CLASS;
    }

    if (matchesColor(inlineGroupBg, HEADER_COLOR_YELLOW)) {
      headerRow.classList.add(BLOCK_HEADER_YELLOW_CLASS);
      headerRow.classList.add(GANTT_BLOCK_YELLOW_CLASS);
      return BLOCK_HEADER_YELLOW_CLASS;
    }

    if (matchesColor(inlineGroupBg, HEADER_COLOR_PURPLE_DARK) || matchesColor(inlineGroupBg, HEADER_COLOR_PURPLE_LIGHT)) {
      headerRow.classList.add(BLOCK_HEADER_PURPLE_CLASS);
      headerRow.classList.add(GANTT_BLOCK_PURPLE_CLASS);
      return BLOCK_HEADER_PURPLE_CLASS;
    }

    const rowBg = window.getComputedStyle(headerRow).backgroundColor;
    const headerCell = headerRow.querySelector(".day-cell, div");
    const cellBg = headerCell ? window.getComputedStyle(headerCell).backgroundColor : "";

    if (matchesColor(rowBg, HEADER_COLOR_GREEN) || matchesColor(cellBg, HEADER_COLOR_GREEN)) {
      headerRow.classList.add(BLOCK_HEADER_GREEN_CLASS);
      headerRow.classList.add(GANTT_BLOCK_GREEN_CLASS);
      return BLOCK_HEADER_GREEN_CLASS;
    }

    if (matchesColor(rowBg, HEADER_COLOR_YELLOW) || matchesColor(cellBg, HEADER_COLOR_YELLOW)) {
      headerRow.classList.add(BLOCK_HEADER_YELLOW_CLASS);
      headerRow.classList.add(GANTT_BLOCK_YELLOW_CLASS);
      return BLOCK_HEADER_YELLOW_CLASS;
    }

    if (matchesColor(rowBg, HEADER_COLOR_PURPLE_DARK) || matchesColor(cellBg, HEADER_COLOR_PURPLE_DARK)
        || matchesColor(rowBg, HEADER_COLOR_PURPLE_LIGHT) || matchesColor(cellBg, HEADER_COLOR_PURPLE_LIGHT)) {
      headerRow.classList.add(BLOCK_HEADER_PURPLE_CLASS);
      headerRow.classList.add(GANTT_BLOCK_PURPLE_CLASS);
      return BLOCK_HEADER_PURPLE_CLASS;
    }

    return null;
  }

  function findNearestBlockHeader(row) {
    const rightBody = row?.closest("#right-body");
    if (!rightBody) {
      return null;
    }

    let cursor = row;
    while (cursor) {
      cursor = findPreviousMatchingElementAcrossContainers(cursor, ".day-row", rightBody);
      if (!cursor) {
        return null;
      }

      if (cursor.classList?.contains("group")) {
        return cursor;
      }
    }

    return null;
  }

  function getBlockBandColor(row) {
    const headerRow = findNearestBlockHeader(row);
    const headerClass = ensureBlockHeaderClass(headerRow);
    if (headerClass === BLOCK_HEADER_YELLOW_CLASS) {
      return BAND_COLOR_YELLOW;
    }

    if (headerClass === BLOCK_HEADER_GREEN_CLASS) {
      return BAND_COLOR_GREEN;
    }

    if (headerClass === BLOCK_HEADER_PURPLE_CLASS) {
      return BAND_COLOR_PURPLE;
    }

    return null;
  }

  function paintRangeForRow(dayRow, leftRow, calendarContext) {
    clearRangeCells(dayRow);
    syncOutOfMonthRowState(dayRow, leftRow, calendarContext);
    
    const bandColor = getBlockBandColor(dayRow);
    if (bandColor) {
      dayRow.style.setProperty("--ganttBarColor", bandColor);
      dayRow.style.setProperty("--gantt-band-color", bandColor);
    } else {
      dayRow.style.removeProperty("--ganttBarColor");
      dayRow.style.removeProperty("--gantt-band-color");
    }

    if (!isDataRow(dayRow, leftRow)) {
      return;
    }

    const startCell = leftRow.querySelector('[data-column-key="startDate"]');
    const endCell = leftRow.querySelector('[data-column-key="endDate"]');
    if (!startCell || !endCell) {
      return;
    }

    const visibleInterval = getVisibleDayIntervalForRow(leftRow, calendarContext);
    if (!visibleInterval) {
      return;
    }

    const rangeCells = [];
    dayRow.querySelectorAll(`.day-cell[${DAY_ATTR}]`).forEach((cell) => {
      const day = Number.parseInt(cell.getAttribute(DAY_ATTR), 10);
      if (Number.isInteger(day) && day >= visibleInterval.startDay && day <= visibleInterval.endDay) {
        cell.classList.add(RANGE_CELL_CLASS);
        rangeCells.push(cell);
      }
    });

    if (!rangeCells.length) {
      return;
    }

    rangeCells[0].classList.add(RANGE_START_CLASS);
    rangeCells[rangeCells.length - 1].classList.add(RANGE_END_CLASS);

    if (visibleInterval.startsBeforeMonth) {
      rangeCells[0].classList.add(RANGE_START_FLAT_CLASS);
      appendRangeMarker(rangeCells[0], RANGE_MARKER_START_CLASS);
    }

    if (visibleInterval.endsAfterMonth) {
      rangeCells[rangeCells.length - 1].classList.add(RANGE_END_FLAT_CLASS);
      appendRangeMarker(rangeCells[rangeCells.length - 1], RANGE_MARKER_END_CLASS);
    }
  }

  function updateHeaderDayCountCell(cell, count) {
    if (!cell) {
      return;
    }

    let countNode = cell.querySelector(`.${BLOCK_DAY_COUNT_CLASS}`);
    if (!countNode) {
      countNode = document.createElement("span");
      countNode.className = BLOCK_DAY_COUNT_CLASS;
      cell.appendChild(countNode);
    }

    countNode.textContent = count > 0 ? String(count) : "";
  }

  function extractBlockMaxSimultaneous(value) {
    const normalized = (value || "").replace(/\s+/g, " ").trim();
    if (!normalized) {
      return null;
    }

    const scopedMatch = normalized.match(/m[aá]x(?:imo)?\.?\s*(\d+)\s*simult[aá]neas/i);
    if (scopedMatch) {
      const scopedValue = Number.parseInt(scopedMatch[1], 10);
      return Number.isInteger(scopedValue) ? scopedValue : null;
    }

    const fallbackMatch = normalized.match(/\(\s*m[aá]x(?:imo)?\.?[^\d]*(\d+)\s*simult[aá]neas\s*\)/i);
    if (fallbackMatch) {
      const fallbackValue = Number.parseInt(fallbackMatch[1], 10);
      return Number.isInteger(fallbackValue) ? fallbackValue : null;
    }

    return null;
  }

  function renderBlockDailyCounts(dayRows, leftRows, calendarContext) {
    const blockEntries = [];
    let activeBlock = null;

    dayRows.forEach((dayRow, rowIndex) => {
      const leftRow = leftRows[rowIndex] || null;

      if (dayRow.classList.contains("group")) {
        const headerClass = ensureBlockHeaderClass(dayRow);
        if (!headerClass) {
          activeBlock = null;
          return;
        }

        activeBlock = {
          headerRow: dayRow,
          maxSimultaneous: extractBlockMaxSimultaneous(leftRow?.textContent || ""),
          counts: new Array(32).fill(0),
          hasVisibleRows: false,
        };

        activeBlock.persistedCounts = `${dayRow.dataset.blockDailyCounts || ""}`
          .split(",")
          .map((value) => Number.parseInt(value, 10));

        blockEntries.push(activeBlock);
        return;
      }

      if (!activeBlock || !isDataRow(dayRow, leftRow)) {
        return;
      }

      activeBlock.hasVisibleRows = true;

      const visibleInterval = getVisibleDayIntervalForRow(leftRow, calendarContext);
      if (!visibleInterval) {
        return;
      }

      const fromDay = Math.max(1, visibleInterval.startDay);
      const toDay = Math.min(calendarContext.daysInMonth, visibleInterval.endDay);
      for (let day = fromDay; day <= toDay; day += 1) {
        activeBlock.counts[day] += 1;
      }
    });

    blockEntries.forEach((entry) => {
      if (entry.hasVisibleRows) {
        return;
      }

      if (entry.persistedCounts.length < calendarContext.daysInMonth) {
        return;
      }

      for (let day = 1; day <= calendarContext.daysInMonth; day += 1) {
        const persistedValue = entry.persistedCounts[day - 1];
        entry.counts[day] = Number.isInteger(persistedValue) ? persistedValue : 0;
      }
    });

    blockEntries.forEach(({ headerRow, counts, maxSimultaneous }) => {
      const calendarCells = headerRow.querySelectorAll(`.day-cell[${DAY_ATTR}]`);
      calendarCells.forEach((cell) => {
        const day = Number.parseInt(cell.getAttribute(DAY_ATTR), 10);
        const count = Number.isInteger(day) ? counts[day] : 0;
        updateHeaderDayCountCell(cell, count);

        const isOverMax = Number.isInteger(maxSimultaneous) && count > maxSimultaneous;
        cell.classList.toggle(BLOCK_OVER_MAX_CLASS, Boolean(isOverMax));
      });
    });
  }

  function markCalendarCells(root) {
    const headerTrack = root.querySelector("#right-header-track");
    if (!headerTrack) {
      return;
    }

    const calendarContext = getCalendarContext(root);
    const calendarColumns = detectCalendarColumns(headerTrack);
    if (!calendarColumns.length) {
      return;
    }

    const headerDayCells = [...headerTrack.querySelectorAll(".day-cell")];
    headerDayCells.forEach((cell) => {
      cell.classList.remove(WEEKEND_CELL_CLASS);
      cell.removeAttribute(DAY_ATTR);
    });

    const allDayCells = root.querySelectorAll("#right-body .day-row .day-cell");
    allDayCells.forEach((cell) => {
      cell.classList.remove(RANGE_CELL_CLASS);
      cell.classList.remove(RANGE_START_CLASS);
      cell.classList.remove(RANGE_END_CLASS);
      cell.classList.remove(RANGE_START_FLAT_CLASS);
      cell.classList.remove(RANGE_END_FLAT_CLASS);
      cell.classList.remove(WEEKEND_CELL_CLASS);
      cell.classList.remove(BLOCK_OVER_MAX_CLASS);
      cell.removeAttribute(DAY_ATTR);
      cell.querySelectorAll(`.${RANGE_MARKER_CLASS}`).forEach((marker) => marker.remove());
      cell.querySelector(`.${BLOCK_DAY_COUNT_CLASS}`)?.remove();
    });

    calendarColumns.forEach(({ columnIndex, day }) => {
      const headerCell = headerDayCells[columnIndex];
      if (!headerCell) {
        return;
      }

      const isInsideCurrentMonth = day <= calendarContext.daysInMonth;
      if (!isInsideCurrentMonth) {
        return;
      }

      headerCell.setAttribute(DAY_ATTR, String(day));
      headerCell.classList.toggle(WEEKEND_CELL_CLASS, isWeekendDay(day, calendarContext));
    });

    const dayRows = [...root.querySelectorAll("#right-body .day-row")];
    const leftRows = [...root.querySelectorAll("#left-body .left-row")];
    dayRows.forEach((row, rowIndex) => {
      const rowCells = [...row.children];
      calendarColumns.forEach(({ columnIndex, day }) => {
        const targetCell = rowCells[columnIndex];
        if (!targetCell) {
          return;
        }

        const isInsideCurrentMonth = day <= calendarContext.daysInMonth;
        if (!isInsideCurrentMonth) {
          targetCell.removeAttribute(DAY_ATTR);
          targetCell.classList.remove(WEEKEND_CELL_CLASS);
          return;
        }

        targetCell.setAttribute(DAY_ATTR, String(day));
        targetCell.classList.toggle(WEEKEND_CELL_CLASS, isWeekendDay(day, calendarContext));
      });

      if (!isDataRow(row, leftRows[rowIndex])) {
        return;
      }

      paintRangeForRow(row, leftRows[rowIndex], calendarContext);
    });

    renderBlockDailyCounts(dayRows, leftRows, calendarContext);
  }

  function paintSingleRowByLeftRow(root, leftRow) {
    if (!leftRow) {
      return;
    }

    const leftRows = [...root.querySelectorAll("#left-body .left-row")];
    const rowIndex = leftRows.indexOf(leftRow);
    if (rowIndex < 0) {
      return;
    }

    const dayRows = root.querySelectorAll("#right-body .day-row");
    const dayRow = dayRows[rowIndex];
    if (!dayRow) {
      return;
    }

    paintRangeForRow(dayRow, leftRow, getCalendarContext(root));
  }

  function paintAllRows(root) {
    const dayRows = [...root.querySelectorAll("#right-body .day-row")];
    const leftRows = [...root.querySelectorAll("#left-body .left-row")];
    dayRows.forEach((dayRow, rowIndex) => {
      paintRangeForRow(dayRow, leftRows[rowIndex], getCalendarContext(root));
    });
  }

  function getRowIndexByElement(root, row, selector) {
    const rows = [...root.querySelectorAll(selector)];
    return rows.indexOf(row);
  }

  function getBlockDataRows(root, dayRow) {
    const dayRows = [...root.querySelectorAll("#right-body .day-row")];
    const startIndex = dayRows.indexOf(dayRow);
    if (startIndex < 0) {
      return [];
    }

    let blockStart = startIndex;
    while (blockStart > 0 && !dayRows[blockStart].classList.contains("group")) {
      blockStart -= 1;
    }

    if (!dayRows[blockStart]?.classList.contains("group")) {
      return [];
    }

    const blockRows = [];
    for (let index = blockStart + 1; index < dayRows.length; index += 1) {
      const currentRow = dayRows[index];
      if (currentRow.classList.contains("group")) {
        break;
      }

      blockRows.push(currentRow);
    }

    return blockRows;
  }

    function getBlockRows(root, headerRow) {
    if (!headerRow?.classList.contains("group")) {
      return [];
    }

    return [headerRow, ...getBlockDataRows(root, headerRow)];
  }

  function getLeftRowForDayRow(root, dayRow) {
    const rowIndex = getRowIndexByElement(root, dayRow, "#right-body .day-row");
    if (rowIndex < 0) {
      return null;
    }

    const leftRows = root.querySelectorAll("#left-body .left-row");
    return leftRows[rowIndex] || null;
  }

  function parseDateFromCellValue(value, calendarContext) {
    const normalized = (value || "").trim();
    if (!normalized) {
      return null;
    }

    const parsed = parseDateValue(value, calendarContext);
    if (parsed) {
      return parsed.date;
    }

    const monthMap = {
      ene: 0, enero: 0, feb: 1, febrero: 1, mar: 2, marzo: 2,
      abr: 3, abril: 3, may: 4, mayo: 4, jun: 5, junio: 5,
      jul: 6, julio: 6, ago: 7, agosto: 7, sep: 8, sept: 8, septiembre: 8,
      oct: 9, octubre: 9, nov: 10, noviembre: 10, dic: 11, diciembre: 11,
    };
    const textMatch = normalized.toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").match(/^(\d{1,2})\s*([a-z]+)(?:\s*(\d{2,4}))?$/i);
    if (textMatch) {
      const day = Number.parseInt(textMatch[1], 10);
      const month = monthMap[textMatch[2]];
      const yearPart = textMatch[3];
      const year = yearPart ? Number.parseInt(yearPart.length === 2 ? `20${yearPart}` : yearPart, 10) : calendarContext.year;
      if (Number.isInteger(month)) {
        return new Date(year, month, day);
      }
    }

    return null;
  }

  function formatTooltipDate(date) {
    const MONTHS_ES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"];
    return `${String(date.getDate()).padStart(2, "0")} ${MONTHS_ES[date.getMonth()] || ""}`.trim();
  }

  function getBarDateRangeData(leftRow) {
    if (!leftRow) {
      return null;
    }

    const startText = leftRow.querySelector('[data-column-key="startDate"]')?.textContent?.trim() || "";
    const endText = leftRow.querySelector('[data-column-key="endDate"]')?.textContent?.trim() || "";
    if (!startText || !endText) {
      return null;
    }

    const calendarContext = getCalendarContext(document);
    const startDate = parseDateFromCellValue(startText, calendarContext);
    const endDate = parseDateFromCellValue(endText, calendarContext);
    if (startDate && endDate && !Number.isNaN(startDate.getTime()) && !Number.isNaN(endDate.getTime())) {
      const oneDay = 24 * 60 * 60 * 1000;
      const durationDays = Math.floor((endDate - startDate) / oneDay) + 1;
      return {
        line1: `${formatTooltipDate(startDate)} - ${formatTooltipDate(endDate)}`,
        line2: `${Math.max(1, durationDays)} días`,
      };
    }

    return null;
  }

  function ensureFocusTooltip(root) {
    if (focusTooltip && root.contains(focusTooltip)) {
      return focusTooltip;
    }

    focusTooltip = document.createElement("div");
    focusTooltip.className = TOOLTIP_CLASS;
    focusTooltip.hidden = true;
    const rightBody = root.querySelector("#right-body");
    rightBody?.appendChild(focusTooltip);
    return focusTooltip;
  }

  function clearBlockFocus(root) {
    if (hoverFocusTimerId !== null) {
      window.clearTimeout(hoverFocusTimerId);
      hoverFocusTimerId = null;
      pendingHoverFocusRow = null;
    }

    if (activeFocusBlockRows.length) {
      activeFocusBlockRows.forEach((blockRow) => {
        blockRow.classList.remove(FOCUS_DIM_CLASS, FOCUS_ACTIVE_CLASS);
        const leftRow = getLeftRowForDayRow(root, blockRow);
        leftRow?.classList.remove(FOCUS_DIM_CLASS, FOCUS_ACTIVE_CLASS);
      });
    }

    activeFocusRow = null;
    activeFocusBlockRows = [];
    if (focusTooltip) {
      focusTooltip.hidden = true;
      focusTooltip.textContent = "";
    }
  }

    function rowIncludesDay(root, dayRow, day) {
    const leftRow = getLeftRowForDayRow(root, dayRow);
    if (!isDataRow(dayRow, leftRow)) {
      return false;
    }

    const startText = leftRow.querySelector('[data-column-key="startDate"]')?.textContent || "";
    const endText = leftRow.querySelector('[data-column-key="endDate"]')?.textContent || "";
    const calendarContext = getCalendarContext(root);
    const visibleInterval = getVisibleDayIntervalForRow(leftRow, calendarContext);
    if (!visibleInterval) {
      return false;
    }

    return day >= visibleInterval.startDay && day <= visibleInterval.endDay;
  }

  function clearGlobalHeaderDayFocus(root) {
    if (headerHoverTimerId !== null) {
      window.clearTimeout(headerHoverTimerId);
      headerHoverTimerId = null;
    }

    pendingHeaderHoverDay = null;
    activeHeaderHoverDay = null;
    clearGlobalHeaderDayBand(root);
    
    const dayRows = root.querySelectorAll(`#right-body .day-row.${GLOBAL_DAY_DIM_CLASS}`);
    dayRows.forEach((dayRow) => {
      dayRow.classList.remove(GLOBAL_DAY_DIM_CLASS);
    });

    const leftRows = root.querySelectorAll(`#left-body .left-row.${GLOBAL_DAY_DIM_CLASS}`);
    leftRows.forEach((leftRow) => {
      leftRow.classList.remove(GLOBAL_DAY_DIM_CLASS);
    });
  }

  function applyGlobalHeaderDayFocus(root, day) {
    if (activeHeaderHoverDay === day) {
      return;
    }

    activeHeaderHoverDay = day;
    const dayRows = [...root.querySelectorAll("#right-body .day-row")];
    dayRows.forEach((dayRow) => {
      const leftRow = getLeftRowForDayRow(root, dayRow);
      if (!isDataRow(dayRow, leftRow)) {
        dayRow.classList.remove(GLOBAL_DAY_DIM_CLASS);
        leftRow?.classList.remove(GLOBAL_DAY_DIM_CLASS);
        return;
      }

      const includesDay = rowIncludesDay(root, dayRow, day);
      dayRow.classList.toggle(GLOBAL_DAY_DIM_CLASS, !includesDay);
      leftRow.classList.toggle(GLOBAL_DAY_DIM_CLASS, !includesDay);
    });
  }

  function applyBlockFocus(root, dayRow) {
    if (activeFocusRow === dayRow) {
      return;
    }

    clearBlockFocus(root);

    const blockRows = getBlockDataRows(root, dayRow).filter((row) => {
      const leftRow = getLeftRowForDayRow(root, row);
      return isDataRow(row, leftRow);
    });

    if (!blockRows.length) {
      return;
    }

    activeFocusRow = dayRow;
    activeFocusBlockRows = blockRows;

    blockRows.forEach((blockRow) => {
      blockRow.classList.add(FOCUS_DIM_CLASS);
      const leftRow = getLeftRowForDayRow(root, blockRow);
      leftRow?.classList.add(FOCUS_DIM_CLASS);
    });

    dayRow.classList.remove(FOCUS_DIM_CLASS);
    dayRow.classList.add(FOCUS_ACTIVE_CLASS);

    const leftActiveRow = getLeftRowForDayRow(root, dayRow);
    leftActiveRow?.classList.remove(FOCUS_DIM_CLASS);
    leftActiveRow?.classList.add(FOCUS_ACTIVE_CLASS);

    const tooltip = ensureFocusTooltip(root);
    const rowRangeCells = dayRow.querySelectorAll(`.day-cell.${RANGE_CELL_CLASS}`);
    const firstCell = rowRangeCells[0];
    const lastCell = rowRangeCells[rowRangeCells.length - 1];
    const tooltipData = getBarDateRangeData(leftActiveRow);
    if (!tooltip || !firstCell || !lastCell || !tooltipData) {
      return;
    }

    tooltip.innerHTML = `<div>${tooltipData.line1}</div><div>${tooltipData.line2}</div>`;
    tooltip.hidden = false;

    const rightBody = root.querySelector("#right-body");
    if (!rightBody) {
      return;
    }

    const bodyRect = rightBody.getBoundingClientRect();
    const firstRect = firstCell.getBoundingClientRect();
    const lastRect = lastCell.getBoundingClientRect();
    const centerX = (firstRect.left + lastRect.right) / 2;
    const top = firstRect.top - bodyRect.top - tooltip.offsetHeight - 8 + rightBody.scrollTop;
    const left = centerX - bodyRect.left - tooltip.offsetWidth / 2 + rightBody.scrollLeft;

    tooltip.style.left = `${Math.max(4, left)}px`;
    tooltip.style.top = `${Math.max(4, top)}px`;
  }

  function clearBlockDayFocus(root) {
    if (blockDayHoverTimerId !== null) {
      window.clearTimeout(blockDayHoverTimerId);
      blockDayHoverTimerId = null;
    }

    pendingBlockDayHover = null;

    if (activeBlockDayHover) {
      const { dimRows, hoveredHeaderCell, hoveredGlobalHeaderCell } = activeBlockDayHover;
      dimRows.forEach((dayRow) => {
        dayRow.classList.remove(BLOCK_DAY_DIM_CLASS);
        const leftRow = getLeftRowForDayRow(root, dayRow);
        leftRow?.classList.remove(BLOCK_DAY_DIM_CLASS);
      });

      hoveredHeaderCell?.classList.remove(BLOCK_DAY_HOVER_CELL_CLASS);
      hoveredGlobalHeaderCell?.classList.remove(BLOCK_DAY_HOVER_CELL_CLASS);
    }

    activeBlockDayHover = null;
    clearBlockDayBand(root);
  }

  function applyBlockDayFocus(root, headerRow, headerCell, day) {
    clearBlockDayFocus(root);

    const blockDataRows = getBlockDataRows(root, headerRow).filter((row) => {
      const leftRow = getLeftRowForDayRow(root, row);
      return isDataRow(row, leftRow);
    });

    if (!blockDataRows.length) {
      return;
    }

    const dimRows = [];
    blockDataRows.forEach((dayRow) => {
      const includesDay = rowIncludesDay(root, dayRow, day);
      if (includesDay) {
        return;
      }

      dayRow.classList.add(BLOCK_DAY_DIM_CLASS);
      const leftRow = getLeftRowForDayRow(root, dayRow);
      leftRow?.classList.add(BLOCK_DAY_DIM_CLASS);
      dimRows.push(dayRow);
    });
    
    const globalHeaderCell = [...root.querySelectorAll("#right-header-track .day-cell")].find((cell) => {
      const cellDay = parseDayLabel(cell.getAttribute(DAY_ATTR) || cell.textContent);
      return cellDay === day;
    }) || null;

    headerCell.classList.add(BLOCK_DAY_HOVER_CELL_CLASS);
    globalHeaderCell?.classList.add(BLOCK_DAY_HOVER_CELL_CLASS);
    activeBlockDayHover = {
      hoveredHeaderCell: headerCell,
      hoveredGlobalHeaderCell: globalHeaderCell,
      headerCell,
      headerRow,
      day,
      dimRows,
      blockRows: getBlockRows(root, headerRow),
    };
    applyBlockDayBand(root);
  }

  function attachHoverFocusListeners(root) {
    const getEventTargetElement = (event) => {
      if (event.target instanceof Element) {
        return event.target;
      }

      if (event.target instanceof Node) {
        return event.target.parentElement;
      }

      return null;
    };

    const getRangeCellFromEvent = (event) => {
      if (!(event.target instanceof Element)) {
        return null;
      }

      return event.target.closest(`.day-cell.${RANGE_CELL_CLASS}`);
    };

    root.addEventListener("mouseover", (event) => {
      const targetElement = getEventTargetElement(event);
      const blockDayCountNode = targetElement?.closest(`.${BLOCK_DAY_COUNT_CLASS}`) || null;
      const blockHeaderDayCell = blockDayCountNode?.closest(`#right-body .day-row.group .day-cell`) || null;
      if (!blockHeaderDayCell) {
        return;
      }

      const hoverCount = blockHeaderDayCell.querySelector(`.${BLOCK_DAY_COUNT_CLASS}`);
      const countText = hoverCount?.textContent?.trim() || "";
      if (!countText) {
        return;
      }

      const related = event.relatedTarget instanceof Element ? event.relatedTarget : null;
      if (related && blockHeaderDayCell.contains(related)) {
        return;
      }

      const day = parseDayLabel(blockHeaderDayCell.getAttribute(DAY_ATTR) || blockHeaderDayCell.textContent);
      const headerRow = blockHeaderDayCell.closest(".day-row.group");
      if (day === null || !headerRow) {
        return;
      }

      if (
        activeBlockDayHover
        && activeBlockDayHover.day === day
        && activeBlockDayHover.headerRow === headerRow
        && activeBlockDayHover.hoveredHeaderCell === blockHeaderDayCell
      ) {
        return;
      }

      if (blockDayHoverTimerId !== null) {
        window.clearTimeout(blockDayHoverTimerId);
      }

      pendingBlockDayHover = null;
      clearBlockFocus(root);
      clearGlobalHeaderDayFocus(root);
      applyBlockDayFocus(root, headerRow, blockHeaderDayCell, day);
    }, true);

    root.addEventListener("mouseout", (event) => {
      const targetElement = getEventTargetElement(event);
      const fromCountNode = targetElement?.closest(`.${BLOCK_DAY_COUNT_CLASS}`) || null;
      const fromBlockHeaderCell = targetElement?.closest("#right-body .day-row.group .day-cell") || null;

      if (!fromCountNode && !fromBlockHeaderCell) {
        return;
      }

      const related = event.relatedTarget instanceof Element ? event.relatedTarget : null;
      if (related && ((fromCountNode && fromCountNode.contains(related)) || (fromBlockHeaderCell && fromBlockHeaderCell.contains(related)))) {
        return;
      }

      clearBlockDayFocus(root);
    }, true);

    root.addEventListener("mouseover", (event) => {
      if (!activeBlockDayHover) {
        return;
      }

      const targetElement = getEventTargetElement(event);
      if (!targetElement) {
        clearBlockDayFocus(root);
        return;
      }

      const hoveredCountNode = targetElement.closest(`.${BLOCK_DAY_COUNT_CLASS}`);
      const hoveredBlockHeaderCell = targetElement.closest("#right-body .day-row.group .day-cell");
      const isInsideActiveHover = Boolean(
        hoveredCountNode
        && activeBlockDayHover.hoveredHeaderCell
        && activeBlockDayHover.hoveredHeaderCell.contains(hoveredCountNode),
      ) || Boolean(
        hoveredBlockHeaderCell
        && activeBlockDayHover.hoveredHeaderCell === hoveredBlockHeaderCell,
      );

      if (!isInsideActiveHover) {
        clearBlockDayFocus(root);
      }
    }, true);

    root.addEventListener("mousedown", (event) => {
      if (event.button !== 0) {
        return;
      }

      const targetCell = getRangeCellFromEvent(event);
      if (!targetCell) {
        return;
      }

      isRangePressActive = true;

      clearBlockDayFocus(root);

      const dayRow = targetCell.closest(".day-row");
      if (!dayRow || dayRow.classList.contains("group")) {
        return;
      }

      if (activeFocusRow === dayRow || pendingHoverFocusRow === dayRow) {
        return;
      }

      if (hoverFocusTimerId !== null) {
        window.clearTimeout(hoverFocusTimerId);
      }

      pendingHoverFocusRow = null;
      applyBlockFocus(root, dayRow);
    }, true);

    root.addEventListener("mouseover", (event) => {
      if (!isRangePressActive) {
        return;
      }

      const targetCell = getRangeCellFromEvent(event);
      if (!targetCell) {
        return;
      }

      const dayRow = targetCell.closest(".day-row");
      if (!dayRow || dayRow.classList.contains("group") || activeFocusRow === dayRow) {
        return;
      }

      applyBlockFocus(root, dayRow);
    }, true);

    const handleRangePressRelease = () => {
      if (!isRangePressActive) {
        return;
      }

      isRangePressActive = false;
      clearBlockFocus(root);
    };

    root.addEventListener("mouseup", handleRangePressRelease, true);
    window.addEventListener("mouseup", handleRangePressRelease, true);
    root.addEventListener("mouseleave", handleRangePressRelease, true);

    root.addEventListener("mouseover", (event) => {
      const headerCell = event.target instanceof Element ? event.target.closest("#right-header-track .day-cell") : null;
      if (!headerCell) {
        return;
      }

      const day = parseDayLabel(headerCell.getAttribute(DAY_ATTR) || headerCell.textContent);
      if (day === null) {
        return;
      }

      if (activeHeaderHoverDay === day || pendingHeaderHoverDay === day) {
        return;
      }

      if (headerHoverTimerId !== null) {
        window.clearTimeout(headerHoverTimerId);
      }

      pendingHeaderHoverDay = day;
      headerHoverTimerId = window.setTimeout(() => {
        headerHoverTimerId = null;
        const dayToFocus = pendingHeaderHoverDay;
        pendingHeaderHoverDay = null;
        if (!Number.isInteger(dayToFocus)) {
          return;
        }

        clearBlockFocus(root);
        clearBlockDayFocus(root);
        applyGlobalHeaderDayBand(root, headerCell);
        applyGlobalHeaderDayFocus(root, dayToFocus);
      }, BAR_HOVER_FOCUS_DELAY_MS);
    }, true);

    root.addEventListener("mouseout", (event) => {
      const targetElement = getEventTargetElement(event);
      const fromHeaderCell = targetElement?.closest("#right-header-track .day-cell") || null;
      if (!fromHeaderCell) {
        return;
      }

      const related = event.relatedTarget instanceof Element ? event.relatedTarget : null;
      if (related?.closest?.("#right-header-track .day-cell")) {
        return;
      }

      clearGlobalHeaderDayFocus(root);
    }, true);

    root.addEventListener("scroll", () => {
      if (activeHeaderHoverDay === null && !activeBlockDayHover) {
        return;
      }

      if (activeHeaderHoverDay !== null) {
        updateGlobalHeaderDayBandPosition(root);
      }

      if (activeBlockDayHover) {
        updateBlockDayBandPosition(root);
      }
    }, true);

    window.addEventListener("resize", () => {
      if (activeHeaderHoverDay === null && !activeBlockDayHover) {
        return;
      }

      if (activeHeaderHoverDay !== null) {
        updateGlobalHeaderDayBandPosition(root);
      }

      if (activeBlockDayHover) {
        updateBlockDayBandPosition(root);
      }
    });
  }

  function isDataRow(dayRow, leftRow) {
    if (!dayRow || dayRow.classList.contains("group")) {
      return false;
    }

    if (!leftRow || leftRow.classList.contains("group")) {
      return false;
    }

    const hasListoCheckbox = leftRow.querySelector('input[type="checkbox"].listo-checkbox');
    const hasEditableTitleCell = leftRow.querySelector('.title-cell[data-column-key="title"], .title-cell__input, .title-cell__text');
    return Boolean(hasListoCheckbox || hasEditableTitleCell);
  }

  function scheduleMark(root) {
    if (rafId !== null) {
      return;
    }

    rafId = window.requestAnimationFrame(() => {
      rafId = null;
      markCalendarCells(root);
    });
  }

  function startCalendarMarkObserver(root) {
    const targetNode = root.querySelector(OBSERVER_TARGET_SELECTOR) || root.querySelector("#right-body");
    if (!targetNode) {
      return;
    }

    if (observer) {
      observer.disconnect();
    }

    observer = new MutationObserver((mutationList) => {
      const hasStructuralChange = mutationList.some((mutation) => mutation.type === "childList");
      if (!hasStructuralChange) {
        return;
      }

      scheduleMark(root);
    });

    observer.observe(targetNode, {
      childList: true,
      subtree: true,
    });
  }

  function attachDateEditRepaintListeners(root) {
    const flushPendingRepaint = () => {
      repaintRafId = null;
      repaintRafId2 = null;
      repaintTimeoutId = null;

      if (pendingFullRepaint) {
        pendingRowsToRepaint.clear();
        pendingFullRepaint = false;
        markCalendarCells(root);
        return;
      }

      const rowsToRepaint = [...pendingRowsToRepaint];
      pendingRowsToRepaint.clear();
      rowsToRepaint.forEach((leftRow) => {
        paintSingleRowByLeftRow(root, leftRow);
      });
    };

    const schedulePostUpdateRepaint = ({ leftRow = null, full = false } = {}) => {
      if (full) {
        pendingFullRepaint = true;
      } else if (leftRow) {
        pendingRowsToRepaint.add(leftRow);
      } else {
        pendingFullRepaint = true;
      }

      if (repaintRafId !== null || repaintRafId2 !== null || repaintTimeoutId !== null) {
        return;
      }

      repaintRafId = window.requestAnimationFrame(() => {
        repaintRafId2 = window.requestAnimationFrame(() => {
          repaintTimeoutId = window.setTimeout(flushPendingRepaint, 0);
        });
      });
    };

    const isDateColumnEvent = (event) => {
      const targetCell = event.target?.closest?.(DATE_COLUMN_SELECTOR);
      return targetCell || null;
    };

    const repaintFromEvent = (event) => {
      const targetCell = isDateColumnEvent(event);
      if (!targetCell) {
        return;
      }

      const leftRow = targetCell.closest(".left-row");
      if (!leftRow) {
        return;
      }

      schedulePostUpdateRepaint({ leftRow });
    };

    root.addEventListener("input", repaintFromEvent, true);
    root.addEventListener("change", repaintFromEvent, true);
    root.addEventListener("focusout", repaintFromEvent, true);
    root.addEventListener("paste", (event) => {
      const eventTarget = event.target instanceof Element ? event.target : null;
      const pastedInsideGantt = eventTarget?.closest?.(GANTT_BODY_SELECTOR);
      if (!pastedInsideGantt) {
        return;
      }

      schedulePostUpdateRepaint({ full: true });
    }, true);

    const handleDocumentPaste = (event) => {
      const eventTarget = event.target instanceof Element ? event.target : null;
      const targetInsideMonthBlock = eventTarget?.closest?.(".month-block") || root.contains(eventTarget);
      if (!targetInsideMonthBlock) {
        return;
      }

      repaintUntilStable(root, { maxMs: 600 });
    };

    document.addEventListener("paste", handleDocumentPaste, true);
    document.addEventListener("beforeinput", (event) => {
      if (event.inputType !== "insertFromPaste") {
        return;
      }

      handleDocumentPaste(event);
    }, true);

    root.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") {
        return;
      }
      
      repaintFromEvent(event);
    }, true);
  }

  function run() {
    const monthBlock = document.querySelector(".month-block");
    if (!monthBlock) {
      return;
    }

    markCalendarCells(monthBlock);
    startCalendarMarkObserver(monthBlock);
    attachDateEditRepaintListeners(monthBlock);
    attachHoverFocusListeners(monthBlock);

    document.addEventListener("calendar:month-change", (event) => {
      const target = event.target;
      const isDocumentTarget = target === document;
      const isRelatedElement = target instanceof Element
        && (target.contains(monthBlock) || monthBlock.contains(target));

      if (!isDocumentTarget && !isRelatedElement) {
        return;
      }

      markCalendarCells(monthBlock);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", run, { once: true });
  } else {
    run();
  }
})();
