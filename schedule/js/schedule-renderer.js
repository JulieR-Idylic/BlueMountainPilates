"use strict";

document.addEventListener("DOMContentLoaded", () => {
  loadSchedule();
});


async function loadSchedule() {
  const container = document.getElementById("live-schedule");

  if (!container) {
    console.error(
      "Schedule renderer could not find #live-schedule."
    );
    return;
  }

  try {
    const response = await fetch(
      "data/schedule.json",
      {
        cache: "no-store"
      }
    );

    if (!response.ok) {
      throw new Error(
        `Schedule JSON request failed: ${response.status}`
      );
    }

    const schedule = await response.json();

    validateSchedule(schedule);

    renderSchedule(
      container,
      schedule
    );
  }
  catch (error) {
    console.error(
      "Unable to load live schedule:",
      error
    );

    renderLoadError(container);
  }
}


function validateSchedule(schedule) {
  if (!schedule) {
    throw new Error(
      "Schedule data is empty."
    );
  }

  if (schedule.schemaVersion !== 2) {
    throw new Error(
      `Unsupported schedule schema version: ${schedule.schemaVersion}`
    );
  }

  if (
    !Array.isArray(schedule.weeks) ||
    schedule.weeks.length === 0
  ) {
    throw new Error(
      "Schedule contains no weeks."
    );
  }
}


function renderSchedule(
  container,
  schedule
) {
  container.innerHTML = "";

  window.schedulePublishedAt =
    schedule.publishedAt ?? null;

  schedule.weeks.forEach(
    (week, index) => {
      const section = buildWeekSection(
        week,
        index
      );

      container.appendChild(section);
    }
  );
}


function buildWeekSection(
  week,
  index
) {
  const section = document.createElement(
    "section"
  );

  section.className = "schedule-week";

  const heading = document.createElement(
    "h2"
  );

  heading.className = "schedule-week-heading";
  heading.textContent = getWeekLabel(
    week.key,
    index
  );

  section.appendChild(heading);

  if (window.schedulePublishedAt) {
    const updated = document.createElement(
      "p"
    );

    updated.className =
      "schedule-last-updated";

    updated.textContent =
      `Schedule last updated: ${formatPublishedTime(
        window.schedulePublishedAt
      )}`;

    section.appendChild(updated);
  }

  const frame = document.createElement(
    "div"
  );

  frame.className = "schedule-table-frame";

  const table = document.createElement(
    "table"
  );

  table.className = "schedule-grid";

  table.setAttribute(
    "aria-label",
    heading.textContent
  );

  addColumnWidths(
    table,
    week.columnWidths
  );

  const tbody = document.createElement(
    "tbody"
  );

  week.rows.forEach(
    (row, rowIndex) => {
      const tr = buildRow(
        row,
        week.rowHeights?.[rowIndex]
      );

      tbody.appendChild(tr);
    }
  );

  table.appendChild(tbody);
  frame.appendChild(table);
  section.appendChild(frame);

  return section;
}


function buildRow(
  row,
  rowHeight
) {
  const tr = document.createElement(
    "tr"
  );

  tr.classList.add(
    "schedule-row"
  );

  if (row.tone) {
    tr.classList.add(
      `schedule-tone-${row.tone}`
    );
  }

  if (
    rowHeight !== null &&
    rowHeight !== undefined
  ) {
    tr.style.height =
      excelPointsToPixels(rowHeight);
  }

  row.cells.forEach(
    (cell) => {
      const td = buildCell(cell);
      tr.appendChild(td);
    }
  );

  return tr;
}


function buildCell(cell) {
  const td = document.createElement(
    "td"
  );

  td.classList.add(
    "schedule-cell"
  );

  if (cell.style) {
    td.classList.add(
      `schedule-${cell.style}`
    );
  }

  if (cell.bold) {
    td.classList.add(
      "schedule-bold"
    );
  }

  if (
    cell.rowspan &&
    cell.rowspan > 1
  ) {
    td.rowSpan = cell.rowspan;
  }

  if (
    cell.colspan &&
    cell.colspan > 1
  ) {
    td.colSpan = cell.colspan;
  }

  td.textContent =
    cell.text ?? "";

  return td;
}


function addColumnWidths(
  table,
  columnWidths
) {
  if (
    !Array.isArray(columnWidths) ||
    columnWidths.length === 0
  ) {
    return;
  }

  const colgroup =
    document.createElement(
      "colgroup"
    );

  columnWidths.forEach(
    (excelWidth) => {
      const col =
        document.createElement(
          "col"
        );

      col.style.width =
        excelColumnWidthToPixels(
          excelWidth
        );

      colgroup.appendChild(col);
    }
  );

  table.appendChild(colgroup);
}


function getWeekLabel(
  key,
  index
) {
  if (key === "Web_ThisWeek") {
    return "This Week";
  }

  if (key === "Web_NextWeek") {
    return "Next Week";
  }

  const plusMatch =
    /^Web_WeekPlus(\d+)$/.exec(key);

  if (plusMatch) {
    const offset =
      Number(plusMatch[1]);

    return `${offset} Weeks Ahead`;
  }

  return `Week ${index + 1}`;
}


function formatPublishedTime(
  value
) {
  if (!value) {
    return "";
  }

  const date = new Date(value);

  return new Intl.DateTimeFormat(
    "en-US",
    {
      month: "short",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit",
      timeZone: "America/Los_Angeles",
    }
  ).format(date);
}


function excelColumnWidthToPixels(
  excelWidth
) {
  const width =
    Number(excelWidth);

  if (!Number.isFinite(width)) {
    return "90px";
  }

  const pixels =
    Math.round(
      (width * 7) + 5
    );

  return `${pixels}px`;
}


function excelPointsToPixels(
  points
) {
  const value =
    Number(points);

  if (!Number.isFinite(value)) {
    return "";
  }

  const pixels =
    Math.round(
      value * 96 / 72
    );

  return `${pixels}px`;
}


function renderLoadError(
  container
) {
  container.innerHTML = "";

  const message =
    document.createElement(
      "p"
    );

  message.className =
    "schedule-load-error";

  message.textContent =
    "The live schedule is temporarily unavailable. Please try again shortly.";

  container.appendChild(message);
}