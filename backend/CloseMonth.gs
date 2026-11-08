/**
 * Personal finances: "Close month"
 *
 * Replace the old version of this file in Extensions > Apps Script with this one, and save.
 * Two ways to close a month:
 *   - On any device: tick the "Close month" box in Recurring (cell G3).
 *   - On a computer: menu Finances > Close month (asks for confirmation).
 *
 * What it does:
 *   1. Copies every ticked recurring item into Transactions, dated in the open month,
 *      marked as paid. Balances do not change: the money just moves from one list to the other.
 *   2. Clears all the ticks in Recurring.
 *   3. Moves the open month (Recurring B3) and the Dashboard month to the next month.
 * Unticked items are not copied: they did not happen.
 */

const OPEN_MONTH_CELL = 'B3';
const CLOSE_BOX_CELL = 'G3';
const ITEMS_RANGE = 'A5:G39';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Finances')
    .addItem('Close month', 'closeMonthWithConfirm')
    .addToUi();
}

// Runs on every edit. Only reacts when the "Close month" box is ticked.
function onEdit(e) {
  const range = e.range;
  if (range.getSheet().getName() !== 'Recurring') return;
  if (range.getA1Notation() !== CLOSE_BOX_CELL) return;
  if (String(e.value).toUpperCase() !== 'TRUE') return;
  try {
    const n = closeMonth();
    e.source.toast(n + ' items saved in Transactions. New month open.', 'Month closed', 6);
  } finally {
    range.setValue(false);
  }
}

function closeMonthWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const label = openMonthLabel_();
  const answer = ui.alert(
    'Close ' + label + '?',
    'Ticked recurring items will be saved in Transactions and all ticks will be cleared.',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer !== ui.Button.OK) return;
  const n = closeMonth();
  ui.alert(label + ' closed', n + ' items saved in Transactions.', ui.ButtonSet.OK);
}

function closeMonth() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const rec = ss.getSheetByName('Recurring');
  const tx = ss.getSheetByName('Transactions');

  // Open month, read in the spreadsheet's time zone so the day never shifts.
  const open = rec.getRange(OPEN_MONTH_CELL).getValue();
  const year = Number(Utilities.formatDate(open, tz, 'yyyy'));
  const month = Number(Utilities.formatDate(open, tz, 'M')); // 1-12
  const daysInMonth = new Date(year, month, 0).getDate();

  // Ticked items become paid transactions.
  const rows = rec.getRange(ITEMS_RANGE).getValues();
  const out = [];
  rows.forEach(function (r) {
    const description = r[0], paid = r[6];
    if (!description || paid !== true) return;
    let day = 1;
    if (r[5] instanceof Date) day = Number(Utilities.formatDate(r[5], tz, 'd'));
    else if (Number(r[5]) > 0) day = Number(r[5]);
    day = Math.min(day, daysInMonth);
    out.push([
      '=DATE(' + year + ',' + month + ',' + day + ')', // date as a formula: no time-zone surprises
      description, // A Description
      r[1],        // Category
      r[2],        // From
      r[3],        // To
      r[4],        // Amount
      true         // Paid
    ]);
  });

  if (out.length) {
    const colA = tx.getRange('A2:A').getValues();
    let lastRow = 1;
    for (let i = 0; i < colA.length; i++) if (colA[i][0] !== '') lastRow = i + 2;
    tx.getRange(lastRow + 1, 1, out.length, 7).setValues(out);
  }

  // Clear the ticks (only on rows that have an item).
  const ticks = rows.map(function (r) { return [r[0] ? false : r[6]]; });
  rec.getRange('G5:G39').setValues(ticks);

  // Move to the next month.
  const nextYear = month === 12 ? year + 1 : year;
  const nextMonth = month === 12 ? 1 : month + 1;
  const nextFormula = '=DATE(' + nextYear + ',' + nextMonth + ',1)';
  rec.getRange(OPEN_MONTH_CELL).setFormula(nextFormula);
  ss.getSheetByName('Dashboard').getRange('B2').setFormula(nextFormula);
  SpreadsheetApp.flush();
  rollWeekly_(nextYear, nextMonth);

  return out.length;
}

// Weekly items ("Name (week 1)", "Name (week 2)"…) keep their weekday in the new month.
// A month with 5 of that weekday gets a 5th row; a month with 4 loses it.
function rollWeekly_(y, m) {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const rec = ss.getSheetByName('Recurring');
  const cols = Math.min(17, rec.getMaxColumns());
  const vals = rec.getRange(5, 1, 35, cols).getValues();
  const groups = {}, order = [];
  vals.forEach(function (r, i) {
    const mm = String(r[0]).match(/^(.*) \(week (\d+)\)$/);
    if (!mm) return;
    if (!groups[mm[1]]) { groups[mm[1]] = []; order.push(mm[1]); }
    groups[mm[1]].push({ i: i, r: r, n: Number(mm[2]) });
  });
  const last = new Date(y, m, 0).getDate();
  order.forEach(function (base) {
    const g = groups[base].sort(function (a, b) { return a.n - b.n; });
    const f = g[0].r[5];
    if (!(f instanceof Date)) return;
    const wd = Number(Utilities.formatDate(f, tz, 'u')) % 7; // 0 = Sunday
    const days = [];
    for (let d = 1; d <= last; d++) if (new Date(y, m - 1, d).getDay() === wd) days.push(d);
    g.forEach(function (x, k) {
      const row = 5 + x.i;
      if (k < days.length) {
        rec.getRange(row, 1).setValue(base + ' (week ' + (k + 1) + ')');
        rec.getRange(row, 6).setFormula('=DATE(' + y + ',' + m + ',' + days[k] + ')');
      } else {
        rec.getRange(row, 1, 1, 6).clearContent();
        rec.getRange(row, 7).setValue(false);
      }
    });
    for (let k = g.length; k < days.length; k++) {
      const col = rec.getRange(5, 1, 35, 1).getValues();
      let free = -1;
      for (let i = 0; i < col.length; i++) if (col[i][0] === '') { free = 5 + i; break; }
      if (free < 0) break;
      const src = g[0].r;
      rec.getRange(free, 1, 1, 6).setValues([[base + ' (week ' + (k + 1) + ')', src[1], src[2], src[3], src[4], '=DATE(' + y + ',' + m + ',' + days[k] + ')']]);
      rec.getRange(free, 7).setValue(false);
      if (cols >= 17) rec.getRange(free, 17).setValue(src[16] === true);
    }
  });
}

function openMonthLabel_() {
  const ss = SpreadsheetApp.getActive();
  const open = ss.getSheetByName('Recurring').getRange(OPEN_MONTH_CELL).getValue();
  return Utilities.formatDate(open, ss.getSpreadsheetTimeZone(), 'MMMM yyyy');
}
