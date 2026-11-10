/**
 * Fyn app: connection between the phone app and this spreadsheet.
 *
 * Install (once):
 *   1. Extensions > Apps Script. Click + next to Files > Script, name it "Api", paste this whole file, save.
 *   2. Run the function "createAppKey" (select it at the top, click Run, allow access).
 *      A window shows your secret key. Copy it.
 *   3. Deploy > New deployment > type "Web app".
 *      Execute as: Me. Who has access: Anyone. Deploy, copy the Web app URL.
 *   4. Open the Fyn app on your phone and paste the URL and the key.
 *
 * Without the key nobody can read or change anything. To block a lost phone,
 * run "createAppKey" again: the old key stops working at once.
 */

const SALDO = {
  KEY_PROP: 'SALDO_KEY',
  TX: 'Transactions', REC: 'Recurring', ACC: 'Accounts', CAT: 'Categories', BUD: 'Budget', DASH: 'Dashboard',
  REC_FIRST: 5, REC_LAST: 39,
  NOTE_COL: 10 // J in Transactions
};

function createAppKey() {
  const key = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
  PropertiesService.getScriptProperties().setProperty(SALDO.KEY_PROP, key);
  ensureNoteColumn_();
  try {
    SpreadsheetApp.getUi().alert('Your Fyn key', key + '\n\nPaste it in the app. Keep it private. Running this again makes a new key and blocks the old one.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) { Logger.log('Fyn key: ' + key); }
  return key;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const out = run_(() => {
    checkKey_(p.key);
    if (p.payload) return runAction_(JSON.parse(p.payload));
    return readAll_();
  });
  if (p.frame) return frame_(out, p.frame);
  return emit_(out, p.callback);
}

// Second way in: the phone opens this in a hidden frame and gets the answer by message.
function frame_(out, id) {
  const safeId = /^[A-Za-z0-9_]{1,40}$/.test(id) ? id : 'x';
  const s = JSON.stringify(out).replace(/</g, '\\u003c');
  const html = '<!DOCTYPE html><html><body><script>window.top.postMessage({fyn:"' + safeId + '",out:' + s + '},"*");<\/script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) {}
  return emit_(run_(() => { checkKey_(body.key); return runAction_(body); }));
}

function runAction_(body) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const fn = ACTIONS_[body.action];
    if (!fn) throw new Error('Unknown action');
    const result = fn(body) || {};
    SpreadsheetApp.flush();
    result.data = readAll_();
    return result;
  } finally { lock.releaseLock(); }
}

function run_(fn) {
  try { return { ok: true, result: fn() }; }
  catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}

// The phone asks with ?callback=name so Safari never blocks the answer.
function emit_(out, cb) {
  const s = JSON.stringify(out);
  if (cb && /^[A-Za-z_$][\w$]{0,40}$/.test(cb)) return ContentService.createTextOutput(cb + '(' + s + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(key) {
  const real = PropertiesService.getScriptProperties().getProperty(SALDO.KEY_PROP);
  if (!real || !key || key !== real) throw new Error('Wrong key');
}

// ---------- reading ----------
function ss_() { return SpreadsheetApp.getActive(); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function ymd_(d) { return d instanceof Date ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : ''; }
function num_(v) { return typeof v === 'number' ? v : (v === '' || v == null ? null : Number(String(v).replace(/[^0-9.\-]/g, '')) || null); }

function readAll_() {
  const ss = ss_();
  const tx = ss.getSheetByName(SALDO.TX), rec = ss.getSheetByName(SALDO.REC), acc = ss.getSheetByName(SALDO.ACC);
  const cat = ss.getSheetByName(SALDO.CAT), bud = ss.getSheetByName(SALDO.BUD), dash = ss.getSheetByName(SALDO.DASH);

  const txLast = Math.max(2, tx.getLastRow());
  const txCols = Math.max(9, Math.min(tx.getMaxColumns(), SALDO.NOTE_COL));
  const txVals = tx.getRange(2, 1, txLast - 1, txCols).getValues();
  const txs = [];
  txVals.forEach((r, i) => {
    if (r[1] === '' && r[5] === '') return;
    txs.push({ row: i + 2, date: ymd_(r[0]), desc: String(r[1]), cat: String(r[2]), from: String(r[3]), to: String(r[4]), amount: num_(r[5]) || 0, paid: r[6] === true, note: txCols >= SALDO.NOTE_COL ? String(r[SALDO.NOTE_COL - 1] || '') : '' });
  });

  const open = rec.getRange('B3').getValue();
  const recVals = rec.getRange(SALDO.REC_FIRST, 1, SALDO.REC_LAST - SALDO.REC_FIRST + 1, Math.min(17, rec.getMaxColumns())).getValues();
  const recs = [];
  recVals.forEach((r, i) => {
    if (!r[0]) return;
    let day = 1;
    if (r[5] instanceof Date) day = Number(Utilities.formatDate(r[5], tz_(), 'd'));
    else if (Number(r[5]) > 0) day = Number(r[5]);
    recs.push({ row: SALDO.REC_FIRST + i, desc: String(r[0]), cat: String(r[1]), from: String(r[2]), to: String(r[3]), amount: num_(r[4]) || 0, day: day, paid: r[6] === true, varies: r[16] === true });
  });

  const accLast = Math.max(2, acc.getLastRow());
  const accVals = acc.getRange(2, 1, accLast - 1, 14).getValues();
  const accounts = [];
  accVals.forEach((r, i) => {
    if (!r[0]) return;
    accounts.push({ row: i + 2, name: String(r[0]), type: String(r[1]), balanceApp: num_(r[2]), balanceDate: ymd_(r[3]), balanceNow: num_(r[4]) || 0, expectedEnd: num_(r[5]), limit: num_(r[7]), available: num_(r[8]), invested: num_(r[9]), profit: num_(r[11]), note: String(r[12] || ''), colour: String(r[13] || '') });
  });

  const catVals = cat.getRange(2, 1, Math.max(1, cat.getLastRow() - 1), 3).getValues();
  const cats = catVals.filter((r) => r[0]).map((r) => ({ name: String(r[0]), group: String(r[1]), budget: num_(r[2]) }));

  const pct = (v) => (typeof v === 'number' ? (v > 1 ? v / 100 : v) : null);
  const targets = { needs: pct(bud.getRange('B4').getValue()), wants: pct(bud.getRange('B5').getValue()), savings: pct(bud.getRange('B6').getValue()) };

  return {
    today: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'),
    openMonth: ymd_(open),
    cashNow: num_(dash.getRange('B4').getValue()),
    expectedEnd: num_(dash.getRange('B5').getValue()),
    netWorth: num_(dash.getRange('B8').getValue()),
    accounts: accounts, tx: txs, rec: recs, cats: cats, targets: targets,
    sheetUrl: ss.getUrl()
  };
}

// ---------- writing ----------
function dateFormula_(s) { const p = String(s).split('-').map(Number); return '=DATE(' + p[0] + ',' + p[1] + ',' + p[2] + ')'; }
function ensureNoteColumn_() {
  const tx = ss_().getSheetByName(SALDO.TX);
  if (tx.getMaxColumns() < SALDO.NOTE_COL) tx.insertColumnsAfter(tx.getMaxColumns(), SALDO.NOTE_COL - tx.getMaxColumns());
  if (tx.getRange(1, SALDO.NOTE_COL).getValue() === '') tx.getRange(1, SALDO.NOTE_COL).setValue('Note');
}
function txRowValues_(t) {
  return [dateFormula_(t.date), t.desc || '', t.cat || '', t.from || '', t.to || '', Number(t.amount), t.paid === true];
}
function sameTx_(sheet, row, sig) {
  const r = sheet.getRange(row, 1, 1, 7).getValues()[0];
  const cur = String(r[1]) + '|' + Number(num_(r[5]) || 0).toFixed(2);
  if (sig && cur !== sig) throw new Error('This entry changed in the sheet. Refresh and try again.');
}
function firstEmptyTxRow_(tx) {
  const vals = tx.getRange(2, 1, Math.max(1, tx.getMaxRows() - 1), 6).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === '' && vals[i][1] === '' && vals[i][5] === '') return i + 2;
  tx.insertRowsAfter(tx.getMaxRows(), 50);
  return vals.length + 2;
}
function firstEmptyRecRow_(rec) {
  const vals = rec.getRange(SALDO.REC_FIRST, 1, SALDO.REC_LAST - SALDO.REC_FIRST + 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) if (vals[i][0] === '') return SALDO.REC_FIRST + i;
  throw new Error('Plan is full (35 items). Remove one in the sheet first.');
}

const ACTIONS_ = {
  addTx: function (b) {
    ensureNoteColumn_();
    const tx = ss_().getSheetByName(SALDO.TX);
    const row = firstEmptyTxRow_(tx);
    tx.getRange(row, 1, 1, 7).setValues([txRowValues_(b.tx)]);
    tx.getRange(row, SALDO.NOTE_COL).setValue(b.tx.note || '');
    return { row: row };
  },
  updateTx: function (b) {
    ensureNoteColumn_();
    const tx = ss_().getSheetByName(SALDO.TX);
    sameTx_(tx, b.row, b.sig);
    tx.getRange(b.row, 1, 1, 7).setValues([txRowValues_(b.tx)]);
    tx.getRange(b.row, SALDO.NOTE_COL).setValue(b.tx.note || '');
    return { row: b.row };
  },
  deleteTx: function (b) {
    const tx = ss_().getSheetByName(SALDO.TX);
    sameTx_(tx, b.row, b.sig);
    const old = tx.getRange(b.row, 1, 1, 7).getValues()[0];
    const note = tx.getMaxColumns() >= SALDO.NOTE_COL ? tx.getRange(b.row, SALDO.NOTE_COL).getValue() : '';
    tx.getRange(b.row, 1, 1, 6).clearContent();
    tx.getRange(b.row, 7).setValue(false);
    if (tx.getMaxColumns() >= SALDO.NOTE_COL) tx.getRange(b.row, SALDO.NOTE_COL).clearContent();
    return { removed: { date: ymd_(old[0]), desc: old[1], cat: old[2], from: old[3], to: old[4], amount: num_(old[5]), paid: old[6] === true, note: note } };
  },
  setRecPaid: function (b) {
    const rec = ss_().getSheetByName(SALDO.REC);
    if (String(rec.getRange(b.row, 1).getValue()) !== b.desc) throw new Error('Plan changed in the sheet. Refresh and try again.');
    if (b.amount != null) rec.getRange(b.row, 5).setValue(Number(b.amount));
    rec.getRange(b.row, 7).setValue(b.paid === true);
    return {};
  },
  saveRec: function (b) {
    // b.rows: [{row|null, desc, cat, from, to, amount, day}], b.remove: [row]
    const rec = ss_().getSheetByName(SALDO.REC);
    const open = rec.getRange('B3').getValue();
    const y = Number(Utilities.formatDate(open, tz_(), 'yyyy')), m = Number(Utilities.formatDate(open, tz_(), 'M'));
    if (rec.getMaxColumns() < 17) rec.insertColumnsAfter(rec.getMaxColumns(), 17 - rec.getMaxColumns());
    if (rec.getRange(4, 17).getValue() === '') rec.getRange(4, 17).setValue('Changes each time');
    (b.remove || []).slice().sort((p, q) => q - p).forEach((row) => { rec.getRange(row, 1, 1, 6).clearContent(); rec.getRange(row, 7).setValue(false); rec.getRange(row, 17).clearContent(); });
    (b.rows || []).forEach((r) => {
      const row = r.row || firstEmptyRecRow_(rec);
      const day = Math.min(Number(r.day) || 1, new Date(y, m, 0).getDate());
      rec.getRange(row, 1, 1, 6).setValues([[r.desc, r.cat || '', r.from || '', r.to || '', Number(r.amount), '=DATE(' + y + ',' + m + ',' + day + ')']]);
      rec.getRange(row, 17).setValue(r.varies === true);
      if (!r.row) rec.getRange(row, 7).setValue(r.paid === true);
    });
    return {};
  },
  setBalance: function (b) {
    const acc = ss_().getSheetByName(SALDO.ACC);
    if (String(acc.getRange(b.row, 1).getValue()) !== b.name) throw new Error('Accounts changed in the sheet. Refresh and try again.');
    acc.getRange(b.row, 3).setValue(Number(b.value));
    const t = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    acc.getRange(b.row, 4).setFormula(dateFormula_(t));
    return {};
  },
  addAccount: function (b) {
    const acc = ss_().getSheetByName(SALDO.ACC);
    const name = String(b.name || '').trim();
    if (!name) throw new Error('Give the account a name.');
    const names = acc.getRange(2, 1, 41, 1).getValues().map((r) => String(r[0]).trim().toLowerCase());
    if (names.indexOf(name.toLowerCase()) >= 0) throw new Error('There is already an account called ' + name + '.');
    let row = -1;
    for (let i = 0; i < names.length; i++) if (names[i] === '') { row = i + 2; break; }
    if (row < 0) throw new Error('Accounts is full (41). Remove one in the sheet first.');
    const types = ['Bank', 'Cash', 'Credit card', 'Investment', 'Crypto', 'Loan', 'Property'];
    const type = types.indexOf(b.type) >= 0 ? b.type : 'Bank';
    const owed = type === 'Credit card' || type === 'Loan';
    const bal = Number(b.balance) || 0;
    const t = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
    acc.getRange(row, 1, 1, 2).setValues([[name, type]]);
    acc.getRange(row, 3).setValue(owed ? -Math.abs(bal) : bal);
    acc.getRange(row, 4).setFormula(dateFormula_(t));
    if (type === 'Credit card' && Number(b.limit) > 0) acc.getRange(row, 8).setValue(Number(b.limit));
    if (b.note) acc.getRange(row, 13).setValue(String(b.note));
    if (b.colour) acc.getRange(row, 14).setValue(String(b.colour));
    return { row: row };
  },
  closeMonth: function () {
    return { saved: closeMonth() }; // from the existing "Close month" script in this project
  }
};
