/**
 * Sheet1 -> website tab auto-sync (Google Apps Script)
 *
 * Sheet1 me row bharo (A:J, same columns jaise baaki tabs me hain). Jaise hi row complete hoti hai,
 * wo us website ke tab me TOP par (header ke neeche) insert ho jaati hai; purana data neeche khisak jaata hai.
 * Sheet1 ki column K me "Synced" likha aata hai taaki dobara copy na ho.
 */

const SOURCE_SHEET = 'Sheet1';
const HEADER_LABEL = 'Month';          // tab ki header row pehchaanne ke liye (column A)
const WEBSITE_COL = 2;                 // B = Original Website Tag
const STATUS_COL = 11;                 // K = Sync Status (Sheet1 me)
const DATA_COLS = 10;                  // A:J copy hota hai
const REQUIRED_COLS = [1, 2, 3, 4];    // Month, Website, Project, Status bhare hon tabhi sync hoga
const LAST_UPDATED_COL = 10;           // J — khaali ho to aaj ki date khud bhar dega

// Jahan "Website Tag" ka naam aur tab ka naam alag hai, wahan yahan likho:  'tag (lowercase)': 'Tab Name'
const ALIASES = {
  'happy robot': 'Happy Robot',
  'customer support': 'Customer support',
  'all stores - seo': 'SEO  SEM',
  'seo / sem all stores': 'SEO  SEM',
  'samsung parts': 'Samsung ',
  'pim data': 'PIM',
};

/** Simple trigger — koi setup nahi chahiye. Sheet1 me edit/paste hote hi chalta hai. */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SOURCE_SHEET) return;
  const first = Math.max(e.range.getRow(), 2);
  const last = e.range.getLastRow();
  if (last >= first) syncRows_(sh, first, last);
}

/** Menu se manually saari pending rows sync karne ke liye. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Task Sync')
    .addItem('Sync pending rows now', 'syncAllPending')
    .addToUi();
}

function syncAllPending() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SOURCE_SHEET);
  if (sh && sh.getLastRow() >= 2) syncRows_(sh, 2, sh.getLastRow());
}

function syncRows_(src, from, to) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const ss = src.getParent();
    const rows = src.getRange(from, 1, to - from + 1, STATUS_COL).getValues();
    // Upar se neeche: Sheet1 ki sabse neeche wali (latest) row sabse upar pahunchti hai.
    rows.forEach((row, i) => {
      const r = from + i;
      if (String(row[STATUS_COL - 1]).indexOf('Synced') === 0) return;
      if (!REQUIRED_COLS.every(c => String(row[c - 1]).trim() !== '')) return;

      const tab = findTab_(ss, String(row[WEBSITE_COL - 1]));
      const cell = src.getRange(r, STATUS_COL);
      if (!tab) { cell.setValue('Tab not found: ' + row[WEBSITE_COL - 1]); return; }

      const values = row.slice(0, DATA_COLS);
      if (String(values[LAST_UPDATED_COL - 1]).trim() === '') values[LAST_UPDATED_COL - 1] = new Date();
      insertAtTop_(tab, values);
      cell.setValue('Synced ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd MMM HH:mm') + ' -> ' + tab.getName());
    });
  } finally {
    lock.releaseLock();
  }
}

function findTab_(ss, tag) {
  const key = tag.trim().toLowerCase();
  const name = ALIASES[key];
  if (name && ss.getSheetByName(name)) return ss.getSheetByName(name);
  return ss.getSheets().find(s => s.getName().trim().toLowerCase() === key) || null;
}

function insertAtTop_(tab, values) {
  const colA = tab.getRange(1, 1, Math.min(tab.getMaxRows(), 20), 1).getValues();
  const hdr = colA.findIndex(r => String(r[0]).trim() === HEADER_LABEL) + 1 || 4;
  const top = hdr + 1;
  const width = Math.max(tab.getLastColumn(), DATA_COLS);
  // Row `top` ke neeche nayi row daalte hain (taaki "Total tasks" waali ranges khud expand hon),
  // phir purani top row ko ek neeche copy karke nayi entry top par likhte hain.
  tab.insertRowAfter(top);
  tab.getRange(top, 1, 1, width).copyTo(tab.getRange(top + 1, 1, 1, width));
  tab.getRange(top, 1, 1, DATA_COLS).setValues([values]);
}
