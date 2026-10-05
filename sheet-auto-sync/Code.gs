/**
 * Worksheet -> website tab auto-sync (Google Apps Script)
 *
 * Worksheet me row bharo (A:J, same columns jaise baaki tabs me hain). Jaise hi row complete hoti hai,
 * wo us website ke tab me TOP par (header ke neeche) insert ho jaati hai; purana data neeche khisak jaata hai.
 * Worksheet ki K me "Synced" aur L me ID aati hai; baad me A:J me koi bhi badlav tab me update hota hai.
 */

const SOURCE_SHEET = 'Worksheet';
const HEADER_LABEL = 'Month';          // tab ki header row pehchaanne ke liye (column A)
const WEBSITE_COL = 2;                 // B = Original Website Tag
const STATUS_COL = 11;                 // K = Sync Status (Worksheet me)
const ID_COL = 12;                     // L = Sync ID (script khud bharti hai, mat chhedo)
const META_KEY = 'syncId';
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

/** Installable trigger se chalta hai (installTrigger() ek baar chalao). Worksheet me edit/paste hote hi sync. */
function handleEdit(e) {
  if (e && e.range) SpreadsheetApp.getActive().toast('Task Sync chal raha hai: ' + e.range.getSheet().getName(), 'Task Sync', 3);
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName().trim().toLowerCase() !== SOURCE_SHEET.toLowerCase()) return;
  const first = Math.max(e.range.getRow(), 2);
  const last = e.range.getLastRow();
  if (last >= first) syncRows_(sh, first, last);
}

/** EK BAAR chalao (Run dabao, permission Allow karo) — edit trigger install ho jaata hai. */
function installTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'handleEdit').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

/** Menu se manually saari pending rows sync karne ke liye. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Task Sync')
    .addItem('Sync pending rows now', 'syncAllPending')
    .addToUi();
}

function syncAllPending() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheets().find(s => s.getName().trim().toLowerCase() === SOURCE_SHEET.toLowerCase());
  if (!sh) { ss.toast('Sheet "' + SOURCE_SHEET + '" nahi mili. Tab ke naam check karo.', 'Task Sync', 10); return; }
  if (sh.getLastRow() < 2) { ss.toast('Worksheet me koi data row nahi hai.', 'Task Sync', 5); return; }
  syncRows_(sh, 2, sh.getLastRow());
  ss.toast('Sync pura hua. K column dekho.', 'Task Sync', 5);
}

function syncRows_(src, from, to) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const ss = src.getParent();
    const rows = src.getRange(from, 1, to - from + 1, ID_COL).getValues();
    // Upar se neeche: Worksheet ki sabse neeche wali (latest) row sabse upar pahunchti hai.
    rows.forEach((row, i) => {
      const r = from + i;
      try {
      const values = row.slice(0, DATA_COLS);
      const synced = String(row[STATUS_COL - 1]).indexOf('Synced') === 0;
      const id = String(row[ID_COL - 1]);

      if (synced) {
        // Pehle se sync ho chuki row me baad me jo bhi column (A:J) badla, wo tab me update karo.
        const tab = findTab_(ss, String(row[WEBSITE_COL - 1]));
        if (tab && id) updateExisting_(tab, id, values);
        return;
      }
      if (!REQUIRED_COLS.every(c => String(row[c - 1]).trim() !== '')) return;

      const tab = findTab_(ss, String(row[WEBSITE_COL - 1]));
      const cell = src.getRange(r, STATUS_COL);
      if (!tab) { cell.setValue('Tab not found: ' + row[WEBSITE_COL - 1]); return; }

      if (String(values[LAST_UPDATED_COL - 1]).trim() === '') values[LAST_UPDATED_COL - 1] = new Date();
      const newId = Utilities.getUuid();
      insertAtTop_(tab, values, newId);
      src.getRange(r, ID_COL).setValue(newId);
      cell.setValue('Synced ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd MMM HH:mm') + ' -> ' + tab.getName());
      } catch (err) {
        src.getRange(r, STATUS_COL).setValue('ERROR: ' + err.message);
      }
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

function metaRows_(tab, id) {
  let f = tab.createDeveloperMetadataFinder().withKey(META_KEY);
  if (id) f = f.withValue(id);
  return f.find().filter(m => m.getLocation().getRow());
}

function insertAtTop_(tab, values, id) {
  const colA = tab.getRange(1, 1, Math.min(tab.getMaxRows(), 20), 1).getValues();
  const hdr = colA.findIndex(r => String(r[0]).trim() === HEADER_LABEL) + 1 || 4;
  const top = hdr + 1;
  const width = Math.max(tab.getLastColumn(), DATA_COLS);
  // Row `top` ke neeche nayi row daalte hain (taaki "Total tasks" waali ranges khud expand hon),
  // phir purani top row ko ek neeche copy karke nayi entry top par likhte hain.
  const old = metaRows_(tab).filter(m => m.getLocation().getRow().getRow() === top)[0];
  const oldId = old ? old.getValue() : null;
  if (old) old.remove();
  tab.insertRowAfter(top);
  tab.getRange(top, 1, 1, width).copyTo(tab.getRange(top + 1, 1, 1, width));
  tab.getRange(top, 1, 1, DATA_COLS).setValues([values]);
  if (oldId) tab.getRange(top + 1, 1, 1, tab.getMaxColumns()).addDeveloperMetadata(META_KEY, oldId);
  tab.getRange(top, 1, 1, tab.getMaxColumns()).addDeveloperMetadata(META_KEY, id);
}

function updateExisting_(tab, id, values) {
  const m = metaRows_(tab, id)[0];
  if (!m) return;
  const r = m.getLocation().getRow().getRow();
  const cur = tab.getRange(r, 1, 1, DATA_COLS).getValues()[0];
  if (String(values[LAST_UPDATED_COL - 1]).trim() === '') values[LAST_UPDATED_COL - 1] = cur[LAST_UPDATED_COL - 1];
  tab.getRange(r, 1, 1, DATA_COLS).setValues([values]);
}
