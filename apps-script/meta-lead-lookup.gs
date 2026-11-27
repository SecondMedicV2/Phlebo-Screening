// Meta lead lookup for thyrocarejobs.in
// Deployed as a Google Apps Script web app (Execute as: me, Access: Anyone).
// GET ?phone=9876543210 -> latest matching row from "Experiment Phlebo hiring" › DUMP.
//
// DUMP columns used: A id (l:…), N DMLT/experience?, O bike/scooty?, Q full name,
// S city, T post code, U email, X Phone Number (Text, 10 digits).

const SHEET_ID = '1G4B2RId0nqJu2qnfmzplfyti8pCZYqLESKbp79b6-Fc';
const SHEET_GID = 1047572462; // DUMP tab
const PHONE_COLUMN = 24; // X
const MAX_LOOKUPS_PER_MINUTE = 120;
const CACHE_SECONDS = 600;

function doGet(e) {
  const phone = String((e && e.parameter && e.parameter.phone) || '').replace(/\D/g, '').slice(-10);
  if (!/^[6-9]\d{9}$/.test(phone)) return json_({ found: false });
  if (!allowLookup_()) return json_({ found: false, error: 'busy' });

  const cache = CacheService.getScriptCache();
  const cached = cache.get('p' + phone);
  if (cached) return text_(cached);

  const out = JSON.stringify(lookup_(phone));
  cache.put('p' + phone, out, CACHE_SECONDS);
  return text_(out);
}

function lookup_(phone) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheets().find((s) => s.getSheetId() === SHEET_GID) || ss.getSheetByName('DUMP');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false };

  const matches = sheet.getRange(2, PHONE_COLUMN, lastRow - 1, 1)
    .createTextFinder(phone).matchEntireCell(true).findAll();
  if (!matches.length) return { found: false };

  const row = matches[matches.length - 1].getRow(); // newest lead for this number
  const v = sheet.getRange(row, 1, 1, 21).getDisplayValues()[0];
  const yes = (s) => String(s).trim().toLowerCase() === 'yes';
  return {
    found: true,
    lead: String(v[0]).replace(/^l:/, '').trim(),
    name: String(v[16]).trim(),
    city: String(v[18]).trim(),
    pin: String(v[19]).replace(/\D/g, '').slice(0, 6),
    email: String(v[20]).trim(),
    bike: yes(v[14]),
    dmlt: yes(v[13]),
  };
}

// Global cap on lookups per minute so numbers can't be scanned quickly.
function allowLookup_() {
  const cache = CacheService.getScriptCache();
  const key = 'rl' + Math.floor(Date.now() / 60000);
  const count = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(count), 120);
  return count <= MAX_LOOKUPS_PER_MINUTE;
}

function json_(obj) { return text_(JSON.stringify(obj)); }
function text_(s) { return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON); }
