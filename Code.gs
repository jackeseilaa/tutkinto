// Google Apps Script: vastaanottaa tutkintoraportit sovelluksesta ja tallentaa ne Driveen.
// Sidottu Google Sheets -tiedostoon (Laajennukset > Apps Script).
// Julkaisu: Ota käyttöön > Uusi käyttöönotto > Verkkosovellus
//   Suorita nimellä: Minä | Pääsy: Kaikki
// Kopioi /exec-osoite sovelluksen index.html-tiedoston SYNC_URL-muuttujaan.

const SHEET_NAME = 'Tutkinnot';
const FOLDER_NAME = 'Tutkintoraportit';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rec = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Duplikaattisuojaus: id on sarakkeessa N (14)
    const last = sh.getLastRow();
    if (last > 0) {
      const ids = sh.getRange(1, 14, last, 1).getValues().flat();
      if (ids.indexOf(rec.id) !== -1) return out({ ok: true, duplicate: true });
    }

    // Sarakkeet A-M kuten aiemmin, N = id, O = koko raportti JSONina
    sh.appendRow(rec.row.concat([rec.id, JSON.stringify(rec.report)]));

    // JSON-varmuuskopio Driveen
    const it = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
    const name = (rec.report.kokelasNimi || 'kokelas') + '_' + (rec.report.paiva || '') + '_' + rec.id + '.json';
    folder.createFile(name, JSON.stringify(rec.report, null, 2), MimeType.PLAIN_TEXT);

    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out({ ok: true, service: 'tutkinto' }); }

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
