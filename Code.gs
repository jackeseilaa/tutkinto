// Tutkintoraportit SNOP: vastaanottaa raportit sovelluksesta, tallentaa Sheetsiin
// ja Driveen, ja tekee hyväksytystä tutkinnosta todistus-PDF:n Slides-pohjasta.
const BACKEND_VERSION = '2.3';
const SHEET_NAME = 'Tutkinnot';
const FOLDER_NAME = 'Tutkintoraportit';        // JSON-varmuuskopiot
const PDF_FOLDER_NAME = 'Todistukset SNOP';     // valmiit todistus-PDF:t
const TEMPLATE_ID = '1m2NMN1QQvPnHYrQTFiiz9lxyA0fSFErNcO5npsjICeY'; // Slides-pohja (vuokraveneen kuljettaja, M)

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

    // Sarakkeet A-K, L = todistus-PDF (täytetään alla), M = pvm, N = id, O = raportti JSON
    sh.appendRow(rec.row.concat([rec.id, JSON.stringify(rec.report)]));
    const rowNum = sh.getLastRow();

    // JSON-varmuuskopio
    const folder = getFolder(FOLDER_NAME);
    const name = (rec.report.kokelasNimi || 'kokelas') + '_' + (rec.report.paiva || '') + '_' + rec.id + '.json';
    folder.createFile(name, JSON.stringify(rec.report, null, 2), MimeType.PLAIN_TEXT);

    // Todistus-PDF (vain hyväksytty + vuokravene valittu). Virhe ei estä tallennusta.
    let pdfUrls = [], pdfError = '';
    try {
      if (rec.report.paatos === 'hyvaksytty' && rec.report.vuokravene) {
        pdfUrls = makeCertificates(rec);
        if (pdfUrls.length) sh.getRange(rowNum, 12).setValue(pdfUrls.join('\n'));
      }
    } catch (err) {
      pdfError = String(err);
      sh.getRange(rowNum, 12).setValue('PDF-virhe: ' + pdfError);
    }

    return out({ ok: true, pdfUrls: pdfUrls, pdfError: pdfError });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function makeCertificates(rec) {
  const r = rec.report;
  const row = rec.row;
  const nimi = [row[0], row[1]].filter(String).join(' ');
  const syntyma = row[5] || '';
  const paikkaAika = [row[7], row[12]].filter(String).join(', ');
  const outFolder = getFolder(PDF_FOLDER_NAME);
  const urls = [];

  r.vuokravene.split('').forEach(function (laji) {          // 'M' ja/tai 'S'
    const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy('tmp_' + rec.id + '_' + laji);
    try {
      const pres = SlidesApp.openById(copy.getId());
      pres.replaceAllText('Etunimi Sukunimi', nimi);
      pres.replaceAllText('dd.mm.yyyy', syntyma);
      setPlaceAndDate(pres, paikkaAika);
      if (laji === 'S') {
        pres.replaceAllText('(M)', '(S)');
        pres.replaceAllText('moottoriveneell\u00e4', 'purjeveneell\u00e4');
      }
      pres.saveAndClose();
      const pdfName = 'Vuokraveneen kuljettaja ' + laji + ' - ' + nimi + ' - ' + (row[12] || '') + '.pdf';
      const pdf = outFolder.createFile(copy.getAs('application/pdf').setName(pdfName));
      urls.push(pdf.getUrl());
    } finally {
      copy.setTrashed(true);
    }
  });
  return urls;
}

// Korvaa 'Paikka ja aika' -tekstin, levittää kentän ja keskittää, ettei rivi mene allekirjoituksen päälle.
function setPlaceAndDate(pres, text) {
  pres.getSlides()[0].getShapes().forEach(function (sh) {
    let s = '';
    try { s = sh.getText().asString(); } catch (e) { return; }
    if (s.indexOf('Paikka ja aika') === -1) return;
    const centerX = sh.getLeft() + sh.getWidth() / 2;
    const w = 330;
    sh.getText().setText(text);
    sh.setWidth(w);
    sh.setLeft(centerX - w / 2);
    sh.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  });
}

function getFolder(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function doGet() {
  return out({
    ok: true,
    service: 'tutkinto',
    version: BACKEND_VERSION,
    pdfFolderUrl: getFolder(PDF_FOLDER_NAME).getUrl()
  });
}

// Aja kerran editorissa: antaa skriptille tarvittavat luvat (Sheets, Drive, Slides).
function authorize() {
  SpreadsheetApp.getActiveSpreadsheet();
  getFolder(PDF_FOLDER_NAME);
  SlidesApp.openById(TEMPLATE_ID).getSlides().length;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
