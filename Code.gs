// Tutkintoraportit SNOP: vastaanottaa raportit sovelluksesta, tallentaa Sheetsiin,
// tekee hyväksytystä tutkinnosta todistus-PDF:n (Slides-pohja) ja tulostettavan tutkintoraportti-PDF:n.
const BACKEND_VERSION = '2.6';
const SHEET_NAME = 'Tutkinnot';
const REPORT_FOLDER_NAME = 'Tutkintoraportit';        // tulostettavat tutkintoraportti-PDF:t
const JSON_FOLDER_NAME = 'Tutkintoraportit data';     // JSON-varmuuskopiot
const PDF_FOLDER_NAME = 'Todistukset SNOP';           // valmiit todistus-PDF:t
const TEMPLATE_ID = '1m2NMN1QQvPnHYrQTFiiz9lxyA0fSFErNcO5npsjICeY'; // Slides-pohja (vuokraveneen kuljettaja, M)

const HEADERS = ['Etunimi', 'Sukunimi', 'Katuosoite', 'Postinumero', 'Postitoimipaikka', 'Syntymäaika', 'Sähköposti',
  'Tutkintopaikka', 'Tarkastaja', 'ICC', 'Vuokravene', 'Todistus (PDF)', 'Tutkinto pvm', 'ID', 'Tutkintoraportti (PDF)'];

const SECTION_TITLES = {
  1: '1. Valmistelut ja turvallisuus',
  2: '2. Aluksen käsittely',
  3: '3. Moottoriajo ja merimiestaito (moottorivene)',
  4: '4. Purjehdus ja merimiestaito (purjevene)',
  5: '5. Navigointi ja tilanteen hallinta',
  6: '6. Henkilö yli laidan (POB)'
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rec = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    ensureHeader(sh);

    // Duplikaattisuojaus: id on sarakkeessa N (14)
    const last = sh.getLastRow();
    if (last > 0) {
      const ids = sh.getRange(1, 14, last, 1).getValues().flat();
      if (ids.indexOf(rec.id) !== -1) return out({ ok: true, duplicate: true });
    }

    // A-K tiedot, L = todistus (täytetään), M = pvm, N = id, O = tutkintoraportti (täytetään)
    sh.appendRow(rec.row.concat([rec.id, '']));
    const rowNum = sh.getLastRow();
    // appendRow muuntaa "00140" -> 140 ja "08.10.2026" -> päivämäärä; kirjoitetaan nämä uudelleen tekstinä
    [4, 6, 13].forEach(function (c) {
      sh.getRange(rowNum, c).setNumberFormat('@').setValue(String(rec.row[c - 1] == null ? '' : rec.row[c - 1]));
    });

    // JSON-varmuuskopio (raakadata)
    const jname = (rec.report.kokelasNimi || 'kokelas') + '_' + (rec.report.paiva || '') + '_' + rec.id + '.json';
    getFolder(JSON_FOLDER_NAME).createFile(jname, JSON.stringify(rec.report, null, 2), MimeType.PLAIN_TEXT);

    // Tulostettava tutkintoraportti-PDF (aina)
    let reportUrl = '', reportError = '';
    try {
      reportUrl = makeReportPdf(rec);
      setLinks(sh.getRange(rowNum, 15), [['Tutkintoraportti', reportUrl]]);
    } catch (err) {
      reportError = String(err);
      sh.getRange(rowNum, 15).setValue('Raporttivirhe: ' + reportError);
    }

    // Todistus-PDF (vain hyväksytty + vuokravene valittu)
    let pdfUrls = [], pdfError = '';
    try {
      if (rec.report.paatos === 'hyvaksytty' && rec.report.vuokravene) {
        const made = makeCertificates(rec);
        pdfUrls = made.map(function (m) { return m.url; });
        setLinks(sh.getRange(rowNum, 12), made.map(function (m) { return ['Todistus ' + m.laji, m.url]; }));
      }
    } catch (err) {
      pdfError = String(err);
      sh.getRange(rowNum, 12).setValue('PDF-virhe: ' + pdfError);
    }

    return out({ ok: true, pdfUrls: pdfUrls, pdfError: pdfError, reportUrl: reportUrl, reportError: reportError });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Lisää otsikkorivi tarvittaessa (myös jo olemassa olevaan taulukkoon).
function ensureHeader(sh) {
  // Tekstimuoto: postinumero (D), syntymäaika (F), tutkinto pvm (M), id (N) – etunollat ja päivämäärät säilyvät sellaisinaan
  [4, 6, 13, 14].forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  if (sh.getLastRow() > 0 && sh.getRange(1, 1).getValue() === HEADERS[0]) return;
  if (sh.getLastRow() > 0) sh.insertRowBefore(1);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#0a4272').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(12, 130);
  sh.setColumnWidth(15, 150);
}

// Aja kerran editorissa: muuntaa vanhat rivit (L = pelkkä URL -> linkki, O = JSON -> tutkintoraportti-PDF).
function migrateOldRows() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  ensureHeader(sh);
  for (let r = 2; r <= sh.getLastRow(); r++) {
    const l = String(sh.getRange(r, 12).getValue());
    if (l.indexOf('http') === 0) {
      const urls = l.split(/\s+/).filter(String);
      const laji = String(sh.getRange(r, 11).getValue());
      setLinks(sh.getRange(r, 12), urls.map(function (u, i) { return ['Todistus ' + (laji.length === urls.length ? laji.charAt(i) : (i + 1)), u]; }));
    }
    const o = String(sh.getRange(r, 15).getValue());
    if (o.charAt(0) === '{') {
      const report = JSON.parse(o);
      const row = sh.getRange(r, 1, 1, 13).getValues()[0].map(String);
      try {
        setLinks(sh.getRange(r, 15), [['Tutkintoraportti', makeReportPdf({ id: String(sh.getRange(r, 14).getValue()), row: row, report: report })]]);
      } catch (err) {
        sh.getRange(r, 15).setValue('Raporttivirhe: ' + err);
      }
    }
  }
}

// Kirjoittaa soluun klikattavat linkit (yksi per rivi): [[teksti, url], ...]
function setLinks(range, items) {
  const text = items.map(function (i) { return i[0]; }).join('\n');
  let b = SpreadsheetApp.newRichTextValue().setText(text);
  let pos = 0;
  items.forEach(function (i) {
    b = b.setLinkUrl(pos, pos + i[0].length, i[1]);
    pos += i[0].length + 1;
  });
  range.setRichTextValue(b.build());
}

function makeCertificates(rec) {
  const r = rec.report;
  const row = rec.row;
  const nimi = [row[0], row[1]].filter(String).join(' ');
  const syntyma = row[5] || '';
  const paikkaAika = [row[7], row[12]].filter(String).join(', ');
  const outFolder = getFolder(PDF_FOLDER_NAME);
  const made = [];

  r.vuokravene.split('').forEach(function (laji) {          // 'M' ja/tai 'S'
    const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy('tmp_' + rec.id + '_' + laji);
    try {
      const pres = SlidesApp.openById(copy.getId());
      pres.replaceAllText('Etunimi Sukunimi', nimi);
      pres.replaceAllText('dd.mm.yyyy', syntyma);
      setPlaceAndDate(pres, paikkaAika);
      if (laji === 'S') {
        pres.replaceAllText('(M)', '(S)');
        pres.replaceAllText('moottoriveneellä', 'purjeveneellä');
      }
      pres.saveAndClose();
      const pdfName = 'Vuokraveneen kuljettaja ' + laji + ' - ' + nimi + ' - ' + (row[12] || '') + '.pdf';
      const pdf = outFolder.createFile(copy.getAs('application/pdf').setName(pdfName));
      made.push({ laji: laji, url: pdf.getUrl() });
    } finally {
      copy.setTrashed(true);
    }
  });
  return made;
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

// ---------- Tutkintoraportti (selkokielinen, tulostettava PDF) ----------

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
}

function reportHtml(rec) {
  const r = rec.report;
  const kv = function (k, v) { return '<tr><td class="k">' + esc(k) + '</td><td>' + (esc(v) || '–') + '</td></tr>'; };
  const lajit = (r.vuokravene || '').split('').map(function (c) { return c === 'M' ? 'moottorialus' : 'purjealus'; }).join(' + ');
  const icc = (r.icc || '').split('').map(function (c) { return c === 'M' ? 'moottorialus' : 'purjealus'; }).join(' + ');

  let secs = '';
  [1, 2, 3, 4, 5, 6].forEach(function (n) {
    const items = (r.checks && r.checks[n]) || [];
    const anyChecked = items.some(function (i) { return i.rastitettu; });
    const motor = (r.vuokravene || '').indexOf('M') !== -1 || (r.icc || '').indexOf('M') !== -1;
    const sail = (r.vuokravene || '').indexOf('S') !== -1 || (r.icc || '').indexOf('S') !== -1;
    if ((n === 3 && !anyChecked && !motor) || (n === 4 && !anyChecked && !sail)) return;
    const done = items.filter(function (i) { return i.rastitettu; }).length;
    secs += '<h3>' + esc(SECTION_TITLES[n]) + ' <span class="cnt">(' + done + ' / ' + items.length + ')</span></h3><table class="chk">';
    items.forEach(function (i) {
      const idx = i.teksti.indexOf(':');
      const t = idx > 0 && idx < 60 ? '<b>' + esc(i.teksti.slice(0, idx + 1)) + '</b>' + esc(i.teksti.slice(idx + 1)) : esc(i.teksti);
      secs += '<tr><td class="box ' + (i.rastitettu ? 'yes' : 'no') + '">' + (i.rastitettu ? '[X]' : '[&nbsp;&nbsp;]') + '</td><td>' + t + '</td></tr>';
    });
    secs += '</table>';
  });

  const paatos = r.paatos === 'hyvaksytty' ? '<span class="ok">HYVÄKSYTTY</span>'
    : r.paatos === 'uusinta' ? '<span class="bad">UUSINTA VAADITAAN</span>' : '(ei merkitty)';

  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:10pt;color:#1a2a3a;} tr{page-break-inside:avoid;} .keep{page-break-inside:avoid;}' +
    'h1{font-size:19pt;color:#0a4272;margin:0 0 2px 0;} .sub{color:#6c7a8a;margin:0 0 12px 0;}' +
    'h2{font-size:13pt;color:#0a4272;border-bottom:2px solid #0a4272;padding:2px 0;margin:14px 0 6px 0;page-break-after:avoid;}' +
    'h3{font-size:11pt;color:#0a4272;margin:10px 0 4px 0;border-bottom:1px solid #c5d5e8;page-break-after:avoid;} .cnt{color:#6c7a8a;font-weight:normal;font-size:9.5pt;}' +
    'table{border-collapse:collapse;width:100%;} td{padding:2px 6px;vertical-align:top;}' +
    'td.k{width:32%;color:#4a5a6a;} .chk td{border-bottom:1px solid #eef2f7;} td.box{width:34px;font-family:Courier New,monospace;white-space:nowrap;}' +
    'td.yes{color:#1b7a36;font-weight:bold;} td.no{color:#9aa7b4;}' +
    '.ok{color:#1b7a36;font-weight:bold;font-size:14pt;} .bad{color:#b02a37;font-weight:bold;font-size:14pt;}' +
    '.box2{border:1px solid #c5d5e8;padding:6px 8px;min-height:30px;} .foot{margin-top:18px;color:#6c7a8a;font-size:9pt;}' +
    '</style></head><body>' +
    '<h1>Tutkintoraportti</h1><p class="sub">Suomen Navigoinninopettajat / Tutkinnot – veneilytaidon tutkinto</p>' +
    '<h2>Kokelas</h2><table>' + kv('Nimi', r.kokelasNimi) + kv('Syntymäaika', r.syntymaika) + kv('Osoite', r.osoite) +
      kv('Puhelin', r.puhelin) + kv('Sähköposti', r.kokelasEmail) + '</table>' +
    '<h2>Alus ja tutkinto</h2><table>' + kv('Alus', [r.alusNimi, r.alustyyppi].filter(String).join(', ')) +
      kv('Rekisterinumero', r.rekisteri) + kv('Moottoriteho', r.moottori) + kv('Tutkintopäivä', formatFi(r.paiva)) +
      kv('Vesialue / sijainti', r.vesialue) + kv('Tuuli', r.tuuli) + kv('Näkyvyys', r.nakyvyys) +
      kv('Vuokravene', lajit) + kv('ICC', icc) + kv('Tarkastaja', r.vastaanottajaNimi) + '</table>' +
    '<h2>Suoritetut tehtävät</h2>' + secs +
    '<div class="keep"><h2>Arviointi</h2>' +
    '<h3>Vahvuudet</h3><div class="box2">' + (esc(r.vahvuudet) || '–') + '</div>' +
    '<h3>Kehitettävää</h3><div class="box2">' + (esc(r.kehitettavaa) || '–') + '</div>' +
    '<h3>Päätös</h3><p>' + paatos + '</p></div>' +
    '<h2>Tarkastaja</h2><table>' + kv('Nimi', r.vastaanottajaNimi) + kv('Sähköposti', r.vastaanottajaEmail) + kv('Puhelin', r.vastaanottajaPuh) + '</table>' +
    '<p class="foot">Raportti luotu ' + esc(Utilities.formatDate(new Date(), 'Europe/Helsinki', 'd.M.yyyy HH:mm')) + ' SNOP Tutkintoraportti -sovelluksella.</p>' +
    '</body></html>';
}

function formatFi(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
  return m ? (parseInt(m[3], 10) + '.' + parseInt(m[2], 10) + '.' + m[1]) : (iso || '');
}

function makeReportPdf(rec) {
  const r = rec.report;
  const name = 'Tutkintoraportti - ' + (r.kokelasNimi || 'kokelas') + ' - ' + (formatFi(r.paiva) || '') + '.pdf';
  const blob = Utilities.newBlob(reportHtml(rec), 'text/html', 'raportti.html').getAs('application/pdf').setName(name);
  return getFolder(REPORT_FOLDER_NAME).createFile(blob).getUrl();
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
    pdfFolderUrl: getFolder(PDF_FOLDER_NAME).getUrl(),
    reportFolderUrl: getFolder(REPORT_FOLDER_NAME).getUrl()
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
