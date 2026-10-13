/**
 * Ozonteh: oferta comercială + ofertele produselor alese, într-un singur PDF.
 *
 * Instalare (o singură dată pentru fiecare foaie de calcul):
 *   1. Fișier → Salvează ca Foi de calcul Google (dacă titlul are .XLSX).
 *   2. Extensii → Apps Script → lipiți tot fișierul acesta → Salvați.
 *   3. Reîncărcați foaia de calcul: apare meniul „Ozonteh”.
 *
 * Cum lucrează: copiază temporar foaia fiecărei oferte de produs în fișierul
 * curent, ascunde celelalte foi, exportă PDF-ul (oferta comercială primă,
 * apoi fișele în ordinea din tabel), apoi șterge copiile și reface totul.
 *
 * Depinde de pozițiile din scripts/genereaza_lista.py: produsele în B19:B28,
 * linkul la oferta produsului în coloana ascunsă I.
 */

var FOAIE_OFERTA = 'Ofertă comercială';
var PRIMA_LINIE = 19;
var NR_LINII = 10;
var COL_PRODUS = 2;        // B
var COL_LINK_OFERTA = 9;   // I (ascunsă)
var DOSAR_PDF = 'Oferte Ozonteh (PDF)';
var PREFIX_TEMPORAR = '_fisa_';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ozonteh')
    .addItem('Descarcă oferta completă (PDF)', 'descarcaOfertaCompleta')
    .addToUi();
}

function descarcaOfertaCompleta() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var oferta = ss.getSheetByName(FOAIE_OFERTA);
  if (!oferta) {
    ui.alert('Nu găsesc foaia „' + FOAIE_OFERTA + '”.');
    return;
  }
  stergeTemporare(ss);
  SpreadsheetApp.flush();

  var produse = [];
  var faraOferta = [];
  oferta.getRange(PRIMA_LINIE, 1, NR_LINII, COL_LINK_OFERTA).getValues().forEach(function (r) {
    var nume = String(r[COL_PRODUS - 1]).trim();
    if (!nume) return;
    var link = String(r[COL_LINK_OFERTA - 1]).trim();
    var id = (link.match(/\/d\/([\w-]+)/) || [])[1];
    if (!id) {
      faraOferta.push(nume);
      return;
    }
    var gid = (link.match(/[#?&]gid=(\d+)/) || [])[1] || '';
    var dejaAdaugat = produse.some(function (p) { return p.id === id && p.gid === gid; });
    if (!dejaAdaugat) produse.push({ nume: nume, id: id, gid: gid });
  });

  var temporare = [];
  var ascunse = [];
  var nepreluate = [];
  try {
    // 1. Ofertele produselor, copiate imediat după oferta comercială.
    produse.forEach(function (p) {
      var foaieSursa;
      try {
        foaieSursa = foaieOferta(SpreadsheetApp.openById(p.id), p.gid);
      } catch (e) {
        nepreluate.push(p.nume + ' (' + e.message + ')');
        return;
      }
      var copie = foaieSursa.copyTo(ss);
      copie.setName(PREFIX_TEMPORAR + (temporare.length + 1));
      ingheataFormule(foaieSursa, copie);
      ss.setActiveSheet(copie);
      ss.moveActiveSheet(oferta.getIndex() + temporare.length + 1);
      temporare.push(copie);
    });

    // 2. Rămân vizibile doar oferta și fișele.
    var pastrate = [oferta.getSheetId()].concat(temporare.map(function (t) { return t.getSheetId(); }));
    ss.setActiveSheet(oferta);
    ss.getSheets().forEach(function (s) {
      if (pastrate.indexOf(s.getSheetId()) === -1 && !s.isSheetHidden()) {
        s.hideSheet();
        ascunse.push(s);
      }
    });
    SpreadsheetApp.flush();

    // 3. Exportul PDF al foilor vizibile.
    var parametri = [
      'format=pdf', 'size=A4', 'portrait=true', 'fitw=true', 'gridlines=false',
      'printtitle=false', 'sheetnames=false', 'pagenum=UNDEFINED', 'fzr=false',
      'horizontal_alignment=CENTER', 'top_margin=0.4', 'bottom_margin=0.4',
      'left_margin=0.4', 'right_margin=0.4'
    ];
    var raspuns = UrlFetchApp.fetch(
      'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + parametri.join('&'),
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
    );
    if (raspuns.getResponseCode() !== 200) {
      throw new Error('exportul PDF a răspuns cu codul ' + raspuns.getResponseCode());
    }
    var fisier = dosar().createFile(raspuns.getBlob().setName(numeFisier(oferta)));
    arataRezultat(fisier, temporare.length, faraOferta, nepreluate);
  } catch (e) {
    ui.alert('Oferta completă nu s-a putut face: ' + e.message);
  } finally {
    ascunse.forEach(function (s) { s.showSheet(); });
    ss.setActiveSheet(oferta);
    temporare.forEach(function (t) { ss.deleteSheet(t); });
  }
}

/** Foaia cu `gid` din link, altfel prima foaie a ofertei. */
function foaieOferta(fisier, gid) {
  var foi = fisier.getSheets();
  for (var i = 0; i < foi.length && gid; i++) {
    if (String(foi[i].getSheetId()) === gid) return foi[i];
  }
  return foi[0];
}

/**
 * GOOGLEFINANCE și IMPORT* se încarcă cu întârziere în copie. Copia primește
 * valorile deja calculate în oferta originală, ca PDF-ul să nu arate „Loading…”.
 */
function ingheataFormule(sursa, copie) {
  var zona = sursa.getDataRange();
  var formule = zona.getFormulas();
  var valori = zona.getValues();
  for (var r = 0; r < formule.length; r++) {
    for (var c = 0; c < formule[r].length; c++) {
      if (/GOOGLEFINANCE|IMPORT/i.test(formule[r][c])) {
        copie.getRange(r + 1, c + 1).setValue(valori[r][c]);
      }
    }
  }
}

/** Copii rămase de la o rulare întreruptă. */
function stergeTemporare(ss) {
  ss.getSheets().forEach(function (s) {
    if (s.getName().indexOf(PREFIX_TEMPORAR) === 0) ss.deleteSheet(s);
  });
}

function dosar() {
  var gasite = DriveApp.getFoldersByName(DOSAR_PDF);
  return gasite.hasNext() ? gasite.next() : DriveApp.createFolder(DOSAR_PDF);
}

function numeFisier(oferta) {
  var client = String(oferta.getRange('C12').getValue()).trim() || 'client';
  var data = oferta.getRange('F8').getValue();
  var zi = Utilities.formatDate(data instanceof Date ? data : new Date(),
                                Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var nr = String(oferta.getRange('C8').getValue()).trim();
  var nume = 'Oferta ' + (nr ? nr + ' ' : '') + client + ' ' + zi;
  return nume.replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() + '.pdf';
}

function arataRezultat(fisier, nrFise, faraOferta, nepreluate) {
  var descarcare = 'https://drive.google.com/uc?export=download&id=' + fisier.getId();
  var html = '<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5">' +
    '<p><b>' + esc(fisier.getName()) + '</b><br>Oferta comercială + ' + nrFise +
    (nrFise === 1 ? ' fișă de produs.' : ' fișe de produs.') + '</p>';
  if (faraOferta.length) {
    html += '<p>Fără ofertă de produs în listă: ' + esc(faraOferta.join(', ')) + '</p>';
  }
  if (nepreluate.length) {
    html += '<p style="color:#C00000">Nu am putut deschide: ' + esc(nepreluate.join('; ')) + '</p>';
  }
  html += '<p><a href="' + descarcare + '" target="_blank">Descarcă PDF-ul</a> &nbsp;·&nbsp; ' +
    '<a href="' + fisier.getUrl() + '" target="_blank">Deschide în Drive</a></p></div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(460).setHeight(200),
    'Oferta completă e gata'
  );
}

function esc(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
