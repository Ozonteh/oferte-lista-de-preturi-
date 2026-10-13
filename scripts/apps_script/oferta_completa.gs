/**
 * Ozonteh: pagina de start a ofertei + ofertele produselor alese, într-un singur PDF.
 *
 * Instalare (o singură dată pentru fiecare foaie de calcul):
 *   1. Fișier → Salvează ca Foi de calcul Google (dacă titlul are .XLSX).
 *   2. Extensii → Apps Script → lipiți tot fișierul acesta → Salvați.
 *   3. Reîncărcați foaia de calcul: apare meniul „Ozonteh”.
 *
 * Meniul are două comenzi:
 *   - „Descarcă oferta completă (PDF)”: salvează PDF-ul în Drive și dă linkul;
 *   - „Tipărește oferta completă”: deschide același PDF, gata de tipar (Ctrl+P).
 *
 * Cum lucrează: copiază temporar foaia fiecărei oferte de produs în fișierul
 * curent, ascunde celelalte foi, exportă PDF-ul (pagina de start, apoi ofertele
 * produselor în ordinea din tabel), apoi șterge copiile și reface totul.
 *
 * Citește câmpurile prin nume (le pune scripts/genereaza_lista.py):
 *   OfertaProduse, OfertaLinkuri, OfertaClient, OfertaNr, OfertaData.
 */

var FOAIE_OFERTA = 'Ofertă comercială';
var DOSAR_PDF = 'Oferte Ozonteh (PDF)';
var PREFIX_TEMPORAR = '_fisa_';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ozonteh')
    .addItem('Descarcă oferta completă (PDF)', 'descarcaOfertaCompleta')
    .addItem('Tipărește oferta completă', 'tiparesteOfertaCompleta')
    .addToUi();
}

function descarcaOfertaCompleta() {
  var rezultat = faPdf();
  if (rezultat) arataRezultat(rezultat, false);
}

function tiparesteOfertaCompleta() {
  var rezultat = faPdf();
  if (rezultat) arataRezultat(rezultat, true);
}

/** Face PDF-ul complet și îl salvează în Drive. Întoarce null dacă n-a mers. */
function faPdf() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var oferta = ss.getSheetByName(FOAIE_OFERTA);
  var produseRange = ss.getRangeByName('OfertaProduse');
  var linkuriRange = ss.getRangeByName('OfertaLinkuri');
  if (!oferta || !produseRange || !linkuriRange) {
    ui.alert('Nu găsesc foaia „' + FOAIE_OFERTA + '” sau tabelul de produse. ' +
             'Importați din nou fișierul Excel generat.');
    return null;
  }
  stergeTemporare(ss);
  SpreadsheetApp.flush();

  var numeProduse = produseRange.getValues();
  var linkuri = linkuriRange.getValues();
  var produse = [];
  var faraOferta = [];
  for (var i = 0; i < numeProduse.length; i++) {
    var nume = String(numeProduse[i][0]).trim();
    if (!nume) continue;
    var link = String(linkuri[i][0]).trim();
    var id = (link.match(/\/d\/([\w-]+)/) || [])[1];
    if (!id) {
      faraOferta.push(nume);
      continue;
    }
    var gid = (link.match(/[#?&]gid=(\d+)/) || [])[1] || '';
    var dejaAdaugat = produse.some(function (p) { return p.id === id && p.gid === gid; });
    if (!dejaAdaugat) produse.push({ nume: nume, id: id, gid: gid });
  }

  var temporare = [];
  var ascunse = [];
  var nepreluate = [];
  try {
    // 1. Ofertele produselor, copiate imediat după pagina de start.
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

    // 2. Rămân vizibile doar pagina de start și ofertele produselor.
    var pastrate = [oferta.getSheetId()].concat(temporare.map(function (t) { return t.getSheetId(); }));
    ss.setActiveSheet(oferta);
    ss.getSheets().forEach(function (s) {
      if (pastrate.indexOf(s.getSheetId()) === -1 && !s.isSheetHidden()) {
        s.hideSheet();
        ascunse.push(s);
      }
    });
    SpreadsheetApp.flush();

    // 3. Exportul PDF al foilor vizibile; scale=4 = fiecare foaie pe o singură pagină A4.
    var parametri = [
      'format=pdf', 'size=A4', 'portrait=true', 'scale=4', 'gridlines=false',
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
    var fisier = dosar().createFile(raspuns.getBlob().setName(numeFisier(ss)));
    return { fisier: fisier, nrFise: temporare.length, faraOferta: faraOferta, nepreluate: nepreluate };
  } catch (e) {
    ui.alert('Oferta completă nu s-a putut face: ' + e.message);
    return null;
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

function valoare(ss, numeCamp) {
  var r = ss.getRangeByName(numeCamp);
  return r ? r.getValue() : '';
}

function numeFisier(ss) {
  var client = String(valoare(ss, 'OfertaClient')).trim() || 'client';
  var data = valoare(ss, 'OfertaData');
  var zi = Utilities.formatDate(data instanceof Date ? data : new Date(),
                                Session.getScriptTimeZone(), 'dd.MM.yyyy');
  var nr = String(valoare(ss, 'OfertaNr')).trim();
  var nume = 'Oferta ' + (nr ? nr + ' ' : '') + client + ' ' + zi;
  return nume.replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() + '.pdf';
}

function arataRezultat(rezultat, pentruTipar) {
  var fisier = rezultat.fisier;
  var descarcare = 'https://drive.google.com/uc?export=download&id=' + fisier.getId();
  var vizualizare = 'https://drive.google.com/file/d/' + fisier.getId() + '/view';
  var html = '<div style="font-family:Arial,sans-serif;font-size:13px;line-height:1.5">' +
    '<p><b>' + esc(fisier.getName()) + '</b><br>Pagina de start + ' + rezultat.nrFise +
    (rezultat.nrFise === 1 ? ' ofertă de produs.' : ' oferte de produs.') + '</p>';
  if (rezultat.faraOferta.length) {
    html += '<p>Fără ofertă de produs în listă: ' + esc(rezultat.faraOferta.join(', ')) + '</p>';
  }
  if (rezultat.nepreluate.length) {
    html += '<p style="color:#C00000">Nu am putut deschide: ' + esc(rezultat.nepreluate.join('; ')) + '</p>';
  }
  if (pentruTipar) {
    html += '<p><a href="' + vizualizare + '" target="_blank">Deschide PDF-ul pentru tipar</a>' +
      ' (în fereastra nouă apăsați Ctrl+P / ⌘P)</p>' +
      '<script>window.open(' + JSON.stringify(vizualizare) + ', "_blank");</script>';
  } else {
    html += '<p><a href="' + descarcare + '" target="_blank">Descarcă PDF-ul</a> &nbsp;·&nbsp; ' +
      '<a href="' + vizualizare + '" target="_blank">Deschide în Drive</a></p>' +
      '<script>window.open(' + JSON.stringify(descarcare) + ', "_blank");</script>';
  }
  html += '</div>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(480).setHeight(210),
    'Oferta completă e gata'
  );
}

function esc(text) {
  return String(text).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
