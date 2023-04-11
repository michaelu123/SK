// Dieses Skript setzt folgende Dateien und Ordner voraus:
//    1. Ein Template im Format eine Google Präsentation
//       Dieses Template hat den Namen "Saisonkarte Template" und einen definierten File-ID
var TEMPLATE_ID = "1TbhwX5hrSuBPt5mkUlN3HU0u74UzwatXsgvihF4FS_8";
//
//    2. Einen Ordner, in dem es temporäre Dateien ablegen und am Skriptende wieder löschen kann
//       Name: Temporäre Daten
var TEMPORARY_FOLDER_ID = "1uQ2YRHfA7NwhzggrBgpgH2ZK4F1pdZH8";
//
//    3. Einen Ordner, in dem es die erzeugte Saisonkarte ablegen und archivieren kann
//       Name: Erzeugte Saisonkarten
// var RESULT_FOLDER_ID = '1zKj4KTj-IARemTKWitHU81IpMDF5GPgM'
//       Jetzt in der geteilten Ablage "AG Tagestouren Saisonkarten zum Abruf / Erzeugte Saisonkarten"
var RESULT_FOLDER_ID = "1QnL8mwAi2S9CKnFIqsphK6etEoarW92C";
//
//    4. Eine Google-Tabelle, in der die Konfigurationsdaten für das aktuelle Jahr gespeichert werden
//       Diese Tabelle muss am Anfang jedes Jahres manuell aktualisiert werden
//       Folgende Daten sind konfiguriert:
//       - Das aktuelle Jahr (z.B. "2020")
//       - Die laufende Nummer (z.B. 1)
//       - Das Start-Datum des Gültigkeitsbereichs (z.B. "1. März 2020")
//       - Das Ende-Datum des Gültigkeitsbereichs (z.B. "28. Februar 2021")
//       Name: Saisonkarte-Basisdaten
var BASIS_DATA_ID = "1nkOwWze7eRosS_w-E33XYk-Qah2mPjpYfC1EbEUgeI0";

let inited = false;
let headers = {};
let ssheet;
let bestSheet;
let büroSheet;
let evSheet;

// Indices are 1-based!!

// Email-Verifikation
let geschicktIndex = 2; // Haben Sie gerade eine Saisonkarte bestellt?
let mitemailIndex = 3; // Mit dieser Email-Adresse (bitte nicht ändern!) :

//Bestellungen
let mailBIndex; // E-Mail-Adresse, 2
let mitgliedsNameBIndex; // ADFC-Mitgliedsname, 3
let mitgliedsNummerBIndex; // ADFC-Mitgliedsnummer, 4
let nochEinmalBIndex; // Saisonkarte noch einmal schicken?. 5
let kontoNameBIndex; // Name des Kontoinhabers (kann leer bleiben falls gleich Mitgliedsname), 6
let ibanBIndex; // IBAN-Kontonummer, 7
let zustimmungBIndex; // Zustimmung zur SEPA-Lastschrift, 8
let verifiedBIndex; // Verifikation, 9
let gesendetBIndex; // Gesendet, 10

// Büro
let mailRIndex; // E-Mail-Adresse (optional), 2
let mitgliedsNameRIndex; // Name des Mitglieds, 3
let mitgliedsNummerRIndex; // ADFC-Mitgliedsnummer, 4
let saisonKarteRIndex; // Saisonkarte, 5

function isEmpty(str) {
  return !str || 0 === str.length; // I think !str is sufficient...
}

function init() {
  let ssheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ssheet.getSheets();
  for (let sheet of sheets) {
    let sheetName = sheet.getName();
    let sheetHeaders = {};
    Logger.log("sheetName %s", sheetName);
    headers[sheetName] = sheetHeaders;
    let numCols = sheet.getLastColumn();
    Logger.log("numCols %s", numCols);
    let row1Vals = sheet.getRange(1, 1, 1, numCols).getValues();
    Logger.log("sheetName %s row1 %s", sheetName, row1Vals);
    for (let i = 0; i < numCols; i++) {
      let v = row1Vals[0][i];
      if (isEmpty(v)) continue;
      sheetHeaders[v] = i + 1;
    }
    Logger.log("sheet %s %s", sheetName, sheetHeaders);

    if (sheet.getName() == "Bestellungen") {
      bestSheet = sheet;
      mailBIndex = sheetHeaders["E-Mail-Adresse"];
      mitgliedsNameBIndex = sheetHeaders["ADFC-Mitgliedsname"];
      mitgliedsNummerBIndex = sheetHeaders["ADFC-Mitgliedsnummer"];
      nochEinmalBIndex = sheetHeaders["Saisonkarte noch einmal schicken?"];
      kontoNameBIndex =
        sheetHeaders[
          "Name des Kontoinhabers (kann leer bleiben falls gleich Mitgliedsname)"
        ];
      ibanBIndex = sheetHeaders["IBAN-Kontonummer"];
      zustimmungBIndex = sheetHeaders["Zustimmung zur SEPA-Lastschrift"];
      verifiedBIndex = sheetHeaders["Verifikation"];
      gesendetBIndex = sheetHeaders["Gesendet"];
    } else if (sheet.getName() == "Büro") {
      büroSheet = sheet;
      mailRIndex = sheetHeaders["E-Mail-Adresse (optional)"];
      mitgliedsNameRIndex = sheetHeaders["Name des Mitglieds"];
      mitgliedsNummerRIndex = sheetHeaders["ADFC-Mitgliedsnummer"];
      saisonKarteRIndex = sheetHeaders["Saisonkarte"];
    } else if (sheet.getName() == "Email-Verifikation") {
      evSheet = sheet;
    }
  }
  inited = true;
}

function sendeSK() {
  let docLock = LockService.getScriptLock();
  let locked = docLock.tryLock(30000);
  if (!locked) {
    SpreadsheetApp.getUi().alert("Konnte Dokument nicht locken");
    return;
  }
  init();
  var evalues = null;
  if (evSheet.getLastRow() < 2) evalues = [];
  else
    evalues = evSheet.getSheetValues(
      2,
      1,
      evSheet.getLastRow() - 1,
      evSheet.getLastColumn(),
    );
  Logger.log("evalues %s", evalues);

  var bvalues = null;
  if (bestSheet.getLastRow() < 2) bvalues = [];
  else
    bvalues = bestSheet.getSheetValues(
      2,
      1,
      bestSheet.getLastRow() - 1,
      bestSheet.getLastColumn(),
    );
  Logger.log("bvalues %s", bvalues);

  for (var ex in evalues) {
    var erow = evalues[ex];
    if (erow[geschicktIndex - 1] != "Ja" || erow[mitemailIndex - 1] == "")
      continue;
    var eaddr = erow[mitemailIndex - 1];
    var img = null;
    for (var bx in bvalues) {
      bx = +bx; // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
      var brow = bvalues[bx];

      // Noch einmal senden?
      if (
        brow[mailBIndex - 1] != "" &&
        brow[nochEinmalBIndex - 1].indexOf("schon eine") > 0 &&
        brow[verifiedBIndex - 1] == "" &&
        brow[gesendetBIndex - 1] == ""
      ) {
        var baddr = brow[mailBIndex - 1];
        if (eaddr != baddr) continue;
        Logger.log("once more %s", baddr);
        img = getImageBest(
          brow[mitgliedsNameBIndex - 1],
          brow[mitgliedsNummerBIndex - 1],
          bvalues,
        );
        if (img == null)
          sendNotFoundEmail(
            baddr,
            brow[mitgliedsNameBIndex - 1],
            brow[mitgliedsNummerBIndex - 1],
          );
        else sendImgEmail(baddr, img);
      } else if (
        brow[mailBIndex - 1] != "" &&
        brow[zustimmungBIndex - 1] != "" &&
        brow[gesendetBIndex - 1] == ""
      ) {
        // nur Zeilen mit Zustimmung und nicht Gesendet
        // brow[zustimmungBIndex - 1] = Zustimmung..., brow[verifiedBIndex - 1] = Verifikation, brow[gesendetBIndex - 1] = Gesendet
        var baddr = brow[mailBIndex - 1];
        if (eaddr != baddr) continue;
        if (brow[verifiedBIndex - 1] == "") {
          // Bestellungen[Verifikation] = Email-Verif[Zeitstempel]
          bestSheet.getRange(bx + 2, 8 + 1).setValue(erow[0]);
          brow[verifiedBIndex - 1] = erow[0];
        }
        if (!isValid(brow[ibanBIndex - 1])) {
          Logger.log("wrong iban %s", baddr);
          sendWrongIbanEmail(baddr, brow[ibanBIndex - 1]);
          bestSheet.getRange(bx + 2, 9 + 1).setValue("Falsche IBAN");
          brow[gesendetBIndex - 1] = "Falsche IBAN";
          continue;
        }
        img = erzeugeKarte(
          brow[mitgliedsNameBIndex - 1],
          brow[mitgliedsNummerBIndex - 1],
          false,
        );
        Logger.log("send card %s", baddr);
        sendImgEmail(baddr, img);
      } else {
        continue;
      }
      // Bestellungen[Gesendet] = aktuelles Datum/Zeit
      if (img == null) {
        bestSheet.getRange(bx + 2, 9 + 1).setValue("nichtgefunden");
        brow[gesendetBIndex - 1] = "nichtgefunden";
      } else {
        bestSheet.getRange(bx + 2, 9 + 1).setValue(img.getName());
        brow[gesendetBIndex - 1] = img.getName();
      }
    }
  }

  var rvalues = null;
  if (büroSheet.getLastRow() < 2) rvalues = [];
  else
    rvalues = büroSheet.getSheetValues(
      2,
      1,
      büroSheet.getLastRow() - 1,
      büroSheet.getLastColumn(),
    );
  Logger.log("büro values %s", rvalues);
  img = null;
  for (var rx in rvalues) {
    rx = +rx; // confusingly, bx is initially a string, and is interpreted as A1Notation in sheet.getRange(bx) !
    var rrow = rvalues[rx];
    var name = rrow[mitgliedsNameRIndex - 1];
    var num = rrow[mitgliedsNummerRIndex - 1];
    if (name == "" || num == "" || rrow[saisonKarteRIndex - 1] != "") continue;
    img = getImageBüro(name, num, rvalues);
    if (img == null) img = getImageBest(name, num, bvalues); // perhaps overly ambitious
    if (img == null) img = erzeugeKarte(name, num, true);
    var baddr = rrow[mailRIndex - 1];
    if (baddr != null && baddr != "") sendImgEmail(baddr, img);
    büroSheet.getRange(rx + 2, saisonKarteRIndex).setValue(img.getName());
    rrow[saisonKarteRIndex - 1] = img.getName();
  }
  docLock.releaseLock();
}

function getImageBest(name, num, bvalues) {
  for (bx2 in bvalues) {
    brow = bvalues[bx2];
    if (
      name == brow[mitgliedsNameBIndex - 1] &&
      num == brow[mitgliedsNummerBIndex - 1] &&
      brow[gesendetBIndex - 1] != ""
    ) {
      fname = brow[gesendetBIndex - 1];
      var resultFolder = DriveApp.getFolderById(RESULT_FOLDER_ID);
      var files = resultFolder.getFilesByName(fname);
      if (files.hasNext()) return files.next();
    }
  }
  return null;
}

function getImageBüro(name, num, rvalues) {
  for (bx2 in rvalues) {
    rrow = rvalues[bx2];
    if (
      name == rrow[mitgliedsNameRIndex - 1] &&
      num == rrow[mitgliedsNummerRIndex - 1] &&
      rrow[saisonKarteRIndex - 1] != ""
    ) {
      fname = rrow[saisonKarteRIndex - 1];
      var resultFolder = DriveApp.getFolderById(RESULT_FOLDER_ID);
      var files = resultFolder.getFilesByName(fname);
      if (files.hasNext()) return files.next();
    }
  }
  return null;
}

var ibanLen = {
  NO: 15,
  BE: 16,
  DK: 18,
  FI: 18,
  FO: 18,
  GL: 18,
  NL: 18,
  MK: 19,
  SI: 19,
  AT: 20,
  BA: 20,
  EE: 20,
  KZ: 20,
  LT: 20,
  LU: 20,
  CR: 21,
  CH: 21,
  HR: 21,
  LI: 21,
  LV: 21,
  BG: 22,
  BH: 22,
  DE: 22,
  GB: 22,
  GE: 22,
  IE: 22,
  ME: 22,
  RS: 22,
  AE: 23,
  GI: 23,
  IL: 23,
  AD: 24,
  CZ: 24,
  ES: 24,
  MD: 24,
  PK: 24,
  RO: 24,
  SA: 24,
  SE: 24,
  SK: 24,
  VG: 24,
  TN: 24,
  PT: 25,
  IS: 26,
  TR: 26,
  FR: 27,
  GR: 27,
  IT: 27,
  MC: 27,
  MR: 27,
  SM: 27,
  AL: 28,
  AZ: 28,
  CY: 28,
  DO: 28,
  GT: 28,
  HU: 28,
  LB: 28,
  PL: 28,
  BR: 29,
  PS: 29,
  KW: 30,
  MU: 30,
  MT: 31,
};

function isValid(iban) {
  iban = iban.replace(/\s/g, "").toUpperCase();
  if (!iban.match(/^[\dA-Z]+$/)) return false;
  var len = iban.length;
  if (len != ibanLen[iban.substr(0, 2)]) return false;
  iban = iban.substr(4) + iban.substr(0, 4);
  for (var s = "", i = 0; i < len; i += 1) s += parseInt(iban.charAt(i), 36);
  for (var m = s.substr(0, 15) % 97, s = s.substr(15); s; s = s.substr(13))
    m = (m + s.substr(0, 13)) % 97;
  return m == 1;
}

function erzeugeKarte(mitgliedsname, mitgliedsnummer, büro) {
  // Diese Funktion erzeugt eine PNG-Grafik, die eine individuelle Saisonkarte darstellt
  // Der Rückgabewert dieser Funktion ist das File-Objekt der Grafik
  //
  //
  //
  // Schritt 1: Die Konfigurationsdaten werden eingelesen und die laufende Nummer um 1 erhöht
  var basisTabelle = SpreadsheetApp.openById(BASIS_DATA_ID).getSheets()[0];
  var jahr = basisTabelle.getRange("B1").getValue();
  var laufendeNummerZelle = basisTabelle.getRange(büro ? "C2" : "B2");
  var nummerSaisonkarte = laufendeNummerZelle.getValue();
  laufendeNummerZelle.setValue(nummerSaisonkarte + 1);
  var startGueltigkeit = basisTabelle.getRange("B3").getValue();
  var endeGueltigkeit = basisTabelle.getRange("B4").getValue();

  // Schritt 2: Für Ordner und Dateinamen werden Skript-Objekte erzeugt
  var temporaryFolder = DriveApp.getFolderById(TEMPORARY_FOLDER_ID);
  var resultFolder = DriveApp.getFolderById(RESULT_FOLDER_ID);
  var resultFilename =
    "Saisonkarte " + jahr + "-#" + nummerSaisonkarte + " " + mitgliedsname;

  // Schritt 3: Damit das Template nicht ruiniert wird, wird eine Kopie erzeugt und diese Kopie in
  // der Slides-App geöffnet, damit sie modifiziert werden kann
  var copyFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(
      "WorkingCopyOfTemplate",
      temporaryFolder,
    ),
    copyId = copyFile.getId(),
    copyDoc = SlidesApp.openById(copyId);

  // Schritt 4: Die Kopie des Templates wird mit den notwendigen individuellen Daten versehen
  copyDoc.replaceAllText("%Jahr", jahr);
  copyDoc.replaceAllText("%Nummer", nummerSaisonkarte);
  copyDoc.replaceAllText("%Mitgliedsname", mitgliedsname);
  copyDoc.replaceAllText("%Mitgliedsnummer", mitgliedsnummer);
  copyDoc.replaceAllText("%Ab", startGueltigkeit);
  copyDoc.replaceAllText("%Bis", endeGueltigkeit);
  copyDoc.saveAndClose();

  // Schritt 5: Leider ist die Slides-App nicht in der Lage, direkt eine PNG-Datei aus Slide 1 zu erzeugen
  //            Deshalb muss ein Umweg gegangen werden. Der Slide muss über eine HTML-URL vom Google-Server
  //            heruntergeladen werden, der Server liefert dann Web-fähige Daten, die unter anderem eine
  //            brauchbare Grafik beinhalten.
  var slides = copyDoc.getSlides();
  var slide_ID = slides[0].getObjectId();
  var presentation_ID = copyFile.getId();
  var temporaryResult = downloadSlide(
    resultFilename,
    presentation_ID,
    slide_ID,
  );

  // Schritt 6: Leider landen die Daten immer im Root-Folder und eine Datei kann per Skript nicht
  //            so einfach in einen anderen Ordner verschoben werden.
  //            Deshalb wird der Download in den Zielorder kopiert und dann gelöscht.
  var resultGraphic = temporaryResult.makeCopy(resultFolder);
  temporaryResult.setTrashed(true);
  copyFile.setTrashed(true);
  return resultGraphic;

  // Fertig, die Grafik ist im Zielordner
}

function downloadSlide(name, presentationId, slideId) {
  var url =
    "https://docs.google.com/presentation/d/" +
    presentationId +
    "/export/png?id=" +
    presentationId +
    "&pageid=" +
    slideId;
  var options = {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken(),
    },
  };
  var response = UrlFetchApp.fetch(url, options);
  var image = response.getAs(MimeType.PNG);
  image.setName(name);
  return DriveApp.createFile(image);
}

function sendWrongIbanEmail(empfaenger, iban) {
  var subject = "Falsche IBAN";
  var body =
    "Die von Ihnen bei der Bestellung der Saisonkarte übermittelte IBAN " +
    iban +
    " ist leider falsch! Bitte wiederholen Sie die Anmeldung mit einer korrekten IBAN.";
  GmailApp.sendEmail(empfaenger, subject, body);
}

function sendNotFoundEmail(empfaenger, name, num) {
  var subject = "Saisonkarte nicht gefunden";
  var body =
    "Eine Saisonkarte für den Namen " +
    name +
    " und die Mitgliedsnummer " +
    num +
    " konnte leider nicht gefunden werden!";
  GmailApp.sendEmail(empfaenger, subject, body);
}

function sendImgEmail(empfaenger, img) {
  var subject = "Ihre Saisonkarte";
  var body =
    "Anbei Ihre Saisonkarte. Bitte auf dem Handy speichern oder ausdrucken. Hinweis: Manchmal fehlt beim Anhang die korrekte Dateiendung .png Falls dies der Fall sein sollte, kann das Mail-Programm die Datei nicht korrekt darstellen. Bitte ggf. manuell korrigieren oder uns Bescheid geben.";
  GmailApp.sendEmail(empfaenger, subject, body, { attachments: img });
}
