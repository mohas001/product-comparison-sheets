function showTop5Products() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Products"); // Name deiner Haupttabelle
  const sheetName = "Top 5 Products";

  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Products' not found.");
    return;
  }

  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();

  const header = data[0];
  const rows = data.slice(1);

  // Nur Zeilen mit gültigem Value Score (Zahl) filtern
  const filtered = rows.filter(row => typeof row[6] === "number" && !isNaN(row[6]));

  if (filtered.length === 0) {
    SpreadsheetApp.getUi().alert("Keine gültigen Daten gefunden (leere oder fehlerhafte 'Value Score'-Werte).");
    return;
  }

  // Nach Value Score absteigend sortieren
  const sorted = filtered.sort((a, b) => b[6] - a[6]);
  const top5 = sorted.slice(0, 5);

  // Neuen oder existierenden Tab holen oder erstellen
  let targetSheet = ss.getSheetByName(sheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(sheetName);
  } else {
    targetSheet.clear(); // Vorherige Inhalte löschen
  }

  // Header und Top 5 reinschreiben
  targetSheet.getRange(1, 1, 1, header.length).setValues([header]);
  targetSheet.getRange(2, 1, top5.length, header.length).setValues(top5);

  // Formatierung der Kopfzeile übernehmen
  const sourceHeaderRange = sourceSheet.getRange(1, 1, 1, header.length);
  const targetHeaderRange = targetSheet.getRange(1, 1, 1, header.length);

  // Hintergrundfarbe, fett, zentriert übernehmen
  targetHeaderRange.setFontWeight("bold");
  targetHeaderRange.setBackground(sourceHeaderRange.getBackgrounds()[0][0]); // gleiche Farbe
  targetHeaderRange.setHorizontalAlignment("center");

  // Optionale Auto-Anpassung der Spaltenbreite
  targetSheet.autoResizeColumns(1, header.length);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Produkt-Tools")
    .addItem("Top 5 Products anzeigen", "showTop5Products")
    .addToUi();
}
