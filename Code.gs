/*******************************
 * CONFIG
 *******************************/
const SHEET_INPUT = "INPUT_LINK";
const SHEET_PARSED = "PARSED_DATA";
const START_ROW = 2;
const BATCH_SIZE = 50;

/*******************************
 * MENU (CUMA 2)
 *******************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ PDF Automation")
    .addItem("Open Panel", "showPDFUI")
    .addItem("Reset RAW & Parsed", "resetAll")
    .addToUi();
}

/*******************************
 * OPEN UI
 *******************************/
function showPDFUI() {
  const html = HtmlService.createHtmlOutputFromFile('pdf_ui')
    .setWidth(700)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'PDF Automation');
}

/*******************************
 * PROGRESS (REALTIME)
 *******************************/
function getProgress() {
  const props = PropertiesService.getScriptProperties();

  const current = Number(props.getProperty("CURRENT_COUNT")) || 0;
  const total = Number(props.getProperty("TOTAL_COUNT")) || 0;

  if (total === 0) return "Status: Idle";

  const percent = Math.floor((current / total) * 100);

  return `Processing ${current}/${total} (${percent}%)`;
}

/*******************************
 * STEP 1 — EXTRACT
 *******************************/
function runExtractRawPDF() {

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_INPUT);
  if (!sheet) throw new Error("Sheet INPUT_LINK not found");

  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return;

  const props = PropertiesService.getScriptProperties();

  let start = Number(props.getProperty("RAW_INDEX")) || START_ROW;

  const links = sheet.getRange(START_ROW,1,lastRow-1).getValues().flat();
  const validLinks = links.filter(l => l);
  const totalCount = validLinks.length;

  props.setProperty("TOTAL_COUNT", totalCount);

  let processedCount = Number(props.getProperty("CURRENT_COUNT")) || 0;

  const end = Math.min(start + BATCH_SIZE - 1, lastRow);

  for (let row = start; row <= end; row++) {

    const link = sheet.getRange(row, 1).getValue();
    if (!link) continue;

    sheet.getRange(row, 3).setValue("⏳ EXTRACTING");

    try {
      const fileId = extractFileId(link);
      const text = extractPdfText(fileId);

      sheet.getRange(row, 2).setValue(text);
      sheet.getRange(row, 3).setValue("DONE");

      processedCount++;
      props.setProperty("CURRENT_COUNT", processedCount);

    } catch (e) {
      sheet.getRange(row, 3).setValue("❌ " + e.message);
    }
  }

  if (end < lastRow) {

    props.setProperty("RAW_INDEX", end + 1);

    deleteTriggerByName_("runExtractRawPDF");

    ScriptApp.newTrigger("runExtractRawPDF")
      .timeBased()
      .after(5000)
      .create();

  } else {

    props.deleteProperty("RAW_INDEX");
    props.deleteProperty("CURRENT_COUNT");
    props.deleteProperty("TOTAL_COUNT");

    deleteTriggerByName_("runExtractRawPDF");

    SpreadsheetApp.getUi().alert("🎉 Extract selesai!");
  }
}

/*******************************
 * STEP 2 — PARSE
 *******************************/
function runParseData() {

  const ss = SpreadsheetApp.getActive();
  const input = ss.getSheetByName(SHEET_INPUT);
  const parsed = ss.getSheetByName(SHEET_PARSED);

  const lastRow = input.getLastRow();
  parsed.getRange("A2:I").clearContent();

  let output = [];

  for (let row = START_ROW; row <= lastRow; row++) {

    const link = input.getRange(row,1).getValue();
    const rawText = input.getRange(row,2).getValue();
    if (!rawText) continue;

    let amount="", currency="", campaign="", kol="", sow="", bankName="", accountNumber="";

    const headerMatch = rawText.match(/sent you\s+([\d,]+)\s+([A-Z]{3})/i);
    if (headerMatch) {
      amount = headerMatch[1];
      currency = headerMatch[2];
    }

    const descMatch = rawText.match(/Bank Transfer([\s\S]*?)Xendit/i);
    if (descMatch) {
      let clean = descMatch[1].replace(/\s+/g," ").trim();
      const kolMatch = clean.match(/--\s*(.+?)\s*\((\d+)\)/);

      if (kolMatch) {
        kol = kolMatch[1];
        sow = kolMatch[2];
        campaign = clean.split("--")[0].trim();
      }
    }

    const lines = rawText.split("\n").map(l=>l.trim()).filter(l=>l);

    for (let i=0;i<lines.length;i++){
      if (/^[A-Z0-9]{15,}$/.test(lines[i])){
        if (lines[i+1]) bankName = lines[i+1];
        if (lines[i+3] && /^\d+$/.test(lines[i+3])) accountNumber = lines[i+3];
        break;
      }
    }

    output.push([link,amount,currency,campaign,kol,sow,bankName,accountNumber]);
  }

  if (output.length){
    parsed.getRange(2,1,output.length,8).setValues(output);
  }

  SpreadsheetApp.getUi().alert("✅ Parse selesai");
}

/*******************************
 * STEP 3 — MATCH
 *******************************/
function runMVPMatch() {

  const ss = SpreadsheetApp.getActive();

  const parsed = ss.getSheetByName("PARSED_DATA");
  const mvp = ss.getSheetByName("MVP_MATCH");

  const parsedData = parsed
    .getRange(2,1,parsed.getLastRow()-1,8)
    .getValues();

  const mvpData = mvp
    .getRange(2,1,mvp.getLastRow()-1,8)
    .getValues();

  function normalize(text){

    return String(text || "")
      .toLowerCase()
      .replace(/đ/g,"d")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g,"")
      .replace(/[.,\s]/g,"")
      .trim();
  }

  let output = [];
  let parsedStatus = [];

  // default kosong semua
  for (let x = 0; x < parsedData.length; x++) {
    parsedStatus.push([""]);
  }

  for (let i = 0; i < mvpData.length; i++) {

    let row3 = mvpData[i];

    let resultLink = "";
    let note = "KOL NOT READY TO PAYMENT";

    for (let j = 0; j < parsedData.length; j++) {

      let row2 = parsedData[j];

      let match =

        normalize(row3[0]) === normalize(row2[3]) &&
        normalize(row3[1]) === normalize(row2[4]) &&
        normalize(row3[2]) === normalize(row2[7]) &&
        normalize(row3[3]) === normalize(row2[6]) &&
        Number(row3[4]) === Number(row2[1]) &&
        normalize(row3[5]) === normalize(row2[5]);

      if (match){

        resultLink = row2[0];
        note = "";

        // ==========================
        // KASIH STATUS CENTANG
        // DI PARSED_DATA KOLOM I
        // ==========================
        parsedStatus[j] = ["✅"];

        break;
      }
    }

    output.push([resultLink, note]);
  }

  // ==========================
  // OUTPUT MATCH RESULT
  // ==========================
  if (output.length){

    mvp
      .getRange(2,7,output.length,2)
      .setValues(output);
  }

  // ==========================
  // OUTPUT STATUS KE KOLOM I
  // ==========================
  if (parsedStatus.length){

    parsed
      .getRange(2,9,parsedStatus.length,1)
      .setValues(parsedStatus);
  }

  SpreadsheetApp.getUi()
    .alert("🚀 Match selesai");
}

/*******************************
 * RESET
 *******************************/
function resetAll() {
  const ss = SpreadsheetApp.getActive();

  ss.getSheetByName(SHEET_INPUT)
    .getRange("B2:C")
    .clearContent();

  ss.getSheetByName(SHEET_PARSED).clear();

  PropertiesService.getScriptProperties().deleteAllProperties();

  SpreadsheetApp.getUi().alert("🔄 Reset selesai");
}

/*******************************
 * HELPERS
 *******************************/
function extractFileId(url) {
  const m = url.match(/[-\w]{25,}/);
  if (!m) throw new Error("INVALID LINK");
  return m[0];
}

function extractPdfText(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();

  const doc = Drive.Files.insert(
    { mimeType: MimeType.GOOGLE_DOCS, title: "TMP" },
    blob
  );

  const text = DocumentApp.openById(doc.id).getBody().getText();

  DriveApp.getFileById(doc.id).setTrashed(true);

  return text;
}

function deleteTriggerByName_(name) {
  ScriptApp.getProjectTriggers().forEach(t=>{
    if (t.getHandlerFunction()===name){
      ScriptApp.deleteTrigger(t);
    }
  });
}
