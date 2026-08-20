/**
 * SRx SoS Tracks Applet — Google Apps Script Web App
 *
 * Reads HCC V2MOM data from rows 25-33, columns C, E, L, M, N
 * and renders a track dashboard similar to gus-apps scale360-sos-tracks.
 *
 * Setup:
 * 1. Go to script.google.com → New Project
 * 2. Paste this file as Code.gs
 * 3. Create a new HTML file named "Index" and paste Index.html
 * 4. Deploy → New Deployment → Web App → Anyone with link
 * 5. Authorize when prompted (grants read access to the spreadsheet)
 */

const SPREADSHEET_ID = '1ufYd6q4YhSUDFFRJpXgjuwoPcKKugd0SB0oOHlj06fM';
const SHEET_GID = 649674725;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SRx SoS Tracks')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTrackData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();

  // Find sheet by gid
  let sheet;
  for (const s of sheets) {
    if (s.getSheetId() === SHEET_GID) {
      sheet = s;
      break;
    }
  }
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }

  // Rows 25-33, Columns C(3), E(5), L(12), M(13), N(14)
  const startRow = 25;
  const endRow = 33;
  const tracks = [];

  for (let row = startRow; row <= endRow; row++) {
    const measure = sheet.getRange(row, 3).getValue();  // C = measure name
    const owner = sheet.getRange(row, 5).getValue();    // E = owner
    const status = sheet.getRange(row, 12).getValue();  // L = status
    const progress = sheet.getRange(row, 13).getValue(); // M = progress/comments
    const notes = sheet.getRange(row, 14).getValue();   // N = notes/updates

    if (measure) {
      tracks.push({
        measure: String(measure),
        owner: String(owner),
        status: String(status),
        progress: String(progress),
        notes: String(notes)
      });
    }
  }

  return tracks;
}
