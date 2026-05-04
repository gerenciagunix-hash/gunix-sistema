const { google } = require('googleapis');
const SPREADSHEET_ID = '1r4oPE6n7Q3HzpBBZlzVeAsgdjL2liRXfgkU8joWOQxA';
const JOTFORM_API_KEY = 'a2bd6cdf693aaa47a0f85b2064702c5c';
const JOTFORM_FORM_ID = '260543438713659';

function getAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS || '{}');
  return new google.auth.GoogleAuth({credentials, scopes: ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']});
}

async function getSheets() {
  const auth = await getAuth();
  return google.sheets({ version: 'v4', auth });
}

async function getRange(sheetName) {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.values.get({spreadsheetId: SPREADSHEET_ID, range: sheetName});
  const rows = res.data.values || [];
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => { const obj = {}; headers.forEach((h,i) => { obj[h] = row[i] || ''; }); return obj; });
}

async function appendRow(sheetName, values) {
  const sheets = await getSheets();
  await sheets.spreadsheets.values.append({spreadsheetId: SPREADSHEET_ID, range: sheetName, valueInputOption: 'RAW', resource: { values: [values] }});
}

async function updateCell(sheetName, rowIndex, colIndex, value) {
  const sheets = await getSheets();
  const col = String.fromCharCode(65 + colIndex);
  await sheets.spreadsheets.values.update({spreadsheetId: SPREADSHEET_ID, range: sheetName+'!'+col+(rowIndex+1), valueInputOption: 'RAW', resource: { values: [[value]] }});
}

module.exports = { getRange, appendRow, updateCell, SPREADSHEET_ID, JOTFORM_API_KEY, JOTFORM_FORM_ID };