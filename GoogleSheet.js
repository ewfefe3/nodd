// Simplified and cleaned-up version of the script
const { google } = require('googleapis');
const path = require('path');
module.exports = { Gpush };
const auth = new google.auth.GoogleAuth({
  keyFile: 'ky.json',
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

let client;
const sheets = google.sheets('v4');
const spreadsheetId = '1btgB5iS7scCgn8j_YwjrpCOFKhPw7p2EDWSTA4tHaN4';
let lastID = 0;

async function start(range) {
  client = await auth.getClient();
  const res = await sheets.spreadsheets.values.get({
    auth: client,
    spreadsheetId,
    range,
  });
  return res.data.values || [];
}

async function sheetExists(name) {
  const client = await auth.getClient();
  const sheets = google.sheets({ version: 'v4', auth: client });
  const response = await sheets.spreadsheets.get({ spreadsheetId });
  return response.data.sheets.some(sheet => sheet.properties.title === name);
}

async function createSheetIfNeeded(name) {
  const exists = await sheetExists(name);
  if (!exists) {
    await sheets.spreadsheets.batchUpdate({
      auth: await auth.getClient(),
      spreadsheetId,
      resource: {
        requests: [{ addSheet: { properties: { title: name } } }],
      },
    });
  }
}

async function readFirstRow(sheet) {
  const rows = await start(`${sheet}!A1:Z1`);
  return rows[0] || [];
}

async function readLastID(sheet) {
  const rows = await start(`${sheet}!A1:A99999`);
  lastID = rows.length ? parseInt(rows.at(-1)) + 1 : 1;
}

async function appendRow(sheet, row, addID = false) {
  await readLastID(sheet);
  if (addID) row.unshift(`${lastID}`);
  await sheets.spreadsheets.values.append({
    auth: await auth.getClient(),
    spreadsheetId,
    range: `${sheet}!A1:W99`,
    valueInputOption: 'USER_ENTERED',
    resource: { values: [row] },
  });
}

function nowtime() {
  const now = new Date();
  return `${now.getHours()}:${now.getMinutes()}:${now.getSeconds()}`;
}

function nowdate() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
}

async function Gpush({ sheet, a, forc = false }) {
  const vals = a.split('|a|');
  const rowMap = new Map();
  const NULLS = [];
  for (const item of vals) {
    const [key, val] = item.split('::');
    if (!val) NULLS.push(key);
    rowMap.set(key, val);
  }

  if (NULLS.length) return `error::100::input ${NULLS} not completed`;

  await createSheetIfNeeded(sheet);
  await readLastID(sheet);

  let Frow = await readFirstRow(sheet);
  if (!Frow.length) {
    const headers = ['ID', ...[...rowMap.keys()]];
    await appendRow(sheet, headers, false);
    Frow = headers;
  }

  const replacements = { '@time': nowtime(), '@date': nowdate() };
  const row = Frow.slice(1).map(key => replacements[rowMap.get(key)] || rowMap.get(key));
  row.unshift(`${lastID}`);

  const ex = await userExists(rowMap.get('Username'), { mrga: 'points:==13300k' });
  if (forc || (!forc && !ex[0])) {
    await appendRow(sheet, row, false);
  }
  return `now put [${a}] in (${sheet}) is Ok`;
}

async function getFilteredUsers(key) {
  const rows = await start('Users!A1:W100');
  const [headers, ...data] = rows;
  return data.map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] || '']))).filter(u => u.Username === key);
}

async function userExists(key, { mrga } = {}) {
  mrga = mrga || `Username:==${key}`;
  const users = await getFilteredUsers(key);
  const field = mrga.split(':==')[0];
  const value = mrga.split(':==')[1];
  const match = users.length && users[0][field] === value;
  return [!!users.length, match, users];
}
/*
Gpush({
  sheet: 'Users 02',
  a: 'Username::Omar|a|email::yacoubi|a|points::13300k|a|Time::@time|a|Date::@date|a|info::no ealse|a|class::proumear',
  forc: true,
});
*/
