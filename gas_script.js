// ============================================
// FC熊野 試合結果報告書 - Google Apps Script v3
// ============================================

const SPREADSHEET_ID = '1OsLyIAeqP8MPGWRpQoSLM4Wrjgn0nGCUrQfeOosAFZg';
const SHEET_PLAYERS = '選手';
const SHEET_MATCHES = '試合';
const SHEET_SUMMARY = '年度累積';

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const p = e.parameter;
  let result;
  try {
    if      (p.action === 'getPlayers')     result = getPlayers();
    else if (p.action === 'savePlayers')    result = savePlayers(JSON.parse(p.data));
    else if (p.action === 'getMatches')     result = getMatches();
    else if (p.action === 'saveMatches')    result = saveMatches(JSON.parse(p.data));
    else if (p.action === 'getSummary')     result = getSummary(p.nendo);
    else if (p.action === 'addMatchRecord') result = addMatchRecord(JSON.parse(p.data));
    else if (p.action === 'getAttendance')    result = getAttendance();
    else if (p.action === 'saveAttendance')   result = saveAttendance(JSON.parse(p.data));
    else if (p.action === 'clearAttendance')  result = clearAttendance();
    else if (p.action === 'deleteAttendanceRow') result = deleteAttendanceRow(p.sid);
    else result = { error: 'Unknown action' };
  } catch(err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ===== 選手シート =====
function getPlayers() {
  const sheet = getSheet(SHEET_PLAYERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(r => r[0]).map(r => ({ id: r[0], name: r[1] }));
}

function savePlayers(players) {
  const sheet = getSheet(SHEET_PLAYERS);
  sheet.clearContents();
  sheet.getRange(1,1,1,2).setValues([['ID', '選手名']]);
  sheet.getRange(1,1,1,2).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
  if (players.length) {
    sheet.getRange(2,1,players.length,2).setValues(players.map(p=>[p.id, p.name]));
  }
  return { success: true };
}

// ===== 試合シート（アプリ内部用） =====
function getMatches() {
  const sheet = getSheet(SHEET_MATCHES);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  // ヘッダー行をスキップしてJSONに復元
  return data.slice(1).filter(r => r[0]).map(r => ({
    id:       r[0],
    num:      r[1],
    opponent: r[2] || '',
    venue:    r[3] || '',
    scores: {
      h1_k: r[4]||0, h2_k: r[5]||0, pk_k: r[6]||0,
      h1_o: r[7]||0, h2_o: r[8]||0, pk_o: r[9]||0
    },
    scorers: r[10] ? JSON.parse(r[10]) : []
  }));
}

function saveMatches(matches) {
  const sheet = getSheet(SHEET_MATCHES);
  sheet.clearContents();

  // ヘッダー
  const headers = ['ID','第N試合','試合名','相手','前半(熊野)','後半(熊野)','PK(熊野)','前半(相手)','後半(相手)','PK(相手)','得点者(JSON)'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(1,1,1,headers.length)
    .setFontWeight('bold')
    .setBackground('#2d7a4f')
    .setFontColor('white');

  if (!matches.length) return { success: true };

  // 選手IDから名前を引くマップを作成
  const players = getPlayers();
  const playerMap = {};
  players.forEach(p => { playerMap[String(p.id)] = p.name; });

  const rows = matches.map(m => {
    // 得点者を「名前（前半）、名前（後半）」形式に変換
    const scorerNames = (m.scorers || []).map(s => {
      const name = playerMap[String(s.playerId)] || '不明';
      const period = s.period === 'h1' ? '前半' : s.period === 'h2' ? '後半' : 'PK';
      return name + '(' + period + ')';
    }).join('、');
    return [
      m.id,
      '第' + m.num + '試合',
      m.venue || '',
      m.opponent || '',
      m.scores.h1_k || 0,
      m.scores.h2_k || 0,
      m.scores.pk_k || 0,
      m.scores.h1_o || 0,
      m.scores.h2_o || 0,
      m.scores.pk_o || 0,
      scorerNames
    ];
  });

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // 列幅を自動調整
  sheet.autoResizeColumns(1, headers.length);

  // 結果に応じて行に色付け
  rows.forEach((row, i) => {
    const kumano = (row[4]||0) + (row[5]||0);
    const aite   = (row[7]||0) + (row[8]||0);
    const bg = kumano > aite ? '#e8f5ee' : kumano < aite ? '#fdecea' : '#fef9e7';
    sheet.getRange(i+2, 1, 1, headers.length).setBackground(bg);
  });

  return { success: true };
}

// ===== 年度累積シート =====
// 列: 年度 | 日付 | 学年 | 勝 | 負 | 分 | 得点 | 失点 | 得失点差 | 選手名... (得点数)
function getSummary(nendo) {
  const sheet = getSheet(SHEET_SUMMARY);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).filter(r => r[0] && String(r[0]) !== '合計')
    .filter(r => !nendo || r[0] == nendo)
    .map(r => {
      const scorers = {};
      for (let i = 9; i < headers.length; i++) {
        if (headers[i] && r[i]) scorers[headers[i]] = r[i];
      }
      return { nendo:r[0], date:r[1], grade:r[2], win:r[3], lose:r[4], draw:r[5], gf:r[6], ga:r[7], scorers };
    });
}

function addMatchRecord(data) {
  const sheet = getSheet(SHEET_SUMMARY);
  const allData = sheet.getDataRange().getValues();
  const baseHeaders = ['年度','日付','学年','勝','負','分','得点','失点','得失点差'];
  let headers;

  if (allData.length === 0 || !allData[0][0]) {
    headers = [...baseHeaders];
    sheet.clearContents();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
  } else {
    headers = allData[0].map(String);
  }

  // 新しい選手カラムを追加
  const scorerNames = Object.keys(data.scorers);
  scorerNames.forEach(name => {
    if (!headers.includes(name)) {
      headers.push(name);
      sheet.getRange(1, headers.length).setValue(name)
           .setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    }
  });

  // 合計行を一旦除去してデータ行のみにする
  const lastRow = sheet.getLastRow();
  if (lastRow > 1 && String(sheet.getRange(lastRow,1).getValue()) === '合計') {
    sheet.deleteRow(lastRow);
  }

  // 行データ作成
  const diff = data.gf - data.ga;
  const row = [
    data.nendo, data.date, data.grade,
    data.win, data.lose, data.draw,
    data.gf, data.ga,
    diff >= 0 ? '+' + diff : String(diff)
  ];
  for (let i = 9; i < headers.length; i++) {
    row.push(data.scorers[headers[i]] || 0);
  }
  sheet.appendRow(row);

  // 合計行を末尾に追加
  updateSummaryTotal(sheet, headers);

  return { success: true };
}

function updateSummaryTotal(sheet, headers) {
  const allData = sheet.getDataRange().getValues();
  const dataRows = allData.slice(1).filter(r => r[0] && String(r[0]) !== '合計');
  if (!dataRows.length) return;

  const totalRow = ['合計', '', ''];
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[3])||0),0));
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[4])||0),0));
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[5])||0),0));
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[6])||0),0));
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[7])||0),0));
  const diff = totalRow[6] - totalRow[7];
  totalRow.push(diff >= 0 ? '+' + diff : String(diff));
  for (let i = 9; i < headers.length; i++) {
    totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[i])||0),0));
  }

  sheet.appendRow(totalRow);
  const newLast = sheet.getLastRow();
  sheet.getRange(newLast, 1, 1, totalRow.length).setFontWeight('bold').setBackground('#e8f5ee');
}

// ===== 出欠確認シート =====
const SHEET_ATTENDANCE = '出欠確認';

function getAttendance() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {};
  // 1行目: ヘッダー（id, 保護者名...）
  const headers = data[0];
  const result = {};
  data.slice(1).forEach(row => {
    if (!row[0]) return;
    const id = row[0];
    result[id] = {};
    headers.slice(1).forEach((h, i) => {
      if (h) result[id][h] = row[i+1] || '';
    });
  });
  return result;
}

function saveAttendance(attData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) sheet = ss.insertSheet(SHEET_ATTENDANCE);
  sheet.clearContents();

  const ids = Object.keys(attData);
  if (!ids.length) return { success: true };

  // 家族名を動的に収集
  const parentsSet = new Set();
  ids.forEach(id => Object.keys(attData[id] || {}).forEach(p => parentsSet.add(p)));
  const parents = Array.from(parentsSet);

  const headers = ['日程ID', ...parents];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');

  const rows = ids.map(id => {
    const att = attData[id] || {};
    return [id, ...parents.map(p => att[p] || '')];
  });
  sheet.getRange(2,1,rows.length,headers.length).setValues(rows);

  rows.forEach((row, ri) => {
    parents.forEach((p, pi) => {
      const v = row[pi+1];
      const cell = sheet.getRange(ri+2, pi+2);
      if (v==='◯') cell.setBackground('#e8f5ee').setFontColor('#1a7a3f');
      else if (v==='×') cell.setBackground('#fdecea').setFontColor('#c0392b');
      else if (v==='△') cell.setBackground('#fef9e7').setFontColor('#b7950b');
    });
  });

  return { success: true };
}

// ===== 出欠データ全削除 =====
function clearAttendance() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (sheet) {
    sheet.clearContents();
    sheet.clearFormats();
  }
  return { success: true };
}

// ===== 出欠データ1行削除（日程ID指定） =====
function deleteAttendanceRow(sid) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ATTENDANCE);
  if (!sheet) return { success: true };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(sid)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return { success: true };
}
