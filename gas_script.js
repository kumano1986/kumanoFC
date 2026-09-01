// ============================================
// FC熊野 試合結果報告書 - Google Apps Script v3
// ============================================

const SPREADSHEET_ID = '1OsLyIAeqP8MPGWRpQoSLM4Wrjgn0nGCUrQfeOosAFZg';
const SHEET_PLAYERS = '選手';
const SHEET_MATCHES = '試合';
const SHEET_SUMMARY = '年度累積';
const SHEET_ATTENDANCE = '出欠確認';
const SHEET_MASTER = 'マスタ'; // 日程・家族・会場・試合名を1シートにJSON保存
const SHEET_LOG = 'ログ'; // 出欠・配車の変更履歴（追記専用・読み出さない）
const SHEET_ARCHIVE = 'アーカイブ済み日程'; // アーカイブした日程を退避（マスタのサイズ削減用）

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  // GETはe.parameter、POSTのFormDataもe.parameterで取得可能
  const p = e.parameter || {};
  // POSTのJSONボディの場合はe.postDataから取得
  if (!p.action && e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.assign(p, body);
    } catch(err) {}
  }
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
    else if (p.action === 'appendLog')      result = appendLog(JSON.parse(p.data));
    // マスタデータ
    else if (p.action === 'getMaster')      result = getMaster();
    else if (p.action === 'saveMaster')     result = saveMaster(JSON.parse(p.data));
    else if (p.action === 'deleteScheduleFromMaster')   result = deleteScheduleFromMaster(p.sid);
    else if (p.action === 'clearAllSchedulesFromMaster') result = clearAllSchedulesFromMaster();
    else if (p.action === 'archiveSchedules')  result = archiveSchedules(JSON.parse(p.ids));
    else if (p.action === 'getMasterSize')     result = getMasterSize();
    else if (p.action === 'removeFamilyFromMaster')    result = removeFamilyFromMaster(p.name);
    else if (p.action === 'removePlayerFromMaster')    result = removePlayerFromMaster(p.id);
    else if (p.action === 'removeVenueFromMaster')     result = removeVenueFromMaster(p.name);
    else if (p.action === 'removeMatchNameFromMaster') result = removeMatchNameFromMaster(p.name);
    else if (p.action === 'removeOpponentFromMaster')  result = removeOpponentFromMaster(p.name);
    else if (p.action === 'renameOpponentInMaster')    result = renameOpponentInMaster(p.oldName, p.newName);
    else result = { error: 'Unknown action' };
  } catch(err) { result = { error: err.toString() }; }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ===== マスタデータ（日程・家族・会場・試合名） =====
function getMaster() {
  const sheet = getSheet(SHEET_MASTER);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2 || !data[1][0]) return {};
  try { return JSON.parse(data[1][0]); } catch(e) { return {}; }
}

function saveMaster(obj) {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(15000); // 最大15秒待機（同時書き込みの競合を防ぐ）
  if (!gotLock) {
    return { error: 'サーバーが混み合っています。もう一度お試しください。' };
  }
  try {
    const sheet = getSheet(SHEET_MASTER);
    const existing = getMaster();
    if (!existing._deleted) existing._deleted = {};

    // schedulesはID基準でマージ：他端末がまだ知らない日程を消さない。
    // ただし、削除マーカー（tombstone）がある項目は、他端末の古いデータで復活させない。
    if (Array.isArray(obj.schedules)) {
      const delMap = getTombstoneMap(existing, 'schedules');
      const existingSchedules = Array.isArray(existing.schedules) ? existing.schedules : [];
      const merged = {};
      existingSchedules.forEach(function(s){ if (s && s.id && !delMap[s.id]) merged[s.id] = s; });
      obj.schedules.forEach(function(s){ if (s && s.id && !delMap[s.id]) merged[s.id] = s; });
      obj.schedules = Object.values(merged);
    }

    // players：ID基準でマージ。受信データの順序（並び替え・編集を尊重）を基本にしつつ、
    // 受信データにないID（他端末がまだ知らない新規選手）は末尾に保持する。重複IDは除去する。
    // 削除マーカーがあるIDは復活させない。
    if (Array.isArray(obj.players)) {
      const delMap = getTombstoneMap(existing, 'players');
      const existingPlayers = Array.isArray(existing.players) ? existing.players : [];
      const seenIds = {};
      const mergedPlayers = [];
      obj.players.forEach(function(p){
        if (p && p.id != null && !seenIds[p.id] && !delMap[p.id]) { seenIds[p.id] = true; mergedPlayers.push(p); }
      });
      existingPlayers.forEach(function(p){
        if (p && p.id != null && !seenIds[p.id] && !delMap[p.id]) { seenIds[p.id] = true; mergedPlayers.push(p); }
      });
      obj.players = mergedPlayers;
    }

    // parents・venues・matchNames・opponents：文字列配列。受信データの順序を基本にしつつ、
    // 受信データにない既存の文字列（他端末がまだ知らない新規項目）は末尾に保持する。重複は除去する。
    // 削除マーカーがある文字列は復活させない。
    ['parents', 'venues', 'matchNames', 'opponents'].forEach(function(key) {
      if (Array.isArray(obj[key])) {
        const delMap = getTombstoneMap(existing, key);
        const existingArr = Array.isArray(existing[key]) ? existing[key] : [];
        const seen = {};
        const result = [];
        obj[key].forEach(function(v){ if (!seen[v] && !delMap[v]) { seen[v] = true; result.push(v); } });
        existingArr.forEach(function(v){ if (!seen[v] && !delMap[v]) { seen[v] = true; result.push(v); } });
        obj[key] = result;
      }
    });

    // carpool・duty：日程IDをキーとするネストオブジェクトなのでキー単位でマージ
    ['carpool', 'duty'].forEach(function(key) {
      if (obj[key] && typeof obj[key] === 'object') {
        const existingVal = (existing[key] && typeof existing[key] === 'object') ? existing[key] : {};
        const mergedVal = Object.assign({}, existingVal);
        Object.keys(obj[key]).forEach(function(sid) {
          if (key === 'carpool' && obj[key][sid] && typeof obj[key][sid] === 'object') {
            mergedVal[sid] = Object.assign({}, mergedVal[sid] || {}, obj[key][sid]);
          } else {
            mergedVal[sid] = obj[key][sid];
          }
        });
        obj[key] = mergedVal;
      }
    });

    // attPlayers・dutyTotal：家族名をキーとするフラットオブジェクトなのでマージ
    ['attPlayers', 'dutyTotal'].forEach(function(key) {
      if (obj[key] && typeof obj[key] === 'object') {
        const existingVal = (existing[key] && typeof existing[key] === 'object') ? existing[key] : {};
        obj[key] = Object.assign({}, existingVal, obj[key]);
      }
    });

    const mergedObj = Object.assign({}, existing, obj);

    sheet.clearContents();
    sheet.getRange(1,1).setValue('マスタデータ（JSON）');
    sheet.getRange(1,1).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    sheet.getRange(2,1).setValue(JSON.stringify(mergedObj));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 日程を1件だけ明示的に削除（saveMasterのマージでは削除できないため）
function deleteScheduleFromMaster(sid) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const existing = getMaster();
    if (Array.isArray(existing.schedules)) {
      existing.schedules = existing.schedules.filter(function(s){ return s.id !== sid; });
    }
    markTombstone(existing, 'schedules', sid);
    const sheet = getSheet(SHEET_MASTER);
    sheet.clearContents();
    sheet.getRange(1,1).setValue('マスタデータ（JSON）');
    sheet.getRange(1,1).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    sheet.getRange(2,1).setValue(JSON.stringify(existing));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 全日程を明示的に削除（配車・当番データも一緒にクリア）
function clearAllSchedulesFromMaster() {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const existing = getMaster();
    if (Array.isArray(existing.schedules)) {
      existing.schedules.forEach(function(s){ markTombstone(existing, 'schedules', s.id); });
    }
    existing.schedules = [];
    existing.carpool = {};
    existing.duty = {};
    const sheet = getSheet(SHEET_MASTER);
    sheet.clearContents();
    sheet.getRange(1,1).setValue('マスタデータ（JSON）');
    sheet.getRange(1,1).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    sheet.getRange(2,1).setValue(JSON.stringify(existing));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 日程をアーカイブ専用シートに退避し、マスタから完全に削除（マスタのサイズ肥大化を防ぐ）
function archiveSchedules(ids) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const existing = getMaster();
    if (!Array.isArray(existing.schedules)) return { success: true };

    const toArchive = existing.schedules.filter(function(s){ return ids.indexOf(s.id) >= 0; });

    if (toArchive.length > 0) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      let archiveSheet = ss.getSheetByName(SHEET_ARCHIVE);
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(SHEET_ARCHIVE);
        const headers = ['日程ID', '日付', '詳細JSON（日程・配車・当番）'];
        archiveSheet.getRange(1,1,1,headers.length).setValues([headers]);
        archiveSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#888').setFontColor('white');
        archiveSheet.setFrozenRows(1);
      }
      toArchive.forEach(function(s) {
        const extra = {
          schedule: s,
          carpool: (existing.carpool && existing.carpool[s.id]) || null,
          duty: (existing.duty && existing.duty[s.id]) || null
        };
        archiveSheet.appendRow([s.id, s.date, JSON.stringify(extra)]);
      });
    }

    // マスタから完全に削除（サイズ削減）
    existing.schedules = existing.schedules.filter(function(s){ return ids.indexOf(s.id) < 0; });
    ids.forEach(function(id){ markTombstone(existing, 'schedules', id); });
    if (existing.carpool) { ids.forEach(function(id){ delete existing.carpool[id]; }); }
    if (existing.duty)    { ids.forEach(function(id){ delete existing.duty[id]; }); }

    const sheet = getSheet(SHEET_MASTER);
    sheet.clearContents();
    sheet.getRange(1,1).setValue('マスタデータ（JSON）');
    sheet.getRange(1,1).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    sheet.getRange(2,1).setValue(JSON.stringify(existing));
    return { success: true, archivedCount: toArchive.length };
  } finally {
    lock.releaseLock();
  }
}

// マスタデータの現在サイズを返す（上限50,000文字への近さを確認するため）
function getMasterSize() {
  const sheet = getSheet(SHEET_MASTER);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2 || !data[1][0]) return { size: 0 };
  return { size: String(data[1][0]).length };
}

// ===== マスタ配列データの個別削除（saveMasterのマージでは削除できないため専用アクションで行う） =====
function removeFromMaster(key, matchFn, tombstoneKeyFn) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const existing = getMaster();
    if (Array.isArray(existing[key])) {
      existing[key].forEach(function(item){
        if (matchFn(item)) markTombstone(existing, key, tombstoneKeyFn(item));
      });
      existing[key] = existing[key].filter(function(item){ return !matchFn(item); });
    }
    const sheet = getSheet(SHEET_MASTER);
    sheet.clearContents();
    sheet.getRange(1,1).setValue('マスタデータ（JSON）');
    sheet.getRange(1,1).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    sheet.getRange(2,1).setValue(JSON.stringify(existing));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}
function removeFamilyFromMaster(name)    { return removeFromMaster('parents',    function(v){ return v === name; }, function(v){ return v; }); }
function removePlayerFromMaster(id)      { return removeFromMaster('players',    function(v){ return v && String(v.id) === String(id); }, function(v){ return v.id; }); }
function removeVenueFromMaster(name)     { return removeFromMaster('venues',     function(v){ return v === name; }, function(v){ return v; }); }
function removeMatchNameFromMaster(name) { return removeFromMaster('matchNames', function(v){ return v === name; }, function(v){ return v; }); }
function removeOpponentFromMaster(name)  { return removeFromMaster('opponents',  function(v){ return v === name; }, function(v){ return v; }); }

// 対戦相手の名称変更：旧名称をマスタから明示的に削除（saveMasterのマージ処理を安全に保つため）
function renameOpponentInMaster(oldName, newName) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const existing = getMaster();
    if (Array.isArray(existing.opponents)) {
      const idx = existing.opponents.indexOf(oldName);
      if (idx >= 0) existing.opponents[idx] = newName;
      else if (existing.opponents.indexOf(newName) < 0) existing.opponents.push(newName);
    }
    markTombstone(existing, 'opponents', oldName);
    const sheet = getSheet(SHEET_MASTER);
    sheet.clearContents();
    sheet.getRange(1,1).setValue('マスタデータ（JSON）');
    sheet.getRange(1,1).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
    sheet.getRange(2,1).setValue(JSON.stringify(existing));
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ===== 削除マーカー（tombstone）: 削除した項目が他端末の古いデータで復活するのを防ぐ =====
const TOMBSTONE_TTL_MS = 10 * 60 * 1000; // 10分間、削除を記憶しておく

function getTombstoneMap(existing, category) {
  if (!existing._deleted) existing._deleted = {};
  if (!existing._deleted[category]) existing._deleted[category] = {};
  const map = existing._deleted[category];
  // 期限切れのマーカーは掃除する（無限に蓄積させない）
  const now = Date.now();
  Object.keys(map).forEach(function(k){ if (now - map[k] > TOMBSTONE_TTL_MS) delete map[k]; });
  return map;
}

function markTombstone(existing, category, key) {
  const map = getTombstoneMap(existing, category);
  map[key] = Date.now();
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
  const headers = data[0].map(String);
  // カラム位置をヘッダー名から特定（新旧フォーマット両対応）
  const idx = {
    nendo: headers.indexOf('年度'),
    date: headers.indexOf('日付'),
    grade: headers.indexOf('学年'),
    matchName: headers.indexOf('試合名'),
    opponent: headers.indexOf('対戦相手'),
    win: headers.indexOf('勝'),
    lose: headers.indexOf('負'),
    draw: headers.indexOf('分'),
    gf: headers.indexOf('得点'),
    ga: headers.indexOf('失点'),
    diff: headers.indexOf('得失点差')
  };
  // 選手カラムは得失点差より後
  const scorerStart = idx.diff + 1;
  return data.slice(1).filter(r => r[idx.nendo] && String(r[idx.nendo]) !== '合計')
    .filter(r => !nendo || r[idx.nendo] == nendo)
    .map(r => {
      const scorers = {};
      for (let i = scorerStart; i < headers.length; i++) {
        if (headers[i] && r[i]) scorers[headers[i]] = r[i];
      }
      return {
        nendo: r[idx.nendo], date: r[idx.date], grade: r[idx.grade],
        matchName: idx.matchName >= 0 ? r[idx.matchName] : '',
        opponent: idx.opponent >= 0 ? r[idx.opponent] : '',
        win: r[idx.win], lose: r[idx.lose], draw: r[idx.draw],
        gf: r[idx.gf], ga: r[idx.ga], scorers
      };
    });
}

function addMatchRecord(data) {
  const sheet = getSheet(SHEET_SUMMARY);
  const allData = sheet.getDataRange().getValues();
  const baseHeaders = ['年度','日付','学年','試合名','対戦相手','勝','負','分','得点','失点','得失点差'];
  let headers;

  if (allData.length === 0 || !allData[0][0]) {
    headers = [...baseHeaders];
    sheet.clearContents();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
  } else {
    headers = allData[0].map(String);
    // 旧フォーマット（試合名・対戦相手なし）の場合はカラムを挿入
    if (headers.indexOf('試合名') < 0) {
      sheet.insertColumnsAfter(3, 2);
      sheet.getRange(1, 4).setValue('試合名').setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
      sheet.getRange(1, 5).setValue('対戦相手').setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');
      headers.splice(3, 0, '試合名', '対戦相手');
    }
  }

  // 試合名・対戦相手をまとめる（複数試合がある場合は連結）
  var matchNames = [], opponents = [];
  if (Array.isArray(data.matches)) {
    data.matches.forEach(function(m) {
      if (m.venue) matchNames.push(m.venue);
      if (m.opponent) opponents.push(m.opponent);
    });
  }
  var matchNameStr = matchNames.length ? matchNames.join(' / ') : '';
  var opponentStr = opponents.length ? opponents.join(' / ') : '';

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
    matchNameStr, opponentStr,
    data.win, data.lose, data.draw,
    data.gf, data.ga,
    diff >= 0 ? '+' + diff : String(diff)
  ];
  for (let i = 11; i < headers.length; i++) {
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

  const totalRow = ['合計', '', '', '', ''];
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[5])||0),0));  // 勝
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[6])||0),0));  // 負
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[7])||0),0));  // 分
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[8])||0),0));  // 得点
  totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[9])||0),0));  // 失点
  const diff = totalRow[8] - totalRow[9];
  totalRow.push(diff >= 0 ? '+' + diff : String(diff));
  for (let i = 11; i < headers.length; i++) {
    totalRow.push(dataRows.reduce((s,r)=>s+(parseInt(r[i])||0),0));
  }

  sheet.appendRow(totalRow);
  const newLast = sheet.getLastRow();
  sheet.getRange(newLast, 1, 1, totalRow.length).setFontWeight('bold').setBackground('#e8f5ee');
}

// ===== 出欠確認シート =====
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
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(15000);
  if (!gotLock) {
    return { error: 'サーバーが混み合っています。もう一度お試しください。' };
  }
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_ATTENDANCE);
    if (!sheet) sheet = ss.insertSheet(SHEET_ATTENDANCE);

    // 既存データを読み込んでマージ（他家族のデータ消失を防ぐ）
    const existing = {};
    const existingData = sheet.getDataRange().getValues();
    if (existingData.length > 1) {
      const exHeaders = existingData[0];
      existingData.slice(1).forEach(row => {
        if (!row[0]) return;
        const id = row[0];
        existing[id] = {};
        exHeaders.slice(1).forEach((h, i) => {
          if (h && row[i+1]) existing[id][h] = row[i+1];
        });
      });
    }

    // 受け取ったデータを既存にマージ（受信データを優先、ただし空文字は既存を保持しない＝削除扱い）
    const ids = Object.keys(attData);
    ids.forEach(id => {
      if (!existing[id]) existing[id] = {};
      const incoming = attData[id] || {};
      Object.keys(incoming).forEach(p => {
        const v = incoming[p];
        if (v === '' || v === null || v === undefined) {
          delete existing[id][p]; // 明示的に空にされた＝削除
        } else {
          existing[id][p] = v;
        }
      });
    });

    // 全日程ID・全家族を集約
    const allIds = Object.keys(existing);
    const parentsSet = new Set();
    allIds.forEach(id => Object.keys(existing[id] || {}).forEach(p => parentsSet.add(p)));
    const parents = Array.from(parentsSet);

    sheet.clearContents();
    if (!allIds.length || !parents.length) return { success: true };

    const headers = ['日程ID', ...parents];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#2d7a4f').setFontColor('white');

    const rows = allIds.map(id => {
      const att = existing[id] || {};
      return [id, ...parents.map(p => att[p] || '')];
    });
    sheet.getRange(2,1,rows.length,headers.length).setValues(rows);

    rows.forEach((row, ri) => {
      parents.forEach((p, pi) => {
        const v = row[pi+1];
        const cell = sheet.getRange(ri+2, pi+2);
        if (v==='◯') cell.setBackground('#e8f5ee').setFontColor('#1a7a3f');
        else if (v==='●') cell.setBackground('#e8eef5').setFontColor('#1a4a7a');
        else if (v==='×') cell.setBackground('#fdecea').setFontColor('#c0392b');
        else if (v==='△') cell.setBackground('#fef9e7').setFontColor('#b7950b');
      });
    });

    return { success: true };
  } finally {
    lock.releaseLock();
  }
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
  const lock = LockService.getScriptLock();
  const gotLock = lock.tryLock(15000);
  if (!gotLock) return { error: 'サーバーが混み合っています。もう一度お試しください。' };
  try {
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
  } finally {
    lock.releaseLock();
  }
}

// ===== ログ追記（追記専用・アプリからは読み出さない保険） =====
function appendLog(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    const headers = ['記録日時', '家族', '日程ID', '日付', '種別', '出欠', '配車区分', '乗車可能', '選手数', '同伴数'];
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#555').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const row = [
    now,
    data.family || '',
    data.sid || '',
    data.date || '',
    data.kind || '',      // '出欠' or '配車'
    data.attendance || '', // ◯●△×
    data.drive || '',      // ok/ng/no
    data.capacity || '',
    data.players || '',
    data.companions || ''
  ];
  sheet.appendRow(row);
  return { success: true };
}
