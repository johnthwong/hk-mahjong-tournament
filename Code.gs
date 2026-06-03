/* ==================================================
   1. ROUTER & INITIALIZATION
   ================================================== */

function doGet(e) {
  const portal = e.parameter.portal;
  let template, title;

  if (portal === 'admin') {
    template = HtmlService.createTemplateFromFile('admin');
    title = '🏆 Tournament Admin';
  } else {
    template = HtmlService.createTemplateFromFile('index');
    title = 'Player Portal';
  }

  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }

/* ==================================================
   2. DATA AGGREGATION & CACHING
   ================================================== */


// Global memory cache to prevent redundant reads
let _cachedDataSS = null;
let _settingsMap = null;
let _cachedSettings = null;
let _sheetDataCache = {};

function clearCache() {
  _sheetDataCache = {};
  _cachedSettings = null;
  _settingsMap = null;
}

function getCachedSheetData(sheetName) {
  if (_sheetDataCache[sheetName] !== undefined) return _sheetDataCache[sheetName];
  const ss = getDataSS();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) { 
      _sheetDataCache[sheetName] = null; 
      return null; 
  }
  _sheetDataCache[sheetName] = sheet.getDataRange().getValues();
  return _sheetDataCache[sheetName];
}

function getDataSS() {
  if (_cachedDataSS) return _cachedDataSS;
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = master.getSheetByName("Settings");
  if (!settingsSheet) return master; 
  
  const data = settingsSheet.getDataRange().getValues();
  let targetId = "";
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == "Active_Tournament_ID") { 
      targetId = data[i][1]; 
      break; 
    }
  }
  
  if (targetId) {
    try { 
      _cachedDataSS = SpreadsheetApp.openById(targetId);
      return _cachedDataSS;
    } catch (e) { return master; }
  }
  return master;
}

function getInitialAdminData() {
  const settings = getFullSettings();
  return {
    settings: settings,
    tournaments: getTournamentList(),
    rulesets: getUniqueRulesets(), 
    url: getSpreadsheetUrl(),
    players: getPlayers(),
    schedule: getScheduleTables(),
    pairingState: getPairingState(),
    penalties: getRecentPenalties(),
    penaltyList: getPenaltyDefinitions(), 
    scoreLog: getScoreLog(),
    allGames: getAllGamesData()
  };
}

function getScoringUpdateData() {
  
  return {
    scoreLog: getScoreLog(),
    allGames: getAllGamesData()
  };
}

function getUniqueRulesets() {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = master.getSheetByName("Penalties_List");
  if (!sheet) return ["Default"];
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ["Default"];

  const headers = data[0].map(h => String(h).toLowerCase());
  const ruleIdx = headers.indexOf("ruleset");
  if (ruleIdx === -1) return ["Default"];

  let rulesets = new Set();
  for(let i = 1; i < data.length; i++) {
    if(data[i][ruleIdx]) rulesets.add(data[i][ruleIdx].toString());
  }
  return Array.from(rulesets).sort();
}

function getPenaltyDefinitions() {
  const ss = getDataSS();
  const data = getCachedSheetData("Penalties_List");
  if (!data || data.length < 2) return [];
  
  const headers = data[0].map(h => String(h).toLowerCase());
  let typeIdx = headers.indexOf("type");
  let foulIdx = headers.indexOf("foul");
  let penIdx = headers.indexOf("penalty");
  let ptIdx = headers.indexOf("point deduction");
  let ruleIdx = headers.indexOf("ruleset");

  if (typeIdx === -1) typeIdx = 0;
  if (foulIdx === -1) foulIdx = 1;
  if (penIdx === -1) penIdx = 2;
  if (ptIdx === -1) ptIdx = 3;

  let activeRuleset = readSetting(ss, "Active_Ruleset", "");

  let defs = [];
  for(let i=1; i<data.length; i++) {
    if (ruleIdx > -1 && activeRuleset) {
        if (String(data[i][ruleIdx]) !== String(activeRuleset)) continue;
    }
    if(data[i][typeIdx]) {
        defs.push({
          type: data[i][typeIdx], foul: data[i][foulIdx], penalty: data[i][penIdx], pointDeduction: data[i][ptIdx]
        });
    }
  }
  return defs;
}
function getAllGamesData() {
  try {
    const pairData = getCachedSheetData("Pairings");
    const scoreData = getCachedSheetData("Scores"); 
    if(!pairData) return {};
    
    let scoredMap = {};
    if (scoreData && scoreData.length > 1) {
        for (let i = 1; i < scoreData.length; i++) {
            let r = String(scoreData[i][1]);
            let t = String(scoreData[i][2]);
            if (!scoredMap[r]) scoredMap[r] = new Set();
            scoredMap[r].add(t);
        }
    }

    const pMap = getPlayerMap();
    let gamesByRound = {};
    let currentRound = 0;

    for(let row of pairData) {
      if(!row[0]) continue; 
      let cell = row[0].toString().toUpperCase();
      if(cell.includes("ROUND")) {
        let match = cell.match(/\d+/);
        currentRound = match ? parseInt(match[0]) : 0;
        continue;
      }
      if(currentRound > 0) {
        let tableId = parseInt(row[0]);
        if (!isNaN(tableId)) {
          if (!gamesByRound[currentRound]) gamesByRound[currentRound] = [];
          const getP = (id) => ({ id: id, name: pMap[id] || id });
          let isScored = (scoredMap[String(currentRound)] && scoredMap[String(currentRound)].has(String(tableId)));
          gamesByRound[currentRound].push({ 
              id: tableId, 
              p1: row[1] ? getP(row[1]) : { id: "?", name: "?" }, 
              p2: row[2] ? getP(row[2]) : { id: "?", name: "?" }, 
              p3: row[3] ? getP(row[3]) : { id: "?", name: "?" }, 
              p4: row[4] ? getP(row[4]) : { id: "?", name: "?" },
              isScored: isScored
          });
        }
      }
    }
    return gamesByRound;
  } catch (e) { return {}; }
}

/* ==================================================
   3. DATABASE & SETTINGS
   ================================================== */

function getSpreadsheetUrl() { return getDataSS().getUrl(); }

function getTournamentList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parents = DriveApp.getFileById(ss.getId()).getParents();
    if (!parents.hasNext()) return [];
    const folder = parents.next();
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    let list = [];
    while (files.hasNext()) {
      let f = files.next();
      if (f.getId() !== ss.getId()) list.push({ name: f.getName(), id: f.getId() });
    }
    return list.sort((a,b) => a.name.localeCompare(b.name));
  } catch (e) { return []; }
}

function switchTournament(fileId) {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(fileId);
  updateSheetSetting(master, "Active_Tournament_ID", fileId);
  updateSheetSetting(master, "Active_Tournament_Name", file.getName());
  clearCache(); 
  _cachedDataSS = null;
  return "Switched to: " + file.getName();
}

function startNewTournament(tournamentName, rulesetName) {
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const cleanName = tournamentName || "New Tournament " + new Date().toLocaleDateString();
  const newSS = SpreadsheetApp.create(cleanName);
  const newId = newSS.getId();
  const masterFile = DriveApp.getFileById(master.getId());
  
  if (masterFile.getParents().hasNext()) {
    DriveApp.getFileById(newId).moveTo(masterFile.getParents().next());
  }

  try {
    // Generates the new 4-column layout for brand new files
    newSS.insertSheet("Players").appendRow(["Player ID", "Name", "Checked In", "ARA ID"]);
    newSS.insertSheet("Settings").appendRow(["Key", "Value"]);
    
    const pList = newSS.insertSheet("Penalties_List");
    pList.appendRow(["Type", "Foul", "Penalty", "Point Deduction"]);
    const masterPList = master.getSheetByName("Penalties_List");
    if (masterPList) {
        const data = masterPList.getDataRange().getValues();
        const h = data[0].map(x => String(x).toLowerCase());
        const rIdx = h.indexOf("ruleset");
        const tIdx = h.indexOf("type");
        const fIdx = h.indexOf("foul");
        const pIdx = h.indexOf("penalty");
        const ptIdx = h.indexOf("point deduction");
        
        if (rIdx > -1 && tIdx > -1 && fIdx > -1 && pIdx > -1) {
            let rowsToAdd = [];
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][rIdx]) === String(rulesetName)) {
                    let ptVal = (ptIdx > -1) ? data[i][ptIdx] : "0";
                    rowsToAdd.push([data[i][tIdx], data[i][fIdx], data[i][pIdx], ptVal]);
                }
            }
            if (rowsToAdd.length > 0) {
                pList.getRange(2, 1, rowsToAdd.length, 4).setValues(rowsToAdd);
            }
        }
    } else {
        pList.appendRow(["Major", "Example Penalty", "Chombo", "-20"]);
    }

    const sc = newSS.insertSheet("Scores");
    sc.appendRow(["Timestamp", "Round", "Game ID", "P1 ID", "Raw P1", "Formatted P1", "P2 ID", "Raw P2", "Formatted P2", "P3 ID", "Raw P3", "Formatted P3", "P4 ID", "Raw P4", "Formatted P4", "Leftover"]);
    const pen = newSS.insertSheet("Penalties");
    pen.appendRow(["Timestamp", "Player ID", "Points Deducted", "Reason", "Round", "Table", "Notes"]);
    
    newSS.insertSheet("Pairings");
    newSS.insertSheet("Leaderboard");
    const def = newSS.getSheetByName("Sheet1");
    if (def) newSS.deleteSheet(def);

    updateSheetSetting(master, "Active_Tournament_ID", newId);
    updateSheetSetting(master, "Active_Tournament_Name", cleanName);
    updateSheetSetting(newSS, "Tournament Name", cleanName);
    updateSheetSetting(newSS, "Rotation Seed", 0);
    updateSheetSetting(newSS, "Round Count", 4);
    updateSheetSetting(newSS, "Tiebreaker_Rule", "split");
    updateSheetSetting(newSS, "Faan_Min", 3);
    updateSheetSetting(newSS, "Faan_Max", 13);
    updateSheetSetting(newSS, "Faan_Min_Points", 8);
    updateSheetSetting(newSS, "Faan_Scaling", "half");
    updateSheetSetting(newSS, "Faan_Custom_Points", "");
    updateSheetSetting(newSS, "Self_Pick_Multiplier", 1.5);
    updateSheetSetting(newSS, "False_Win_Points", "");
    updateSheetSetting(newSS, "Pre_Tourney_Enabled", "false"); 
    updateSheetSetting(newSS, "Tourney_Begun", "false"); 
    updateSheetSetting(newSS, "Active_Ruleset", rulesetName); 

    return "Success: Created '" + cleanName + "'";
  } catch (e) { throw new Error("Setup failed: " + e.message); }
}

function updateSheetSetting(ss, key, val) {
  let sheet = ss.getSheetByName("Settings");
  if (!sheet) sheet = ss.insertSheet("Settings");
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == key) { 
      sheet.getRange(i + 1, 2).setValue(val);
      return; 
    }
  }
  sheet.appendRow([key, val]);
}

function getFullSettings() {
  if (_cachedSettings) return _cachedSettings;
  const master = SpreadsheetApp.getActiveSpreadsheet();
  const dataSS = getDataSS();
  
  const read = (ss, k, d) => {
    let sheet = ss.getSheetByName("Settings");
    if(!sheet) return d;
    const v = sheet.getDataRange().getValues();
    for(let i=0; i<v.length; i++) if(v[i][0] == k) return v[i][1] === "" ? d : v[i][1];
    return d;
  };

  _cachedSettings = {
    activeName: read(master, "Active_Tournament_Name", "No Active Tournament"),
    activeId: read(master, "Active_Tournament_ID", ""),
    roundCount: read(dataSS, "Round Count", 4),
    pairingMode: read(dataSS, "Pairing_Mode", "scramble"),
    topCutEnabled: read(dataSS, "Top_Cut_Enabled", "false"),
    topCutSize: read(dataSS, "Top_Cut_Size", 0),
    topCutRound: read(dataSS, "Top_Cut_Round", 0),
    tiebreakerRule: read(dataSS, "Tiebreaker_Rule", "split"),
    faanMin: read(dataSS, "Faan_Min", 3),
    faanMax: read(dataSS, "Faan_Max", 13),
    faanMinPoints: read(dataSS, "Faan_Min_Points", 8),
    faanScaling: read(dataSS, "Faan_Scaling", "half"),
    faanCustomPoints: read(dataSS, "Faan_Custom_Points", ""),
    selfPickMultiplier: read(dataSS, "Self_Pick_Multiplier", 1.5),
    falseWinPoints: read(dataSS, "False_Win_Points", ""),
    preTourneyEnabled: read(dataSS, "Pre_Tourney_Enabled", "false"),
    tourneyBegun: read(dataSS, "Tourney_Begun", "false"),
    activeRuleset: read(dataSS, "Active_Ruleset", ""),
    theme: read(dataSS, "Theme", "default")
  };
  return _cachedSettings;
}

// Base (unscaled) points for a faan count.
//   full-spicy: pure doubling, 2^faan.
//   half-spicy: same as full up to 4 faan, then doubles every 2 faan, with odd
//   faan = 1.5x the previous even faan (matches the standard HK faan-to-score table).
function faanBase(f, scaling) {
  if (scaling === 'full') return Math.pow(2, f);
  if (f <= 4) return Math.pow(2, f);
  if (f % 2 === 0) return Math.pow(2, f / 2 + 2);
  return 1.5 * Math.pow(2, (f - 1) / 2 + 2);
}

// Build the faan -> points table, scaled so the minimum faan is worth the
// configured point value. Self-pick points = points x self-pick multiplier
// (the winner's total on a self-draw, split among the three opponents).
function computeFaanTable(s) {
  const minF = parseInt(s.faanMin);
  const maxF = parseInt(s.faanMax);
  const minPts = Number(s.faanMinPoints);
  const scaling = String(s.faanScaling || 'half').toLowerCase();
  const spMult = Number(s.selfPickMultiplier) || 1.5;
  const rows = [];

  if (scaling === 'custom') {
    const custom = String(s.faanCustomPoints || '').split(/[\s,]+/).filter(x => x !== '').map(Number);
    for (let f = minF, i = 0; f <= maxF; f++, i++) {
      const pts = (i < custom.length && !isNaN(custom[i])) ? custom[i] : 0;
      rows.push({ faan: f, points: pts, selfPick: Math.round(pts * spMult) });
    }
    return rows;
  }

  const baseMin = faanBase(minF, scaling) || 1;
  for (let f = minF; f <= maxF; f++) {
    const pts = Math.round(minPts * faanBase(f, scaling) / baseMin);
    rows.push({ faan: f, points: pts, selfPick: Math.round(pts * spMult) });
  }
  return rows;
}

// Client-callable: the faan table for the active tournament.
function getFaanTable() {
  return computeFaanTable(getFullSettings());
}

function saveTournamentSettings(form) {
  const ss = getDataSS();
  updateSheetSetting(ss, "Round Count", form.roundCount);
  updateSheetSetting(ss, "Pairing_Mode", form.pairingMode);
  updateSheetSetting(ss, "Top_Cut_Enabled", form.topCutEnabled);
  updateSheetSetting(ss, "Top_Cut_Size", form.topCutSize);
  updateSheetSetting(ss, "Top_Cut_Round", form.topCutRound);
  updateSheetSetting(ss, "Tiebreaker_Rule", form.tiebreakerRule);
  // Clamp faan range: max caps at 13 (HK), min >= 1 and strictly below max.
  let fMax = Math.min(13, Math.max(2, parseInt(form.faanMax) || 13));
  let fMin = Math.max(1, parseInt(form.faanMin) || 3);
  if (fMin >= fMax) fMin = fMax - 1;
  updateSheetSetting(ss, "Faan_Min", fMin);
  updateSheetSetting(ss, "Faan_Max", fMax);
  updateSheetSetting(ss, "Faan_Min_Points", form.faanMinPoints);
  updateSheetSetting(ss, "Faan_Scaling", form.faanScaling);
  updateSheetSetting(ss, "Faan_Custom_Points", form.faanCustomPoints);
  updateSheetSetting(ss, "Self_Pick_Multiplier", form.selfPickMultiplier);
  updateSheetSetting(ss, "False_Win_Points", form.falseWinPoints);
  updateSheetSetting(ss, "Pre_Tourney_Enabled", form.preTourneyEnabled);
  updateSheetSetting(ss, "Theme", form.theme);

  if (form.rulesetName) {
      updateSheetSetting(ss, "Active_Ruleset", form.rulesetName);
      const master = SpreadsheetApp.getActiveSpreadsheet();
      
      // If we are in a sub-tournament, overwrite the local Penalties_List with the newly chosen ruleset
      if (ss.getId() !== master.getId()) {
          const masterPList = master.getSheetByName("Penalties_List");
          let localPList = ss.getSheetByName("Penalties_List");
          if (masterPList && localPList) {
              const data = masterPList.getDataRange().getValues();
              const h = data[0].map(x => String(x).toLowerCase());
              const rIdx = h.indexOf("ruleset");
              const tIdx = h.indexOf("type");
              const fIdx = h.indexOf("foul");
              const pIdx = h.indexOf("penalty");
              const ptIdx = h.indexOf("point deduction");

              if (rIdx > -1 && tIdx > -1 && fIdx > -1 && pIdx > -1) {
                  let rowsToAdd = [["Type", "Foul", "Penalty", "Point Deduction"]];
                  for (let i = 1; i < data.length; i++) {
                      if (String(data[i][rIdx]) === String(form.rulesetName)) {
                          let ptVal = (ptIdx > -1) ? data[i][ptIdx] : "0";
                          rowsToAdd.push([data[i][tIdx], data[i][fIdx], data[i][pIdx], ptVal]);
                      }
                  }
                  if (rowsToAdd.length > 1) { 
                      localPList.clear();
                      localPList.getRange(1, 1, rowsToAdd.length, 4).setValues(rowsToAdd);
                  }
              }
          }
      }
  }
  clearCache();
  return "Settings Saved.";
}

function readSetting(ss, key, def) {
  if (!_settingsMap) {
    _settingsMap = new Map();
    let sheet = ss.getSheetByName("Settings");
    if(sheet) {
      const data = sheet.getDataRange().getValues();
      for(let i=0; i<data.length; i++) {
        _settingsMap.set(data[i][0], data[i][1]);
      }
    }
  }
  let val = _settingsMap.get(key);
  return (val !== undefined && val !== "") ? val : def;
}

/* ==================================================
   4. PLAYER MANAGEMENT
   ================================================== */

function getNextSafeId(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const match = data[i][0].toString().match(/\d+/);
    if (match) { 
        let n = parseInt(match[0], 10); 
        if (n > max) max = n;
    }
  }
  return max + 1;
}

function addPlayer(name, manualId, araId) {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Players");
  if (!sheet) { sheet = ss.insertSheet("Players"); sheet.appendRow(["Player ID", "Name", "Checked In", "ARA ID"]); }
  if (sheet.getRange(1, 4).getValue() !== "ARA ID") sheet.getRange(1, 4).setValue("ARA ID");

  let id = manualId || "P" + getNextSafeId(sheet);
  sheet.appendRow([id, name, "", araId || ""]);
  clearCache(); 
  return getPlayers();
}

function addPlayersBulk(names) {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Players");
  if (!sheet) { sheet = ss.insertSheet("Players"); sheet.appendRow(["Player ID", "Name", "Checked In", "ARA ID"]); }
  if (sheet.getRange(1, 4).getValue() !== "ARA ID") sheet.getRange(1, 4).setValue("ARA ID");

  let nextNum = getNextSafeId(sheet);
  const rows = [];
  names.forEach(line => { 
      if(line.trim()) { 
          let parts = line.split(' - ');
          let n = parts[0].trim();
          let ara = parts.length > 1 ? parts.slice(1).join(' - ').trim() : "";
          rows.push(["P" + nextNum, n, "", ara]); 
          nextNum++; 
      } 
  });
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  clearCache();
  return getPlayers();
}

function updateAraId(playerId, newAraId) {
    const ss = getDataSS();
    const sheet = ss.getSheetByName("Players");
    if (!sheet) return getPlayers();
    if (sheet.getRange(1, 4).getValue() !== "ARA ID") sheet.getRange(1, 4).setValue("ARA ID");

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === playerId.toString()) {
            sheet.getRange(i + 1, 4).setValue(newAraId);
            break;
        }
    }
    clearCache();
    return getPlayers();
}
function deletePlayer(playerId) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet) return getPlayers();
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0].toString() == playerId.toString()) sheet.deleteRow(i + 1);
  }
  clearCache();
  return getPlayers();
}

function togglePlayerDNF(playerId) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet) return getPlayers();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() == playerId.toString()) {
      let cur = data[i][1].toString();
      let neu = cur.startsWith("[DNF] ") ? cur.replace("[DNF] ", "") : "[DNF] " + cur;
      sheet.getRange(i + 1, 2).setValue(neu);
      break;
    }
  }
  clearCache();
  return getPlayers();
}

function getPlayers() {
  let data = getCachedSheetData("Players");
  if (!data) return [];
  
  if (data.length > 0 && data[0][3] !== "ARA ID") {
      const ss = getDataSS();
      const sheet = ss.getSheetByName("Players");
      if (sheet) sheet.getRange(1, 4).setValue("ARA ID");
      clearCache(); // Data structurally changed, wipe cache
      data = getCachedSheetData("Players");
  }
  
  if (data.length <= 1) return [];
  return data.slice(1).map(r => ({ 
      id: r[0], name: r[1],
      isCheckedIn: r[2] === true || r[2] === "true" || r[2] === "TRUE",
      araId: r[3] || ""
  }));
}

function clearAllPlayers() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  clearCache();
  return [];
}

function getPlayerMap() {
  const list = getPlayers();
  let map = {};
  list.forEach(p => map[p.id.toString()] = p.name);
  return map;
}
function renamePlayer(playerId, newName) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet) return;
  const cleanName = newName.trim();
  if (!cleanName) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString() === playerId.toString()) {
      let cur = data[i][1].toString();
      let isDNF = cur.startsWith("[DNF] ");
      let finalName = isDNF ? "[DNF] " + cleanName : cleanName;
      sheet.getRange(i + 1, 2).setValue(finalName);
      break;
    }
  }
  clearCache();
  updateLeaderboardSheet();
}
/* ==================================================
   5. PAIRING LOGIC
   ================================================== */

// A substitute is any player whose ID or name starts with "SUB" (auto-added subs
// use "SUB<n>" IDs / "SUBSTITUTE <n>" names). Subs are filler that only fill out
// the last table — they are not permanent participants.
function isSubPlayer(p) {
  return /^SUB/i.test(String(p.id)) || /^SUB/i.test(String(p.name));
}

function getPairingState() {
  const data = getCachedSheetData("Pairings");
  const sett = getFullSettings();
  const players = getPlayers();
  // Pairing only ever uses non-DNF players (see generateNextRound), so the bucket
  // preview / sub padding must count active players, not the full roster.
  const activePlayers = players.filter(p => !String(p.name).toUpperCase().startsWith("[DNF]"));
  if (activePlayers.length < 4) return { error: "Need at least 4 players." };

  let maxRound = 0;
  if (data && data.length > 1) {
    for(let row of data) {
      if(!row[0]) continue;
      let cell = row[0].toString().toUpperCase();
      if (cell.includes("ROUND")) {
        let match = cell.match(/\d+/);
        if(match && parseInt(match[0]) > maxRound) maxRound = parseInt(match[0]);
      }
    }
  }

  const ss = getDataSS();
  const cutEnabled = String(sett.topCutEnabled).toLowerCase() === "true";
  const cutSize = parseInt(sett.topCutSize) || 0;
  const cutRound = parseInt(sett.topCutRound) || 0;
  let isCutActive = String(readSetting(ss, "Top_Cut_Active", "false")).toLowerCase() === "true";
  
  // STRICT CHECK: Only trigger if we have a valid cutRound AND the next round has surpassed it
  let shouldTrigger = (cutEnabled && cutSize > 0 && !isCutActive && cutRound > 0 && (maxRound + 1) > cutRound);

  let validOptions = [];
  // Subs only fill out the last table, so the effective pool is the REAL (non-sub)
  // active players rounded up to a multiple of 4. Excess subs are benched at
  // generation, so they must not inflate the bucket count.
  let realActiveCount = activePlayers.filter(p => !isSubPlayer(p)).length;
  let pCount = Math.ceil(realActiveCount / 4) * 4;

  // Remove the Top Cut size from the bucket math if it is active/triggering
  let poolSize = (isCutActive || shouldTrigger) ? pCount - cutSize : pCount;
  if (poolSize < 4) poolSize = pCount; // Fallback if math gets weird
  
  let totalTables = Math.floor(poolSize / 4);

  for (let b = 1; b <= totalTables; b++) {
    let baseTables = Math.floor(totalTables / b);
    if (baseTables === 0) break; // Bucket would be empty, stop adding options

    let leftoverTables = totalTables % b;
    let label = "";
    
    // Label prefix based on whether Top Cut is active
    if (isCutActive || shouldTrigger) {
         label = (b === 1) ? `1 Gen Bucket (${poolSize})` : `${b} Gen Buckets`;
    } else {
         label = (b === 1) ? `1 Bucket (${poolSize})` : `${b} Buckets`;
    }

    // Add size summary to the label
    if (b > 1) {
        if (leftoverTables === 0) {
            label += ` (${baseTables * 4} players each)`;
        } else {
            let maxB = (baseTables + 1) * 4;
            let minB = baseTables * 4;
            label += ` (Split of ${maxB} & ${minB})`;
        }
    }
    
    validOptions.push({ val: b, label: label });
  }

  return {
    nextRound: maxRound + 1, totalRounds: sett.roundCount, playerCount: activePlayers.length,
    validBuckets: validOptions, lastBuckets: Number(readSetting(ss, "Last_Bucket_Count", 1))
  };
}

function generateNextRound(bucketCount, addSubs) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { success: false, message: "Server busy: Another admin is currently generating a round. Please try again." }; }

  try {
    const ss = getDataSS();
    let pairSheet = ss.getSheetByName("Pairings");
    if (!pairSheet) pairSheet = ss.insertSheet("Pairings");
    const state = getPairingState();
    const round = state.nextRound;

    const allGames = getAllGamesData();
    const currentRound = round - 1;
    if (currentRound > 0 && allGames[currentRound]) {
      const unscored = allGames[currentRound].filter(g => !g.isScored);
      if (unscored.length > 0) {
        return { success: false, message: `Cannot generate round.\n${unscored.length} table(s) are missing scores in Round ${currentRound}.` };
      }
    }

    let allPlayers = getPlayers();
    let active = allPlayers.filter(p => !p.name.toUpperCase().startsWith("[DNF]"));
    let reals = active.filter(p => !isSubPlayer(p));
    let subs = active.filter(p => isSubPlayer(p));
    // Subs are a filler pool: only enough to round real players up to a multiple
    // of 4. Excess subs are benched (not matched); add new ones only if we're
    // short and the caller opted in.
    let subsNeeded = (4 - (reals.length % 4)) % 4;

    if (addSubs && subs.length < subsNeeded) {
      // Substitutes get "SUB<n>" IDs (so SUB-aware logic can spot them by ID) and
      // "SUBSTITUTE <n>" names. Number off existing SUB IDs to avoid collisions.
      let subNums = allPlayers.map(p => /^SUB(\d+)$/i.exec(p.id)).filter(Boolean).map(m => parseInt(m[1], 10));
      let nextSub = subNums.length ? Math.max(...subNums) + 1 : 1;
      for (let i = subs.length; i < subsNeeded; i++) {
        addPlayer(`SUBSTITUTE ${nextSub}`, `SUB${nextSub}`);
        nextSub++;
      }
      allPlayers = getPlayers();
      active = allPlayers.filter(p => !p.name.toUpperCase().startsWith("[DNF]"));
      reals = active.filter(p => !isSubPlayer(p));
      subs = active.filter(p => isSubPlayer(p));
    }

    // Match all real players plus only the subs needed; bench any extras.
    let players = reals.concat(subs.slice(0, subsNeeded));

    if (players.length % 4 !== 0) {
      return { success: false, message: "Cannot generate round. The number of active players must be a multiple of 4." };
    }

    let historyMap = new Map();
    players.forEach(p => historyMap.set(p.id, new Set()));
    if (pairSheet.getLastRow() > 1) {
      const data = pairSheet.getDataRange().getValues();
      let inData = false;
      for(let row of data) {
        if(!row[0]) continue;
        if(row[0].toString().includes("ROUND")) { inData = true; continue; }
        if(inData && row[1]) {
          let pIds = [row[1], row[2], row[3], row[4]];
          for(let i=0; i<4; i++) {
            for(let j=i+1; j<4; j++) {
              if (historyMap.has(pIds[i])) historyMap.get(pIds[i]).add(pIds[j]);
              if (historyMap.has(pIds[j])) historyMap.get(pIds[j]).add(pIds[i]);
            }
          }
        }
      }
    }

    const mode = String(readSetting(ss, "Pairing_Mode", "scramble")).toLowerCase();
    const cutEnabled = String(readSetting(ss, "Top_Cut_Enabled", "false")).toLowerCase() === "true";
    const cutSize = parseInt(readSetting(ss, "Top_Cut_Size", 0));
    const cutRound = parseInt(readSetting(ss, "Top_Cut_Round", 0));
    let isCutActive = String(readSetting(ss, "Top_Cut_Active", "false")).toLowerCase() === "true";
    let savedCutIDs = readSetting(ss, "Top_Cut_Player_IDs", "").split(",").filter(x => x);
    let buckets = [];
    
    let shouldTrigger = (cutEnabled && cutSize > 0 && !isCutActive && cutRound > 0 && round > cutRound);

    if (shouldTrigger) {
        isCutActive = true;
        const standings = getStandingsData();
        let ranked = players.map(p => {
            let s = standings.find(x => x.id === p.id);
            return { ...p, pts: s ? s.totalScore : -9999 };
        });
        ranked.sort((a,b) => b.pts - a.pts);

        let topPool = ranked.slice(0, cutSize);
        let restPool = ranked.slice(cutSize);
        
        updateSheetSetting(ss, "Top_Cut_Start_Round", round);
        updateSheetSetting(ss, "Top_Cut_Player_IDs", topPool.map(p=>p.id).join(","));
        updateSheetSetting(ss, "Top_Cut_Active", "true");
        
        if (mode === 'swiss') {
            buckets.push(topPool.sort(() => Math.random() - 0.5));
            let remBuckets = Math.max(1, parseInt(bucketCount, 10));
            buckets.push(...sliceIntoSwissBuckets(restPool, remBuckets));
        } else {
            buckets.push(topPool.sort(() => Math.random() - 0.5));
            buckets.push(restPool.sort(() => Math.random() - 0.5));
        }
    }
    else if (isCutActive && savedCutIDs.length > 0) {
        let topPool = players.filter(p => savedCutIDs.includes(p.id));
        let restPool = players.filter(p => !savedCutIDs.includes(p.id));
        
        if (mode === 'swiss') {
            const standings = getStandingsData();
            const getStats = (pid) => standings.find(x => x.id === pid);
            buckets.push(topPool.sort(() => Math.random() - 0.5)); 
            
            restPool.sort((a,b) => {
                let sa = getStats(a.id); let sb = getStats(b.id);
                return (sb ? sb.totalScore : -9999) - (sa ? sa.totalScore : -9999);
            });
            let remBuckets = Math.max(1, parseInt(bucketCount, 10));
            buckets.push(...sliceIntoSwissBuckets(restPool, remBuckets));
        } else {
            buckets.push(topPool.sort(() => Math.random() - 0.5));
            buckets.push(restPool.sort(() => Math.random() - 0.5));
        }
    }
    else {
        if (mode === 'swiss') {
            const standings = getStandingsData();
            let ranked = players.map(p => {
                let s = standings.find(x => x.id === p.id);
                return { ...p, pts: s ? s.totalScore : -9999 };
            });
            ranked.sort((a,b) => b.pts - a.pts);
            buckets.push(...sliceIntoSwissBuckets(ranked, bucketCount));
        }
        else {
            buckets.push(players.sort(() => Math.random() - 0.5));
        }
    }
    
    updateSheetSetting(ss, "Last_Bucket_Count", bucketCount);

    let roundTables = [];
    let tableCounter = 1;
    let totalRepeats = 0;
    let conflictDetails = [];

    buckets.forEach((bucket, bIdx) => {
      let pool = [...bucket];
      let bucketChar = String.fromCharCode(65 + bIdx); 
      let allowRepeats = (isCutActive && bIdx === 0);

      for (let i = pool.length - 1; i > 0; i--) {
          const j = Math.floor(Math.random() * (i + 1));
          [pool[i], pool[j]] = [pool[j], pool[i]];
      }

      let bucketTables = [];
      for (let i = 0; i < pool.length; i += 4) {
          bucketTables.push(pool.slice(i, i + 4));
      }

      if (!allowRepeats && bucketTables.length > 1) {
          const getBucketRepeats = (tables) => {
              let total = 0;
              for (let t of tables) total += countRepeats(t, historyMap);
              return total;
          };

          let currentRepeats = getBucketRepeats(bucketTables);
          let iterations = 0;
          const maxIterations = 5000;

          while (currentRepeats > 0 && iterations < maxIterations) {
              let conflictTableIndices = [];
              for (let i = 0; i < bucketTables.length; i++) {
                  if (countRepeats(bucketTables[i], historyMap) > 0) {
                      conflictTableIndices.push(i);
                  }
              }

              if (conflictTableIndices.length === 0) break;

              let t1Idx = conflictTableIndices[Math.floor(Math.random() * conflictTableIndices.length)];
              let t2Idx = Math.floor(Math.random() * bucketTables.length);

              while (t1Idx === t2Idx) {
                  t2Idx = Math.floor(Math.random() * bucketTables.length);
              }

              let p1Idx = Math.floor(Math.random() * 4);
              let p2Idx = Math.floor(Math.random() * 4);

              let temp = bucketTables[t1Idx][p1Idx];
              bucketTables[t1Idx][p1Idx] = bucketTables[t2Idx][p2Idx];
              bucketTables[t2Idx][p2Idx] = temp;

              let newRepeats = getBucketRepeats(bucketTables);

              if (newRepeats <= currentRepeats) {
                  currentRepeats = newRepeats;
              } else {
                  let tempBack = bucketTables[t1Idx][p1Idx];
                  bucketTables[t1Idx][p1Idx] = bucketTables[t2Idx][p2Idx];
                  bucketTables[t2Idx][p2Idx] = tempBack;
              }
              iterations++;
          }
      }

      bucketTables.forEach(tPlayers => {
          let tRepeats = allowRepeats ? 0 : countRepeats(tPlayers, historyMap);
          totalRepeats += tRepeats;
          let currentTableId = tableCounter++;
          roundTables.push([currentTableId, tPlayers[0].id, tPlayers[1].id, tPlayers[2].id, tPlayers[3].id, bucketChar]);

          if (tRepeats > 0 && !allowRepeats) {
              let tableConflicts = [];
              for(let i=0; i<4; i++) {
                  for(let j=i+1; j<4; j++) {
                      let p1 = tPlayers[i];
                      let p2 = tPlayers[j];
                      if (historyMap.has(p1.id) && historyMap.get(p1.id).has(p2.id)) {
                          tableConflicts.push(`${p1.name} & ${p2.name}`);
                      }
                  }
              }
              if (tableConflicts.length > 0) {
                  conflictDetails.push(`Table ${currentTableId}: ${tableConflicts.join(' | ')}`);
              }
          }
      });
    });

    let output = [[`--- ROUND ${round} (${mode.toUpperCase()}) ---`, "", "", "", "", ""]];
    roundTables.forEach(row => output.push(row));
    output.push(["", "", "", "", "", ""]); 
    pairSheet.getRange(pairSheet.getLastRow() + 1, 1, output.length, 6).setValues(output);
    
    clearCache();
    let msg = `Generated Round ${round} pairings!`;
    return { success: true, message: msg, repeats: totalRepeats, conflicts: conflictDetails };
  } catch(e) {
    return { success: false, message: "Error generating round: " + e.message };
  } finally {
    lock.releaseLock();
  }
}
function deleteLastRound(roundNum) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return false; }
  try {
    const ss = getDataSS();
    const sheet = ss.getSheetByName("Pairings");
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    let startRow = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] && data[i][0].toString().toUpperCase().includes(`--- ROUND ${roundNum} (`)) {
        startRow = i;
        break;
      }
    }
    if (startRow !== -1) {
      sheet.getRange(startRow + 1, 1, sheet.getLastRow() - startRow, 6).clearContent();
    }
    clearCache();
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Delete a round's pairings AND its score rows. Used to regenerate the latest
// round from scratch (e.g. to drop subs that are no longer needed after DNFs).
function deleteRoundAndScores(roundNum) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return false; }
  try {
    const ss = getDataSS();

    const pSheet = ss.getSheetByName("Pairings");
    if (pSheet) {
      const data = pSheet.getDataRange().getValues();
      let startRow = -1;
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][0] && data[i][0].toString().toUpperCase().includes(`--- ROUND ${roundNum} (`)) {
          startRow = i;
          break;
        }
      }
      if (startRow !== -1) {
        pSheet.getRange(startRow + 1, 1, pSheet.getLastRow() - startRow, 6).clearContent();
      }
    }

    // Remove score rows whose Round column (col 2) matches.
    const sSheet = ss.getSheetByName("Scores");
    if (sSheet && sSheet.getLastRow() > 1) {
      const sData = sSheet.getDataRange().getValues();
      for (let i = sData.length - 1; i >= 1; i--) {
        if (String(sData[i][1]) === String(roundNum)) sSheet.deleteRow(i + 1);
      }
    }

    clearCache();
    return true;
  } finally {
    lock.releaseLock();
  }
}

function countRepeats(players, historyMap) {
  let repeats = 0;
  for(let i=0; i<players.length; i++) {
    let pid = players[i].id;
    if(pid === "BYE" || pid.startsWith("SUB")) continue;
    let past = historyMap.get(pid);
    if (!past) continue;
    
    for(let j=i+1; j<players.length; j++) {
      let pid2 = players[j].id;
      if(pid2 === "BYE" || pid2.startsWith("SUB")) continue;
      if(past.has(pid2)) repeats++;
    }
  }
  return repeats;
}

function getScheduleTables() {
  const data = getCachedSheetData("Pairings");
  const pMap = getPlayerMap();
  if (!data) return [];
  
  let schedule = [];
  let currentRoundObj = null;
  for (let row of data) {
    if (!row[0]) continue;
    let cell = row[0].toString().toUpperCase();
    if (cell.includes("ROUND")) {
      let match = cell.match(/\d+/);
      if (match) {
        currentRoundObj = { round: parseInt(match[0]), tables: [] };
        schedule.push(currentRoundObj);
      }
      continue;
    }
    if (currentRoundObj && !isNaN(parseInt(row[0]))) {
      let table = {
        id: parseInt(row[0]), bucket: (row.length > 5 && row[5]) ? row[5] : "",
        p1: pMap[row[1]] || row[1], p2: pMap[row[2]] || row[2],
        p3: pMap[row[3]] || row[3], p4: pMap[row[4]] || row[4]
      };
      currentRoundObj.tables.push(table);
    }
  }
  return schedule.reverse();
}
// ADD THIS HELPER FUNCTION:
function sliceIntoSwissBuckets(pool, bucketCount) {
    let buckets = [];
    let bCount = Math.max(1, parseInt(bucketCount, 10));
    let totalTables = Math.floor(pool.length / 4);
    let baseTables = Math.floor(totalTables / bCount);
    let leftoverTables = totalTables % bCount;
    
    // Evenly distribute leftover tables Top/Bottom
    let extraTables = new Array(bCount).fill(0);
    for (let k = 0; k < leftoverTables; k++) {
        if (k % 2 === 0) extraTables[k / 2]++;
        else extraTables[bCount - 1 - Math.floor(k / 2)]++;
    }
    
    let currentIdx = 0;
    for (let i = 0; i < bCount; i++) {
        let size = (baseTables + extraTables[i]) * 4;
        if (size > 0) {
            buckets.push(pool.slice(currentIdx, currentIdx + size).sort(() => Math.random() - 0.5));
            currentIdx += size;
        }
    }
    return buckets;
}
/* ==================================================
   6. SCORING & PENALTIES
   ================================================== */

function checkIfScored(round, gameId, p1Id, p2Id, p3Id, p4Id) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Scores");
  if (!sheet) return { scored: false };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == round && data[i][2] == gameId) {
      const rowP1 = data[i][3]; const rowP2 = data[i][6]; const rowP3 = data[i][9]; const rowP4 = data[i][12];
      if (p1Id && (rowP1 != p1Id || rowP2 != p2Id || rowP3 != p3Id || rowP4 != p4Id)) {
        sheet.deleteRow(i + 1);
        return { scored: false, mismatchDeleted: true };
      }
      return { scored: true, rowIndex: i + 1, scores: { p1: data[i][4], p2: data[i][7], p3: data[i][10], p4: data[i][13], leftover: data[i][15] } };
    }
  }
  return { scored: false };
}

function saveScores(form) {
  const lock = LockService.getScriptLock();
  try { 
    lock.waitLock(30000); // Wait up to 30 seconds for the lock to clear
  } catch (e) { 
    return { success: false, message: "Server busy: Another admin is currently saving. Please try again." }; 
  }

  try {
    const ss = getDataSS();
    let sheet = ss.getSheetByName("Scores");
    if (!sheet) sheet = ss.insertSheet("Scores");
    // HK net-points model: each player's entered score IS their net result for
    // the game (zero-sum across the table). No Uma, no starting stack, no
    // thousands scaling. The last Scores column (legacy "Leftover") stays 0.
    const g = form.game;
    const res = {
      p1: { raw: Number(g.p1Score), final: Number(g.p1Score) },
      p2: { raw: Number(g.p2Score), final: Number(g.p2Score) },
      p3: { raw: Number(g.p3Score), final: Number(g.p3Score) },
      p4: { raw: Number(g.p4Score), final: Number(g.p4Score) }
    };

    const check = checkIfScored(form.round, g.gameId);
    const rowData = [ new Date(), form.round, g.gameId, g.p1Id, res.p1.raw, res.p1.final, g.p2Id, res.p2.raw, res.p2.final, g.p3Id, res.p3.raw, res.p3.final, g.p4Id, res.p4.raw, res.p4.final, 0 ];
    if (check.scored && check.rowIndex) { sheet.getRange(check.rowIndex, 1, 1, rowData.length).setValues([rowData]); }
    else { sheet.appendRow(rowData); }
    
   
    clearCache(); 
    updateLeaderboardSheet();
    return { success: true, message: check.scored ? "Updated existing score." : "Saved new score." };
  } catch (e) {
    return { success: false, message: "Error saving score: " + e.message };
  } finally {
    lock.releaseLock();
  }
}


function addPenalty(round, table, playerId, points, reason, notes) {
  const lock = LockService.getScriptLock();
  try { 
    lock.waitLock(30000); 
  } catch (e) { 
    return { success: false, message: "Server busy: Another admin is currently saving. Please try again." }; 
  }

  try {
    const ss = getDataSS();
    let sheet = ss.getSheetByName("Penalties");
    if (!sheet) { sheet = ss.insertSheet("Penalties"); sheet.appendRow(["Timestamp", "Player ID", "Points Deducted", "Reason", "Round", "Table", "Notes"]); }
    
    let pts = Number(points);
    if (isNaN(pts)) pts = 0;
    sheet.appendRow([new Date(), playerId, pts, reason, round, table, notes]);
    
   
    clearCache(); 
    updateLeaderboardSheet();
    return { success: true, message: "Penalty Added." };
  } catch (e) {
    return { success: false, message: "Error adding penalty: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// False Win: the offender loses 3*V and each of the other three players at the
// table gains V (zero-sum). V defaults to the points at the maximum faan. Stored
// as penalty rows (positive "Points Deducted" reduces score; negative credits it).
function applyFalseWin(round, table, offenderId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { success: false, message: "Server busy. Please try again." }; }
  try {
    const ss = getDataSS();
    const s = getFullSettings();

    let V = Number(s.falseWinPoints);
    if (!V || isNaN(V) || V <= 0) {
      const tbl = computeFaanTable(s);
      V = tbl.length ? tbl[tbl.length - 1].points : 0;
    }
    if (!V) return { success: false, message: "False Win value is 0 — set it in Settings (or configure the Faan table) first." };

    const games = getAllGamesData()[round] || [];
    const game = games.find(x => String(x.id) === String(table));
    if (!game) return { success: false, message: `Table ${table} not found in Round ${round}.` };
    const seated = ['p1', 'p2', 'p3', 'p4'].map(k => String(game[k].id)).filter(id => id && id !== "?");
    if (seated.indexOf(String(offenderId)) === -1) return { success: false, message: "That player is not seated at the selected table." };

    let pSheet = ss.getSheetByName("Penalties");
    if (!pSheet) { pSheet = ss.insertSheet("Penalties"); pSheet.appendRow(["Timestamp", "Player ID", "Points Deducted", "Reason", "Round", "Table", "Notes"]); }

    const now = new Date();
    const others = seated.filter(id => id !== String(offenderId) && id !== "BYE");
    const rows = [[now, offenderId, 3 * V, "False Win", round, table, `Pays ${V} to each of ${others.length} opponents`]];
    others.forEach(id => rows.push([now, id, -V, "False Win (credit)", round, table, `Credit from ${offenderId}'s false win`]));
    pSheet.getRange(pSheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);

    clearCache();
    updateLeaderboardSheet();
    return { success: true, message: `False Win applied: ${offenderId} −${3 * V}; each of ${others.length} opponents +${V}.` };
  } catch (e) {
    return { success: false, message: "Error applying False Win: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

function getRecentPenalties() {
  const data = getCachedSheetData("Penalties");
  if (!data || data.length <= 1) return [];
  const pMap = getPlayerMap();
  return data.slice(1).reverse().map(r => {
    let dateStr = "N/A";
    try { if (r[0]) dateStr = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "MM/dd HH:mm"); } catch(e) { dateStr = r[0].toString(); }
    return { 
        date: dateStr, rawDate: r[0] ? new Date(r[0]).getTime() : 0,
        playerId: r[1], name: pMap[r[1]] || r[1], 
        points: r[2], reason: r[3], round: (r[4] || "-"), table: (r[5] || "-"), notes: (r[6] || "") 
    };
  });
}

function deletePenalty(rawDate, playerId) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Penalties");
  if (!sheet) return { success: false, message: "Penalties sheet not found." };
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    let rDate = data[i][0] ? new Date(data[i][0]).getTime() : 0;
    let rPid = data[i][1];
    if (rDate === rawDate && String(rPid) === String(playerId)) {
      sheet.deleteRow(i + 1);
      clearCache(); // New: Wipe memory before recalculating
      updateLeaderboardSheet();
      return { success: true, message: "Penalty deleted successfully." };
    }
  }
  return { success: false, message: "Penalty record not found." };
}

function getScoreLog() {
  const data = getCachedSheetData("Scores");
  if (!data || data.length <= 1) return [];
  const pMap = getPlayerMap();
  return data.slice(1).reverse().map(r => {
    let dateStr = "N/A";
    try { if (r[0]) dateStr = Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "MM/dd HH:mm"); } catch(e) { dateStr = r[0].toString(); }
    return { date: dateStr, round: r[1], game: r[2], p1: `${pMap[r[3]] || r[3]} (${r[4]})`, p2: `${pMap[r[6]] || r[6]} (${r[7]})`, p3: `${pMap[r[9]] || r[9]} (${r[10]})`, p4: `${pMap[r[12]] || r[12]} (${r[13]})`, leftover: r[15] };
  });
}

/* ==================================================
   7. LEADERBOARD
   ================================================== */

function getStandingsData() {
  const ss = getDataSS();
  const sData = getCachedSheetData("Scores");
  const pData = getCachedSheetData("Penalties");
  const playersList = getPlayers(); // Use full list to get ARA ID
  
  let startRound = parseInt(readSetting(ss, "Top_Cut_Start_Round", "0"));
  let topIDsRaw = String(readSetting(ss, "Top_Cut_Player_IDs", ""));
  let topIDs = topIDsRaw ? topIDsRaw.split(",").filter(x => x) : [];
  let topSet = new Set(topIDs);
  let hasCut = (startRound > 0 && topIDs.length > 0);

  let stats = {};
  playersList.forEach(p => { 
      stats[p.id] = { 
          id: p.id, name: p.name, araId: p.araId || "", totalPts: 0, postCutPts: 0, preCutPts: 0,
          played: 0, pen: 0, isDNF: p.name.startsWith("[DNF]"), isTopCut: topSet.has(p.id)
      }; 
  });

  if(sData && sData.length > 1) {
    for(let i=1; i<sData.length; i++) {
      if (!sData[i][1]) continue;
      let rNum = parseInt(sData[i][1]);
      [3, 6, 9, 12].forEach((idIdx) => {
        const pid = sData[i][idIdx]; 
        const pts = Number(sData[i][idIdx + 2]); 
        if(pid && stats[pid]) { 
            stats[pid].totalPts += pts;
            stats[pid].played++;
            if (hasCut && rNum >= startRound) stats[pid].postCutPts += pts;
        }
      });
    }
  }
  
  if(pData && pData.length > 1) {
    for(let i=1; i<pData.length; i++) {
      if (!pData[i][1]) continue;
      let rNum = parseInt(pData[i][4]);
      const pid = pData[i][1]; 
      const deductFmt = Number(pData[i][2]);
      if(pid && stats[pid]) { 
          stats[pid].totalPts -= deductFmt; 
          stats[pid].pen += deductFmt;
          if (hasCut && rNum >= startRound) stats[pid].postCutPts -= deductFmt;
      }
    }
  }

  Object.values(stats).forEach(p => {
      p.preCutPts = p.totalPts - p.postCutPts;
      p.sortScore = p.isTopCut ? p.postCutPts : p.totalPts; 
  });

  let topGroup = Object.values(stats).filter(p => p.isTopCut).sort((a,b) => b.sortScore - a.sortScore);
  let restGroup = Object.values(stats).filter(p => !p.isTopCut).sort((a,b) => {
      if (a.isDNF !== b.isDNF) return a.isDNF ? 1 : -1;
      return b.sortScore - a.sortScore;
  });

  const formatP = (p, rank) => ({
      rank: p.isDNF ? "-" : rank, id: p.id, name: p.name, araId: p.araId,
      displayScore: p.totalPts, auxScore: p.postCutPts, totalScore: p.totalPts,
      played: p.played, penalties: p.pen, isDNF: p.isDNF, isTopCut: p.isTopCut
  });

  return [ ...topGroup.map((p, i) => formatP(p, i+1)), ...restGroup.map((p, i) => formatP(p, topGroup.length + i + 1)) ];
}

function updateLeaderboardSheet() {
  const ss = getDataSS();
  let sheet = ss.getSheetByName("Leaderboard");
  if (!sheet) sheet = ss.insertSheet("Leaderboard");
  const standings = getStandingsData();
  
  // Added p.araId to the row output
  const rows = standings.map(p => [ p.rank, p.id, p.name, p.araId, p.played, p.displayScore ]);
  
  sheet.clear();
  // Added "ARA ID" to the headers
  const output = [["Rank", "Player ID", "Name", "ARA ID", "Games Played", "Total Points"], ...rows];
  // Updated range width to 6 columns
  sheet.getRange(1, 1, output.length, 6).setValues(output);
}
/* ==================================================
   8. SWAP & EDITING TOOLS
   ================================================== */

function getHistoryMatrix(excludeRound) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Pairings");
  let historyMap = new Map();
  if (!sheet) return historyMap;
  const data = sheet.getDataRange().getValues();
  let currentRound = 0;
  for (let row of data) {
    if(!row[0]) continue;
    let cell = row[0].toString().toUpperCase();
    if (cell.includes("ROUND")) {
      let match = cell.match(/\d+/);
      currentRound = match ? parseInt(match[0]) : 0;
      continue;
    }
    if (currentRound >= excludeRound) continue;
    if (currentRound > 0 && row[1]) {
      let pIds = [row[1], row[2], row[3], row[4]];
      for (let i = 0; i < 4; i++) {
        for (let j = 0; j < 4; j++) {
          if (i !== j) {
            let p1 = pIds[i];
            let p2 = pIds[j];
            if (!historyMap.has(p1)) historyMap.set(p1, new Set());
            historyMap.get(p1).add(p2);
          }
        }
      }
    }
  }
  return historyMap;
}

// Swap or substitute players in a round's pairings.
//   - Both selected players seated  -> SWAP (exchange their two seats).
//   - One seated + one benched       -> REPLACE (benched player takes the seat,
//                                        the seated player is benched).
// Tables are derived from the pairings, so callers pass only player IDs. Conflicts
// (repeat opponents) and already-scored tables raise a warning that the caller can
// confirm past with force=true.
function swapPairings(round, p1Id, p2Id, force) {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Pairings");
  const data = sheet.getDataRange().getValues();
  const pMap = getPlayerMap();
  const nameOf = (id) => pMap[id] || id;

  if (!p1Id || !p2Id) return { success: false, message: "Select two players." };
  if (p1Id === p2Id) return { success: false, message: "Pick two different players." };

  // Find a player's seat within the round: {rowIdx (0-based), col (1-4), tableId}.
  const locate = (pid) => {
    let inRound = false;
    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      let cell = data[i][0].toString().toUpperCase();
      if (cell.includes("ROUND")) {
        let m = cell.match(/\d+/);
        inRound = (m && parseInt(m[0]) == round);
        continue;
      }
      if (inRound) {
        for (let c = 1; c <= 4; c++) {
          if (data[i][c] == pid) return { rowIdx: i, col: c, tableId: data[i][0] };
        }
      }
    }
    return null;
  };

  const loc1 = locate(p1Id);
  const loc2 = locate(p2Id);
  if (!loc1 && !loc2) {
    return { success: false, message: `Neither player is seated in Round ${round}; nothing to do.` };
  }

  const games = getAllGamesData()[round] || [];
  const isScored = (tid) => { let g = games.find(x => String(x.id) === String(tid)); return !!(g && g.isScored); };
  const histMap = getHistoryMatrix(round);
  // Repeat-opponent conflicts for `inId` joining the table at `loc`, ignoring `outId`.
  const repeatConflicts = (inId, loc, outId) => {
    let out = [];
    let row = data[loc.rowIdx];
    for (let k = 1; k <= 4; k++) {
      if (k === loc.col) continue;
      let opp = row[k];
      if (opp !== "" && opp !== outId && histMap.has(inId) && histMap.get(inId).has(opp)) {
        out.push(`${nameOf(inId)} played ${nameOf(opp)}`);
      }
    }
    return out;
  };

  // --- Both seated: SWAP ---
  if (loc1 && loc2) {
    let warnings = repeatConflicts(p1Id, loc2, p2Id).concat(repeatConflicts(p2Id, loc1, p1Id));
    if (isScored(loc1.tableId) || isScored(loc2.tableId)) {
      warnings.push("A table is already scored — existing scores stay with their original seats; re-enter if needed.");
    }
    if (warnings.length && !force) {
      return { success: false, warning: true, message: "⚠️ " + warnings.join("\n") + "\n\nProceed anyway?" };
    }
    sheet.getRange(loc1.rowIdx + 1, loc1.col + 1).setValue(p2Id);
    sheet.getRange(loc2.rowIdx + 1, loc2.col + 1).setValue(p1Id);
    clearCache();
    return { success: true, message: `✅ Swapped ${nameOf(p1Id)} (Table ${loc1.tableId}) with ${nameOf(p2Id)} (Table ${loc2.tableId}).` };
  }

  // --- One seated: REPLACE (bench the seated player, seat the benched one) ---
  let seat = loc1 || loc2;
  let outId = loc1 ? p1Id : p2Id; // currently seated -> benched
  let inId = loc1 ? p2Id : p1Id;  // currently benched -> takes the seat
  let warnings = repeatConflicts(inId, seat, outId);
  if (isScored(seat.tableId)) {
    warnings.push(`Table ${seat.tableId} is already scored — the existing score stays attributed to ${nameOf(outId)}. Re-enter the score for ${nameOf(inId)} if needed.`);
  }
  if (warnings.length && !force) {
    return { success: false, warning: true, message: "⚠️ " + warnings.join("\n") + "\n\nProceed anyway?" };
  }
  sheet.getRange(seat.rowIdx + 1, seat.col + 1).setValue(inId);
  clearCache();
  return { success: true, message: `✅ ${nameOf(inId)} now seated at Table ${seat.tableId} (replacing ${nameOf(outId)}, now benched).\n\nMark ${nameOf(outId)} as Did Not Finish if they have withdrawn.` };
}

function getPlayerScheduleMatrix() {
  const ss = getDataSS();
  const pairSheet = ss.getSheetByName("Pairings");
  const scoreSheet = ss.getSheetByName("Scores");
  const pMap = getPlayerMap();
  
  let players = {};
  Object.keys(pMap).forEach(k => { 
      players[k] = { id: k, name: pMap[k], tables: {}, scores: {} }; 
  });
  let maxRound = 0;

  if (pairSheet && pairSheet.getLastRow() > 1) {
    const data = pairSheet.getDataRange().getValues();
    let curRound = 0;
    for (let row of data) {
      if(!row[0]) continue;
      let cell = row[0].toString().toUpperCase();
      if (cell.includes("ROUND")) {
        let match = cell.match(/\d+/);
        if (match) {
            curRound = parseInt(match[0]);
            if(curRound > maxRound) maxRound = curRound;
        }
        continue;
      }
      if (curRound > 0 && !isNaN(parseInt(row[0]))) {
        let tId = parseInt(row[0]);
        [1, 2, 3, 4].forEach(c => {
            let pid = row[c];
            if (pid && players[pid]) {
                players[pid].tables[curRound] = tId;
            }
        });
      }
    }
  }

  if (scoreSheet && scoreSheet.getLastRow() > 1) {
    const data = scoreSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let r = data[i][1];
      [3, 6, 9, 12].forEach(idx => {
          let pid = data[i][idx];
          let val = data[i][idx + 2]; 
          if (pid && players[pid]) {
              players[pid].scores[r] = val;
          }
      });
    }
  }

  let list = Object.values(players).sort((a,b) => {
      let na = parseInt(a.id.replace(/\D/g, '')) || 0;
      let nb = parseInt(b.id.replace(/\D/g, '')) || 0;
      return na - nb;
  });
  return { maxRound: maxRound, players: list };
}
// Check in every player at once, skipping substitutes (and withdrawn [DNF]
// players). Existing check-in state for skipped players is left untouched.
function checkInAllPlayers() {
  const ss = getDataSS();
  const sheet = ss.getSheetByName("Players");
  if (!sheet) return getPlayers();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return getPlayers();
  if (data[0][2] !== "Checked In") sheet.getRange(1, 3).setValue("Checked In");

  const colValues = [];
  for (let i = 1; i < data.length; i++) {
    const p = { id: data[i][0], name: data[i][1] };
    const skip = isSubPlayer(p) || String(p.name).toUpperCase().startsWith("[DNF]");
    colValues.push([skip ? data[i][2] : true]);
  }
  sheet.getRange(2, 3, colValues.length, 1).setValues(colValues);
  clearCache();
  return getPlayers();
}

function togglePlayerCheckIn(playerId) {
    const ss = getDataSS();
    const sheet = ss.getSheetByName("Players");
    if (!sheet) return getPlayers();
    const data = sheet.getDataRange().getValues();
    if (data[0][2] !== "Checked In") sheet.getRange(1, 3).setValue("Checked In");
    for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() == playerId.toString()) {
            let cur = data[i][2] === true || data[i][2] === "true" || data[i][2] === "TRUE";
            sheet.getRange(i + 1, 3).setValue(!cur);
            break;
        }
    }
    clearCache();
    return getPlayers();
}

function beginTournamentRepair() {
    const lock = LockService.getScriptLock();
    try { lock.waitLock(30000); } catch (e) { return { success: false, message: "Server busy: Another admin is currently saving. Please try again." }; }

    try {
      const ss = getDataSS();
      const pairSheet = ss.getSheetByName("Pairings");
      const playersSheet = ss.getSheetByName("Players");
      const players = getPlayers();
      
      const missingIds = players.filter(p => !p.isCheckedIn).map(p => String(p.id));
      
      if (missingIds.length === 0) {
          updateSheetSetting(ss, "Tourney_Begun", "true");
          return { success: true, message: "Tournament begun successfully!" };
      }

      const pData = playersSheet.getDataRange().getValues();
      for (let i = 1; i < pData.length; i++) {
          if (missingIds.includes(String(pData[i][0]))) {
              let cur = pData[i][1].toString();
              if (!cur.startsWith("[DNF] ")) {
                  playersSheet.getRange(i + 1, 2).setValue("[DNF] " + cur);
              }
          }
      }
      SpreadsheetApp.flush();

      const pairData = pairSheet.getDataRange().getValues();
      let round1StartIndex = -1;
      let round1EndIndex = -1;
      
      for (let i = 0; i < pairData.length; i++) {
          if (!pairData[i][0]) continue;
          let cell = String(pairData[i][0]).toUpperCase();
          if (cell.includes("ROUND 1")) {
              round1StartIndex = i;
          } else if (cell.includes("ROUND") && round1StartIndex !== -1) {
              round1EndIndex = i;
              break;
          }
      }
      if (round1EndIndex === -1) round1EndIndex = pairData.length;

      if (round1StartIndex === -1) {
          updateSheetSetting(ss, "Tourney_Begun", "true");
          return { success: true, message: "Started, but could not find Round 1 to repair." };
      }

      let tables = [];
      for (let i = round1StartIndex + 1; i < round1EndIndex; i++) {
          let row = pairData[i];
          if (!isNaN(parseInt(row[0]))) {
              tables.push({
                  tableId: row[0],
                  seats: [row[1], row[2], row[3], row[4]],
                  bucket: row[5] || ""
              });
          }
      }

      let moves = [];
      let subsAdded = [];
      tables.forEach(t => {
          for (let s = 0; s < 4; s++) {
              if (missingIds.includes(String(t.seats[s]))) {
                  t.seats[s] = null;
              }
          }
      });
      for (let i = 0; i < tables.length; i++) {
          for (let j = 0; j < 4; j++) {
              if (tables[i].seats[j] === null) {
                  let found = false;
                  for (let k = tables.length - 1; k >= i; k--) {
                      for (let s = 3; s >= 0; s--) {
                          if (k === i && s <= j) continue;
                          if (tables[k].seats[s] !== null) {
                              let movingPid = tables[k].seats[s];
                              let pObj = players.find(p => p.id === movingPid);
                              let pName = pObj ? pObj.name : movingPid;
                              
                              tables[i].seats[j] = movingPid;
                              tables[k].seats[s] = null;
                              
                              // STORE AS OBJECT TO CAPTURE FINAL ID LATER
                              moves.push({ pName: pName, from: tables[k].tableId, toObj: tables[i] });
                              found = true;
                              break;
                          }
                      }
                      if (found) break;
                  }
              }
          }
      }

      tables = tables.filter(t => t.seats.some(seat => seat !== null));
      if (tables.length > 0) {
          let lastTable = tables[tables.length - 1];
          // Number substitutes off existing SUB IDs (collision-safe) and give them
          // "SUB<n>" IDs / "SUBSTITUTE <n>" names, matching generateNextRound.
          let subNums = players.map(p => /^SUB(\d+)$/i.exec(p.id)).filter(Boolean).map(m => parseInt(m[1], 10));
          let nextSub = subNums.length ? Math.max(...subNums) + 1 : 1;

          let needsSub = lastTable.seats.some(seat => seat === null);
          if (needsSub) {
              for (let j = 0; j < 4; j++) {
                  if (lastTable.seats[j] === null) {
                      let subId = "SUB" + nextSub;
                      let subName = "SUBSTITUTE " + nextSub;
                      nextSub++;
                      playersSheet.appendRow([subId, subName, true]);
                      lastTable.seats[j] = subId;
                      
                      // STORE AS OBJECT TO CAPTURE FINAL ID LATER
                      subsAdded.push({ subName: subName, toObj: lastTable });
                  }
              }
              SpreadsheetApp.flush();
          }
      }

      pairSheet.getRange(round1StartIndex + 2, 1, round1EndIndex - (round1StartIndex + 2), 6).clearContent();
      let output = [];
      let tCounter = 1;
      tables.forEach(t => {
          t.finalId = tCounter; // SAVE THE EXACT ASSIGNED NUMBER TO THE OBJECT
          output.push([tCounter++, t.seats[0], t.seats[1], t.seats[2], t.seats[3], t.bucket]);
      });
      output.push(["", "", "", "", "", ""]);

      if (output.length > 0) {
          pairSheet.getRange(round1StartIndex + 2, 1, output.length, 6).setValues(output);
      }

      updateSheetSetting(ss, "Tourney_Begun", "true");
      SpreadsheetApp.flush();
      clearCache();
      
      let finalMessage = "Tournament begun and Round 1 pairings repaired!";
      if (moves.length > 0) {
          // MAP THE FINAL ASSIGNED IDs TO THE LOG STRINGS
          let moveLogs = moves.map(m => `${m.pName} moved from Table ${m.from} -> Table ${m.toObj.finalId}`);
          finalMessage += "\n\n--- Player Movements ---\n" + moveLogs.join("\n");
      }
      if (subsAdded.length > 0) {
          // MAP THE FINAL ASSIGNED IDs TO THE LOG STRINGS
          let subLogs = subsAdded.map(s => `${s.subName} added to Table ${s.toObj.finalId}`);
          finalMessage += "\n\n--- Subs Added ---\n" + subLogs.join("\n");
      }

      return { success: true, message: finalMessage };
    } catch(e) {
      return { success: false, message: "Error repairing pairings: " + e.message };
    } finally {
      lock.releaseLock();
    }
}