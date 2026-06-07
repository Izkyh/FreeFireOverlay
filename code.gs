// =========================================================================
// PUBG MOBILE OVERLAY - PURE SCRIPT AGGREGATION (NO FORMULAS)
// =========================================================================

function getPlacementPts(rank) {
  var table = [12, 9, 8, 7, 6, 5, 4, 3, 2, 1];
  var r = parseInt(rank);
  if (r >= 1 && r <= 10) return table[r - 1];
  return 0;
}

// =========================================================================
// 1. SETUP SHEETS
// =========================================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var DARK_BG = "#111111";
  var WHITE_TXT = "#FFFFFF";

  var historySheet = ss.getSheetByName("MATCH_HISTORY");
  if (historySheet) ss.deleteSheet(historySheet);

  var dbSheet = ss.getSheetByName("DB_TEAMS") || ss.insertSheet("DB_TEAMS");
  dbSheet.getRange("A1:B1").setValues([["NAMA_TEAM", "LOGO_URL"]]).setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");
  
  var liveSheet = ss.getSheetByName("LIVE_MATCH") || ss.insertSheet("LIVE_MATCH");
  liveSheet.clear();
  liveSheet.getRange("A1:F1").setValues([["NAMA_TEAM", "ALIVE", "ELIMS", "TRIGGER ELIMINATE", "TRIGGER WWCD", "RANK SEMENTARA"]])
           .setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");

  liveSheet.getRange("G1").setValue("NAMA MATCH").setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");
  liveSheet.getRange("G2").setValue("Match 1").setBackground("#FFFFFF").setHorizontalAlignment("center");

  liveSheet.getRange("H1").setValue("SIMPAN HASIL").setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");
  liveSheet.getRange("H2").insertCheckboxes().setHorizontalAlignment("center");

  liveSheet.getRange("I1").setValue("JUMLAH TEAM").setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");
  var teamRule = SpreadsheetApp.newDataValidation().requireValueInList(["16", "17", "18", "19", "20"]).build();
  liveSheet.getRange("I2").setDataValidation(teamRule).setValue("20").setHorizontalAlignment("center");

  liveSheet.getRange("J1").setValue("RESET SEMUA").setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");
  liveSheet.getRange("J2").insertCheckboxes().setBackground("#FF4444").setHorizontalAlignment("center");

  var dbNames = dbSheet.getRange("A2:A21").getValues();
  var liveDefaults = [];
  for (var i = 0; i < 20; i++) {
    liveDefaults.push([dbNames[i][0] || "", 4, 0, false, false, ""]);
  }
  liveSheet.getRange(2, 1, 20, 6).setValues(liveDefaults);

  var aliveRule = SpreadsheetApp.newDataValidation().requireValueInList(["4", "3", "2", "1", "0"]).build();
  liveSheet.getRange(2, 2, 20, 1).setDataValidation(aliveRule).setHorizontalAlignment("center");
  liveSheet.getRange(2, 4, 20, 1).insertCheckboxes().setHorizontalAlignment("center");
  liveSheet.getRange(2, 5, 20, 1).insertCheckboxes().setBackground("#FFD700").setHorizontalAlignment("center");

  var mrSheet = ss.getSheetByName("MATCH_RANKING") || ss.insertSheet("MATCH_RANKING");
  mrSheet.getRange("A1:G1").setValues([["MATCH_NAME", "NAMA_TEAM", "RANK", "PLACEMENT_POINT", "KILL_PTS", "TOTAL", "WWCD"]])
         .setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");

  var orSheet = ss.getSheetByName("OVERALL_RANKING") || ss.insertSheet("OVERALL_RANKING");
  orSheet.clear();
  orSheet.getRange("A1:E1").setValues([["NAMA_TEAM", "PLACEMENT_POINT", "KILL_PTS", "TOTAL", "TOTAL_WWCD"]])
         .setBackground(DARK_BG).setFontColor(WHITE_TXT).setFontWeight("bold").setHorizontalAlignment("center");
  
  _updateOverallRanking(ss); // Inisialisasi perhitungan dari script

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert("✅ Bug Diperbaiki! Semua formula dihapus dan diganti kalkulasi Script murni.");
}

// =========================================================================
// 2. LOGIKA AKUMULASI OVERALL RANKING (MENGGANTIKAN FORMULA)
// =========================================================================
function _updateOverallRanking(ss) {
  var dbSheet = ss.getSheetByName("DB_TEAMS");
  var mrSheet = ss.getSheetByName("MATCH_RANKING");
  var orSheet = ss.getSheetByName("OVERALL_RANKING");

  var teams = dbSheet.getRange("A2:A21").getValues();
  var teamStats = {};
  for (var i = 0; i < teams.length; i++) {
    var tName = String(teams[i][0]).trim();
    if (tName) {
      teamStats[tName] = { place: 0, kill: 0, total: 0, wwcd: 0 };
    }
  }

  var lastRow = mrSheet.getLastRow();
  if (lastRow > 1) {
    var historyData = mrSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var j = 0; j < historyData.length; j++) {
      var tName = String(historyData[j][1]).trim();
      if (teamStats[tName]) {
        teamStats[tName].place += (parseInt(historyData[j][3]) || 0);
        teamStats[tName].kill  += (parseInt(historyData[j][4]) || 0);
        teamStats[tName].total += (parseInt(historyData[j][5]) || 0);
        teamStats[tName].wwcd  += (parseInt(historyData[j][6]) || 0);
      }
    }
  }

  var orData = [];
  for (var k = 0; k < 20; k++) {
    var tName = String(teams[k][0]).trim();
    if (tName && teamStats[tName]) {
      // Mengirim angka 0 sebagai default jika poin adalah 0
      orData.push([
        tName, 
        teamStats[tName].place || 0, 
        teamStats[tName].kill || 0, 
        teamStats[tName].total || 0, 
        teamStats[tName].wwcd || 0
      ]);
    } else {
      orData.push(["", "", "", "", ""]);
    }
  }
  orSheet.getRange(2, 1, 20, 5).setValues(orData);
}

// =========================================================================
// 3. INTERAKTIVITAS (onEdit)
// =========================================================================
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var val = e.value;

  if (sheet.getName() === "DB_TEAMS" && col === 1 && row >= 2) {
    var ss = e.source;
    var liveSheet = ss.getSheetByName("LIVE_MATCH");
    var orSheet = ss.getSheetByName("OVERALL_RANKING");
    var allNames = sheet.getRange(2, 1, 20, 1).getValues();
    if(liveSheet) liveSheet.getRange(2, 1, 20, 1).setValues(allNames);
    if(orSheet) orSheet.getRange(2, 1, 20, 1).setValues(allNames);
    return;
  }

  if (sheet.getName() !== "LIVE_MATCH") return;

  if (col === 4 && row >= 2 && val === "TRUE") {
    sheet.getRange(row, 2).setValue(0);
    var totalTeams = parseInt(sheet.getRange("I2").getValue()) || 20;
    var aliveData = sheet.getRange(2, 2, 20, 1).getValues();
    var deadCount = 0;
    for (var i = 0; i < totalTeams; i++) {
      if (parseInt(aliveData[i][0]) === 0) deadCount++;
    }
    var rank = totalTeams - deadCount + 1;
    sheet.getRange(row, 6).setValue(rank);
    sheet.getRange(row, 1, 1, 6).setBackground("#FFCCCC");
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    sheet.getRange(row, 4).uncheck();
  }

  if (col === 5 && row >= 2 && val === "TRUE") {
    sheet.getRange(row, 6).setValue(1);
    sheet.getRange(row, 1, 1, 6).setBackground("#D4EDDA");
  }

  if (col === 9 && row === 2) {
    var jml = parseInt(val) || 20;
    sheet.showRows(2, 20);
    if (jml < 20) sheet.hideRows(jml + 2, 20 - jml);
  }

  if (col === 8 && row === 2 && val === "TRUE") {
    _handleSaveResult(e.source, sheet);
  }

  if (col === 10 && row === 2 && val === "TRUE") {
    var res = Browser.msgBox("⚠️ BAHAYA", "Yakin mereset SEMUA data riwayat match?", Browser.Buttons.YES_NO);
    if (res === "yes") {
      var mrSheet = e.source.getSheetByName("MATCH_RANKING");
      if(mrSheet.getLastRow() > 1) mrSheet.getRange(2, 1, mrSheet.getLastRow()-1, 7).clearContent().setBackground("#FFFFFF");
      _resetLiveMatch(sheet, true);
      _updateOverallRanking(e.source); // Reset Overall Ranking jadi 0
    }
    sheet.getRange("J2").uncheck();
  }
}

// =========================================================================
// 4. SIMPAN HASIL WORKFLOW
// =========================================================================
function _handleSaveResult(ss, liveSheet) {
  var matchName = liveSheet.getRange("G2").getValue();
  if (!matchName || matchName.toString().trim() === "") {
    Browser.msgBox("⚠️ Peringatan", "Isi NAMA MATCH di sel G2!", Browser.Buttons.OK);
    liveSheet.getRange("H2").uncheck();
    return;
  }

  var response = Browser.msgBox("💾 Simpan Data", "Simpan hasil " + matchName + " ke History?", Browser.Buttons.YES_NO);
  if (response !== "yes") {
    liveSheet.getRange("H2").uncheck();
    return;
  }

  var mrSheet = ss.getSheetByName("MATCH_RANKING");
  var totalTeams = parseInt(liveSheet.getRange("I2").getValue()) || 20;
  var liveData = liveSheet.getRange(2, 1, totalTeams, 6).getValues();
  
  // FIX: Hapus awalan "=" agar tidak error. Ganti jadi "-- MATCH 1 --"
  mrSheet.appendRow(["-- " + matchName.toUpperCase() + " --", "", "", "", "", "", ""]);
  var lastRow = mrSheet.getLastRow();
  mrSheet.getRange(lastRow, 1, 1, 7).setBackground("#CCCCCC").setFontWeight("bold");

  var rowsToAppend = [];
  for (var i = 0; i < totalTeams; i++) {
    var teamName = String(liveData[i][0]).trim();
    var elims = parseInt(liveData[i][2]) || 0;
    var rankTemp = liveData[i][5];
    var rank = (rankTemp !== "" && rankTemp !== null && rankTemp !== false && !isNaN(parseInt(rankTemp))) ? parseInt(rankTemp) : 0; 
    var isWWCD = liveData[i][4] === true;
    
    var placePts = getPlacementPts(rank);
    var wwcdVal = isWWCD ? 1 : 0;
    rowsToAppend.push([matchName, teamName, rank === 0 ? "-" : rank, placePts, elims, (placePts + elims), wwcdVal]);
  }

  mrSheet.getRange(mrSheet.getLastRow() + 1, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
  _resetLiveMatch(liveSheet, false);
  
  // UPDATE OVERALL RANKING OTOMATIS
  _updateOverallRanking(ss);
}

function _resetLiveMatch(liveSheet, isFullReset) {
  var clearData = [];
  for (var i = 0; i < 20; i++) clearData.push([4, 0, false, false, ""]);
  liveSheet.getRange(2, 2, 20, 5).setValues(clearData);
  liveSheet.getRange(2, 1, 20, 6).setBackground("#FFFFFF");
  liveSheet.getRange("H2").uncheck();
  if(!isFullReset) liveSheet.getRange("G2").clearContent();
  SpreadsheetApp.flush();
}

// =========================================================================
// 5. API ENDPOINT
// =========================================================================
function doGet(e) {
  var callback = e && e.parameter ? e.parameter.callback : "";
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dbSheet = ss.getSheetByName("DB_TEAMS");
    var liveSheet = ss.getSheetByName("LIVE_MATCH");
    var mrSheet = ss.getSheetByName("MATCH_RANKING");
    var orSheet = ss.getSheetByName("OVERALL_RANKING");

    // ── 1. LOGO MAP ──────────────────────────────────────────────────────────
    var logoMap = {};
    var dbRows = dbSheet.getRange(2, 1, 20, 2).getValues();
    dbRows.forEach(function(r) { if (r[0]) logoMap[String(r[0]).trim()] = r[1] || ""; });

    // ── 2. MATCH HISTORY (Semua Data dari MATCH_RANKING Sheet) ─────────────────
    var match_history = [];
    var postmatch = [];
    if (mrSheet && mrSheet.getLastRow() > 1) {
      var mrRows = mrSheet.getRange(2, 1, mrSheet.getLastRow() - 1, 7).getValues();
      var lastMatchNum = "";

      for (var i = 0; i < mrRows.length; i++) {
        var r = mrRows[i];
        var cellA = String(r[0]).trim();
        
        // Skip separator row but update lastMatchNum if we want, or just rely on column A
        if (cellA.indexOf("--") !== -1) continue;

        if (r[1] && String(r[1]).trim() !== "") {
           var matchName = cellA;
           var mStr = matchName;
           var mMatch = matchName.match(/\d+/);
           if (mMatch) mStr = mMatch[0]; // extract just the number "1", "2"
           
           lastMatchNum = mStr; // Keep track of the last seen match number

           match_history.push({
              match:     mStr,
              name:      String(r[1]).trim(),
              rank:      r[2],
              place_pts: parseInt(r[3]) || 0,
              kill_pts:  parseInt(r[4]) || 0,
              total:     parseInt(r[5]) || 0,
              wwcd:      parseInt(r[6]) || 0,
              logo_url:  logoMap[String(r[1]).trim()] || ""
           });
        }
      }

      // postmatch is the latest match
      postmatch = match_history.filter(function(t) { return t.match === lastMatchNum; });
      postmatch.sort(function(a, b) {
        var ra = (a.rank === "-" || a.rank === "" || a.rank === 0) ? 999 : parseInt(a.rank);
        var rb = (b.rank === "-" || b.rank === "" || b.rank === 0) ? 999 : parseInt(b.rank);
        return ra - rb;
      });
    }

    // ── 3. LIVE DATA (LIVE_MATCH) ─────────────────────────────────────────────
    var totalTeams = parseInt(liveSheet.getRange("I2").getValue()) || 20;
    var live_data = [];
    // Baca 6 kolom (A-F) agar rank sementara (kolom F) ikut terbaca
    var liveRows = liveSheet.getRange(2, 1, totalTeams, 6).getValues();
    liveRows.forEach(function(r, idx) {
      if (!r[0]) return;
      var rankTemp = r[5];
      var rankVal = (rankTemp !== "" && rankTemp !== null && !isNaN(parseInt(rankTemp)))
                    ? parseInt(rankTemp) : "";
      live_data.push({
        id:       idx + 1,
        name:     String(r[0]).trim(),
        logo_url: logoMap[String(r[0]).trim()] || "",
        alive:    parseInt(r[1]) || 0,
        status:   parseInt(r[1]) === 0 ? "DEAD" : "ALIVE",
        kills:    parseInt(r[2]) || 0,
        rank:     rankVal
      });
    });

    // ── 4. OVERALL RANKING ────────────────────────────────────────────────────
    var overall_ranking = [];
    var orRows = orSheet.getRange(2, 1, totalTeams, 5).getValues();
    orRows.forEach(function(r, idx) {
      if (!r[0]) return;
      overall_ranking.push({
        id:        idx + 1,
        name:      String(r[0]).trim(),
        logo_url:  logoMap[String(r[0]).trim()] || "",
        place_pts: parseInt(r[1]) || 0,
        kill_pts:  parseInt(r[2]) || 0,
        total:     parseInt(r[3]) || 0,
        wwcd:      parseInt(r[4]) || 0
      });
    });

    overall_ranking.sort(function(a, b) {
      return b.total - a.total || b.wwcd - a.wwcd || b.kill_pts - a.kill_pts;
    });

    // ── 5. RESPONSE ───────────────────────────────────────────────────────────
    var result = {
      postmatch:       postmatch,
      match_history:   match_history, // history lengkap
      match_ranking:   live_data,   // live data (kompatibel dengan key lama)
      teams:           live_data,   // alias tambahan
      overall_ranking: overall_ranking
    };

    var json = JSON.stringify(result);
    return callback
      ? ContentService.createTextOutput(callback + "(" + json + ")").setMimeType(ContentService.MimeType.JAVASCRIPT)
      : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    var errJson = JSON.stringify({ error: err.toString() });
    return callback
      ? ContentService.createTextOutput(callback + "(" + errJson + ")").setMimeType(ContentService.MimeType.JAVASCRIPT)
      : ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON);
  }
}