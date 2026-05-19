var LIVE_SHEET_NAME = "FF";
var POSTMATCH_SHEET_NAME = "POSTMATCH";

var LIVE_HEADERS = [
  "id",
  "name",
  "alive",
  "kills",
  "rank",
  "wwcd"
];

var POSTMATCH_HEADERS = [
  "match",
  "id",
  "name",
  "rank",
  "kills",
  "wwcd"
];

var DUMMY_TEAMS = [
  "ABN",
  "AUT",
  "B4",
  "BLACK",
  "EPIC",
  "FVG",
  "GS",
  "INCO",
  "IZ",
  "KP",
  "NB",
  "NOT",
  "PNG",
  "S4",
  "SH",
  "SR",
  "TITAN",
  "WSU",
  "XTEAM",
  "TEAM 20"
];

function getPlacementPts(rank) {
  var table = [12, 9, 8, 7, 6, 5, 4, 3, 2, 1];
  if (rank >= 1 && rank <= 10) return table[rank - 1];
  return 0;
}

function toInt(value, fallback) {
  var parsed = parseInt(value, 10);
  return isNaN(parsed) ? fallback : parsed;
}

function normalizeAlive(value) {
  var alive = toInt(value, 4);
  return Math.max(0, Math.min(4, alive));
}

function getStatus(alive) {
  return alive === 0 ? "DEAD" : "ALIVE";
}

function withPoints(row) {
  var placePts = getPlacementPts(row.rank);
  var killPts = row.kills;
  row.place_pts = placePts;
  row.kill_pts = killPts;
  row.total = placePts + killPts;
  return row;
}

function doGet(e) {
  var callback = e && e.parameter ? e.parameter.callback : "";

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var teams = readLiveTeams(ss);
    var postmatch = readPostmatch(ss);

    return sendResponse(
      {
        teams: teams,
        postmatch: postmatch,
        updated: new Date().toISOString()
      },
      callback
    );
  } catch (err) {
    return sendResponse({ error: err.toString() }, callback);
  }
}

function readLiveTeams(ss) {
  var sheet = ss.getSheetByName(LIVE_SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet '" + LIVE_SHEET_NAME + "' tidak ditemukan.");
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var rows = sheet.getRange(2, 1, lastRow - 1, LIVE_HEADERS.length).getValues();
  var teams = [];

  rows.forEach(function (row) {
    var id = toInt(row[0], 0);
    if (!id) return;

    var alive = normalizeAlive(row[2]);
    var kills = Math.max(0, toInt(row[3], 0));
    var rank = Math.max(1, toInt(row[4], id));
    var wwcd = Math.max(0, toInt(row[5], 0));

    teams.push(
      withPoints({
        id: id,
        name: String(row[1] || "TEAM " + id).trim(),
        alive: alive,
        status: getStatus(alive),
        kills: kills,
        rank: rank,
        wwcd: wwcd
      })
    );
  });

  return teams;
}

function readPostmatch(ss) {
  var sheet = ss.getSheetByName(POSTMATCH_SHEET_NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var rows = sheet
    .getRange(2, 1, lastRow - 1, POSTMATCH_HEADERS.length)
    .getValues();
  var results = [];

  rows.forEach(function (row) {
    var match = toInt(row[0], 0);
    var id = toInt(row[1], 0);
    if (!match || !id) return;

    var rank = Math.max(1, toInt(row[3], id));
    var kills = Math.max(0, toInt(row[4], 0));
    var wwcd = Math.max(0, toInt(row[5], rank === 1 ? 1 : 0));

    results.push(
      withPoints({
        match: match,
        id: id,
        name: String(row[2] || "TEAM " + id).trim(),
        rank: rank,
        kills: kills,
        wwcd: wwcd
      })
    );
  });

  results.sort(function (a, b) {
    return a.match - b.match || a.rank - b.rank;
  });

  return results;
}

function setupDummyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupLiveSheet(ss);
  setupPostmatchSheet(ss);
}

function setupLiveSheet(ss) {
  var sheet = ss.getSheetByName(LIVE_SHEET_NAME) || ss.insertSheet(LIVE_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, LIVE_HEADERS.length).setValues([LIVE_HEADERS]);

  var rows = DUMMY_TEAMS.map(function (name, index) {
    var id = index + 1;
    var rank = id;
    var kills = Math.max(0, 21 - id);
    var alive = id === 20 ? 0 : Math.max(1, 5 - Math.ceil(id / 5));
    var wwcd = id === 1 || id === 4 ? 1 : 0;
    return [id, name, alive, kills, rank, wwcd];
  });

  sheet.getRange(2, 1, rows.length, LIVE_HEADERS.length).setValues(rows);
  sheet.autoResizeColumns(1, LIVE_HEADERS.length);
}

function setupPostmatchSheet(ss) {
  var sheet =
    ss.getSheetByName(POSTMATCH_SHEET_NAME) || ss.insertSheet(POSTMATCH_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, POSTMATCH_HEADERS.length).setValues([POSTMATCH_HEADERS]);

  var rows = [];
  for (var match = 1; match <= 2; match++) {
    DUMMY_TEAMS.forEach(function (name, index) {
      var id = index + 1;
      var rank = ((index + (match - 1) * 7) % 20) + 1;
      var kills = Math.max(0, 20 - rank + match);
      var wwcd = rank === 1 ? 1 : 0;
      rows.push([match, id, name, rank, kills, wwcd]);
    });
  }

  sheet.getRange(2, 1, rows.length, POSTMATCH_HEADERS.length).setValues(rows);
  sheet.autoResizeColumns(1, POSTMATCH_HEADERS.length);
}

function sendResponse(data, callback) {
  var json = JSON.stringify(data);
  var output;

  if (callback) {
    output = ContentService.createTextOutput(callback + "(" + json + ")");
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    return output;
  }

  output = ContentService.createTextOutput(json);
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
