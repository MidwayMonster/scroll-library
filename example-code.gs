/**
 * Arcane Scroll Ledger - Google Apps Script backend.
 *
 * SETUP (run once, then redeploy):
 *   1. Apps Script editor > Project Settings > Script Properties, add:
 *        PLAYER_KEY = <your player password>
 *        DM_KEY     = <your DM password>
 *      (Or run setupKeys() below once, then blank out the literals again.)
 *   2. Deploy > Manage deployments > edit the existing deployment > New version.
 *      This keeps the same /exec URL that index.html points at.
 *
 * The keys live only in Script Properties, never in this file or the public repo.
 */

var SHEET_NAME = "Scrolls";

// Column positions (1-based) in the Scrolls sheet.
var COL_ID = 1;
var COL_SPELL = 2;
var COL_SPELL_LEVEL = 3;
var COL_SAVE_DC = 4;
var COL_ATTACK_BONUS = 5;
var COL_CASTING_MODIFIER = 6;
var COL_HOLDER = 7;
var COL_SCRIBED_BY = 8;
var COL_STATUS = 9;
var COL_HOURS_COMPLETED = 10;
var COL_TOTAL_HOURS = 11;
var COL_CREATED_DATE = 12;
var COL_CONSUMED_DATE = 13;

/** Run once from the editor to seed the keys, then clear the literals below. */
function setupKeys() {
  PropertiesService.getScriptProperties().setProperties({
    PLAYER_KEY: "",
    DM_KEY: ""
  });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Returns "dm", "player", null (bad key), or "unconfigured". */
function roleFor(key) {
  var props = PropertiesService.getScriptProperties();
  var dmKey = props.getProperty("DM_KEY");
  var playerKey = props.getProperty("PLAYER_KEY");

  if (!dmKey && !playerKey) return "unconfigured";
  if (dmKey && key === dmKey) return "dm";
  if (playerKey && key === playerKey) return "player";
  return null;
}

function toNumber(value) {
  var n = Number(value);
  return isNaN(n) ? 0 : n;
}

function doGet(e) {
  var key = (e && e.parameter && e.parameter.key) || "";
  var role = roleFor(key);

  if (role === "unconfigured") {
    return jsonOut({ error: "Server keys not configured. Set PLAYER_KEY and DM_KEY in Script Properties." });
  }
  if (!role) return jsonOut({ error: "unauthorized" });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonOut({ error: "Sheet '" + SHEET_NAME + "' not found." });

  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i] || data[i][0] === "") continue;
    rows.push({
      ID: data[i][COL_ID - 1].toString().trim(),
      Spell: data[i][COL_SPELL - 1],
      SpellLevel: data[i][COL_SPELL_LEVEL - 1],
      SaveDC: data[i][COL_SAVE_DC - 1],
      AttackBonus: data[i][COL_ATTACK_BONUS - 1],
      CastingModifier: data[i][COL_CASTING_MODIFIER - 1],
      Holder: data[i][COL_HOLDER - 1],
      ScribedBy: data[i][COL_SCRIBED_BY - 1],
      Status: data[i][COL_STATUS - 1],
      HoursCompleted: data[i][COL_HOURS_COMPLETED - 1],
      TotalHoursRequired: data[i][COL_TOTAL_HOURS - 1],
      CreatedDate: data[i][COL_CREATED_DATE - 1],
      ConsumedDate: data[i][COL_CONSUMED_DATE - 1]
    });
  }

  return jsonOut({ role: role, scrolls: rows });
}

function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ error: "Malformed request body." });
  }

  var role = roleFor(params.key || "");
  if (role === "unconfigured") {
    return jsonOut({ error: "Server keys not configured. Set PLAYER_KEY and DM_KEY in Script Properties." });
  }
  if (role !== "dm") return jsonOut({ error: "unauthorized" });

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonOut({ error: "Sheet '" + SHEET_NAME + "' not found." });

  var today = new Date().toLocaleDateString();

  if (params.action === "create") {
    var newTotal = toNumber(params.totalHoursRequired);
    var newDone = toNumber(params.hoursCompleted);
    var newStatus = params.status === "In Progress" ? "In Progress" : "Ready";

    // A scroll that has met its hour requirement is finished, whatever the form said.
    if (newStatus === "In Progress" && newTotal > 0 && newDone >= newTotal) newStatus = "Ready";
    if (newStatus === "Ready" && newTotal > 0) newDone = newTotal;

    sheet.appendRow([
      "'" + params.id,
      params.spell,
      params.spellLevel,
      params.saveDC,
      "'" + params.attackBonus,
      params.castingModifier,
      params.holder,
      params.scribedBy,
      newStatus,
      newDone,
      newTotal,
      today,
      ""
    ]);
    return jsonOut({ status: "success", appliedStatus: newStatus, hoursCompleted: newDone });
  }

  var data = sheet.getDataRange().getValues();
  var targetId = (params.id || "").toString().trim();

  for (var i = 1; i < data.length; i++) {
    if (data[i][COL_ID - 1].toString().trim() !== targetId) continue;

    var row = i + 1;
    var total = toNumber(data[i][COL_TOTAL_HOURS - 1]);

    if (params.action === "consume") {
      sheet.getRange(row, COL_STATUS).setValue("Consumed");
      sheet.getRange(row, COL_CONSUMED_DATE).setValue(today);
      return jsonOut({ status: "success" });
    }

    if (params.action === "updateHours") {
      var hours = toNumber(params.hours);
      var autoCompleted = total > 0 && hours >= total;
      if (autoCompleted) {
        hours = total;
        sheet.getRange(row, COL_STATUS).setValue("Ready");
      }
      sheet.getRange(row, COL_HOURS_COMPLETED).setValue(hours);
      return jsonOut({ status: "success", autoCompleted: autoCompleted, hoursCompleted: hours });
    }

    if (params.action === "completeScribing") {
      sheet.getRange(row, COL_STATUS).setValue("Ready");
      if (total > 0) sheet.getRange(row, COL_HOURS_COMPLETED).setValue(total);
      return jsonOut({ status: "success", hoursCompleted: total });
    }

    return jsonOut({ error: "Unknown action: " + params.action });
  }

  return jsonOut({ status: "not_found" });
}
