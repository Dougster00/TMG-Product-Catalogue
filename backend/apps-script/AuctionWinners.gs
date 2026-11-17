// TMG Auction Winners and Prize Claims API
// Deploy as a Google Apps Script web app.
// POST winner uploads use Script Property API_KEY.
// Private LSL prize claims use Script Property KIOSK_KEY.

const SPREADSHEET_ID = "1iB-shU0JLF_Vc6lHhPumBSjVQgdc06_jbcfMxA9p6HE";
const SHEET_NAME = "Sheet1";

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const expectedKey = PropertiesService.getScriptProperties().getProperty("API_KEY");
    if (!expectedKey || payload.api_key !== expectedKey) {
      return jsonOutput_({ ok: false, error: "unauthorized" });
    }

    const requestId = clean_(payload.request_id, 160);
    const boardId = clean_(payload.board_id, 80);
    const item = clean_(payload.item, 160);
    const itemObject = clean_(payload.item_object || item, 160);
    const winnerName = clean_(payload.winner_name, 80);
    const winnerUuid = clean_(payload.winner_uuid, 80);
    const bid = Number(payload.winning_bid);
    const closedUtc = clean_(payload.auction_closed_utc, 80);
    if (!requestId || !boardId || !item || !itemObject || !winnerName || !winnerUuid
        || !Number.isFinite(bid) || bid < 0 || !closedUtc) {
      return jsonOutput_({ ok: false, error: "invalid_record" });
    }

    lock.waitLock(10000);
    const sheet = getSheet_();
    if (!sheet) return jsonOutput_({ ok: false, error: "sheet_missing" });

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const duplicate = sheet.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(requestId).matchEntireCell(true).findNext();
      if (duplicate) return jsonOutput_({ ok: true, duplicate: true });
    }

    sheet.appendRow([
      new Date(), requestId, boardId, item, winnerName, winnerUuid,
      Math.floor(bid), "Received", closedUtc, itemObject, ""
    ]);
    return jsonOutput_({ ok: true, duplicate: false });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    if (params.action === "claim") return claimPrize_(params);
    if (params.action === "delivered") return markDelivered_(params);

    // Winner names and bids stay private in the Google Sheet.
    return jsonOutput_({ ok: true, winners: [] });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  }
}

function claimPrize_(params) {
  if (!validKioskKey_(params.claim_key))
    return jsonOutput_({ ok: false, error: "unauthorized" });
  const uuid = clean_(params.uuid, 80).toLowerCase();
  if (!/^[0-9a-f-]{36}$/.test(uuid))
    return jsonOutput_({ ok: false, error: "invalid_uuid" });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet_();
    if (!sheet || sheet.getLastRow() < 2)
      return jsonOutput_({ ok: false, error: "no_prize" });
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (String(row[5]).toLowerCase() === uuid && row[7] === "Received") {
        return jsonOutput_({
          ok: true,
          request_id: String(row[1]),
          item_object: String(row[9] || row[3]),
          item: String(row[3])
        });
      }
    }
    return jsonOutput_({ ok: false, error: "no_prize" });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function markDelivered_(params) {
  if (!validKioskKey_(params.claim_key))
    return jsonOutput_({ ok: false, error: "unauthorized" });
  const requestId = clean_(params.request_id, 160);
  const uuid = clean_(params.uuid, 80).toLowerCase();
  if (!requestId || !/^[0-9a-f-]{36}$/.test(uuid))
    return jsonOutput_({ ok: false, error: "invalid_request" });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getSheet_();
    if (!sheet || sheet.getLastRow() < 2)
      return jsonOutput_({ ok: false, error: "not_found" });
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (String(row[1]) === requestId && String(row[5]).toLowerCase() === uuid) {
        if (row[7] === "Received") {
          const sheetRow = i + 2;
          sheet.getRange(sheetRow, 8).setValue("Delivered");
          sheet.getRange(sheetRow, 11).setValue(new Date());
        }
        return jsonOutput_({ ok: true });
      }
    }
    return jsonOutput_({ ok: false, error: "not_found" });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function validKioskKey_(provided) {
  const expected = PropertiesService.getScriptProperties().getProperty("KIOSK_KEY");
  return !!expected && String(provided || "") === expected;
}

function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (sheet && !sheet.getRange(1, 10).getValue()) sheet.getRange(1, 10).setValue("Prize Object");
  if (sheet && !sheet.getRange(1, 11).getValue()) sheet.getRange(1, 11).setValue("Delivered (UTC)");
  return sheet;
}

function publicOutput_(e, winners, error) {
  const data = { ok: !error, winners: winners };
  if (error) data.error = error;
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$]{0,80}$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(data) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonOutput_(data);
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(value, maxLength) {
  return String(value == null ? "" : value).trim()
    .replace(/[\u0000-\u001f\u007f]/g, "").slice(0, maxLength);
}
