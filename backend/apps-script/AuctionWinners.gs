// TMG Auction Winners API
// Deploy as a Google Apps Script web app after setting Script Property API_KEY.
// Public GET returns only winner name, item, bid, and close time.

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
    const winnerName = clean_(payload.winner_name, 80);
    const winnerUuid = clean_(payload.winner_uuid, 80);
    const bid = Number(payload.winning_bid);
    const closedUtc = clean_(payload.auction_closed_utc, 80);
    if (!requestId || !boardId || !item || !winnerName || !winnerUuid
        || !Number.isFinite(bid) || bid < 0 || !closedUtc) {
      return jsonOutput_({ ok: false, error: "invalid_record" });
    }

    lock.waitLock(10000);
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) return jsonOutput_({ ok: false, error: "sheet_missing" });

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const duplicate = sheet.getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(requestId).matchEntireCell(true).findNext();
      if (duplicate) return jsonOutput_({ ok: true, duplicate: true });
    }

    sheet.appendRow([
      new Date(), requestId, boardId, item, winnerName, winnerUuid,
      Math.floor(bid), "Received", closedUtc
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
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return publicOutput_(e, []);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    const winners = values
      .filter(row => row[7] === "Received")
      .slice(-100)
      .reverse()
      .map(row => ({
        closed: String(row[8]),
        item: String(row[3]),
        winner: String(row[4]),
        bid: Number(row[6]) || 0
      }));
    return publicOutput_(e, winners);
  } catch (err) {
    return publicOutput_(e, [], String(err && err.message || err));
  }
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
  return String(value == null ? "" : value).trim().replace(/[\u0000-\u001f\u007f]/g, "").slice(0, maxLength);
}
