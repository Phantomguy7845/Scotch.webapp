const CONFIG = {
  SHEET_NAME: "Requests",
  ADMIN_KEY: "CHANGE_ME_TO_A_STRONG_ADMIN_KEY",
  TIME_ZONE: "Asia/Bangkok",
  APP_NAME: "Scotch webapps1",
};

const HEADERS = [
  "requestId",
  "submittedAt",
  "employeeName",
  "department",
  "email",
  "phone",
  "tripDate",
  "startTime",
  "endTime",
  "pickup",
  "destination",
  "passengers",
  "purpose",
  "notes",
  "status",
  "approvedAt",
  "approvedBy",
  "emailStatus",
];

function doGet(e) {
  try {
    const action = ((e && e.parameter && e.parameter.action) || "health").toLowerCase();
    if (action === "health") {
      return jsonResponse({
        ok: true,
        app: CONFIG.APP_NAME,
        sheet: CONFIG.SHEET_NAME,
        time: nowIso(),
      });
    }
    if (action === "listrequests") {
      verifyAdminKey(e.parameter.adminKey);
      return jsonResponse({ ok: true, requests: readRequests() });
    }
    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function doPost(e) {
  try {
    const payload = parsePayload(e);
    const action = String(payload.action || "").toLowerCase();

    if (action === "submitrequest") {
      return submitRequest(payload.data || {});
    }
    if (action === "approverequest") {
      return approveRequest(payload);
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (error) {
    return jsonResponse({ ok: false, error: error.message });
  }
}

function submitRequest(data) {
  validateRequired(data, [
    "employeeName",
    "department",
    "email",
    "phone",
    "tripDate",
    "startTime",
    "endTime",
    "pickup",
    "destination",
    "passengers",
    "purpose",
  ]);

  const sheet = getOrCreateSheet();
  const requestId = "REQ-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  const now = nowIso();
  const row = [
    requestId,
    now,
    safeText(data.employeeName),
    safeText(data.department),
    safeText(data.email),
    safeText(data.phone),
    safeText(data.tripDate),
    safeText(data.startTime),
    safeText(data.endTime),
    safeText(data.pickup),
    safeText(data.destination),
    safeText(data.passengers),
    safeText(data.purpose),
    safeText(data.notes || ""),
    "PENDING",
    "",
    "",
    "",
  ];

  sheet.appendRow(row);

  return jsonResponse({
    ok: true,
    requestId: requestId,
    message: "Request submitted successfully.",
  });
}

function approveRequest(payload) {
  verifyAdminKey(payload.adminKey);
  const requestId = safeText(payload.requestId);
  if (!requestId) {
    throw new Error("requestId is required.");
  }

  const approvedBy = safeText(payload.approvedBy || "Fleet Admin");
  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    throw new Error("No requests found.");
  }

  let foundRow = -1;
  for (let i = 1; i < values.length; i += 1) {
    if (String(values[i][0]) === requestId) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    throw new Error("Request not found: " + requestId);
  }

  const statusCell = sheet.getRange(foundRow, 15);
  const currentStatus = String(statusCell.getValue() || "");
  if (currentStatus.toUpperCase() === "APPROVED") {
    return jsonResponse({
      ok: true,
      message: "Request already approved.",
      emailStatus: sheet.getRange(foundRow, 18).getValue(),
    });
  }

  const approvedAt = nowIso();
  statusCell.setValue("APPROVED");
  sheet.getRange(foundRow, 16).setValue(approvedAt);
  sheet.getRange(foundRow, 17).setValue(approvedBy);

  const request = rowToRequest(sheet.getRange(foundRow, 1, 1, HEADERS.length).getValues()[0]);
  const emailResult = sendApprovalEmail(request);
  const emailStatusText = emailResult.ok ? "SENT" : "FAILED: " + emailResult.error;
  sheet.getRange(foundRow, 18).setValue(emailStatusText);

  return jsonResponse({
    ok: true,
    requestId: requestId,
    emailStatus: emailStatusText,
    message: "Request approved.",
  });
}

function readRequests() {
  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const requests = [];
  for (let i = 1; i < values.length; i += 1) {
    requests.push(rowToRequest(values[i]));
  }
  return requests;
}

function rowToRequest(row) {
  const request = {};
  for (let i = 0; i < HEADERS.length; i += 1) {
    request[HEADERS[i]] = row[i] === undefined ? "" : String(row[i]);
  }
  return request;
}

function sendApprovalEmail(request) {
  const recipient = safeText(request.email);
  if (!recipient) {
    return { ok: false, error: "Missing recipient email." };
  }

  const subject = "Vehicle request approved: " + request.requestId;
  const lines = [
    "Your vehicle request has been approved.",
    "",
    "Request ID: " + request.requestId,
    "Employee: " + request.employeeName,
    "Department: " + request.department,
    "Trip date: " + request.tripDate,
    "Time: " + request.startTime + " - " + request.endTime,
    "Pickup: " + request.pickup,
    "Destination: " + request.destination,
    "Passengers: " + request.passengers,
    "Purpose: " + request.purpose,
    "Approved at: " + request.approvedAt,
    "Approved by: " + request.approvedBy,
    "",
    "This is an automated email from " + CONFIG.APP_NAME + ".",
  ];

  try {
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      body: lines.join("\n"),
      name: "Scotch Fleet Service",
    });
    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
  }

  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeaders = HEADERS.every(function (header, idx) {
    return String(firstRow[idx] || "") === header;
  });
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function parsePayload(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  throw new Error("Missing payload.");
}

function verifyAdminKey(inputKey) {
  if (safeText(inputKey) !== CONFIG.ADMIN_KEY) {
    throw new Error("Invalid admin key.");
  }
}

function validateRequired(data, requiredFields) {
  requiredFields.forEach(function (field) {
    if (!safeText(data[field])) {
      throw new Error("Missing required field: " + field);
    }
  });
}

function safeText(value) {
  return String(value == null ? "" : value).trim();
}

function nowIso() {
  return Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
