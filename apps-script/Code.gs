const CONFIG = {
  SHEET_NAME: "Requests",
  ADMIN_KEY: "CHANGE_ME_TO_A_STRONG_ADMIN_KEY",
  TIME_ZONE: "Asia/Bangkok",
  APP_NAME: "Scotch webapps1",
};

const TOPIC_OFF_CYCLE = "OFF_CYCLE_DELIVERY";
const TOPIC_FACTORY = "FACTORY_CAR";
const CASE_FACTORY_DELIVERY = "FACTORY_DELIVERY_DOCS";
const CASE_FACTORY_SHUTTLE = "FACTORY_STAFF_SHUTTLE";

const HEADERS = [
  "requestId",
  "submittedAt",
  "requestTopic",
  "requestTopicLabel",
  "factoryCaseType",
  "factoryCaseLabel",
  "requesterName",
  "requesterPhone",
  "contactEmail",
  "offcycleStoreName",
  "offcycleDeliveryDate",
  "offcycleCrates",
  "offcycleAmount",
  "factoryJobName",
  "factoryDeliveryDate",
  "factoryDeliveryTime",
  "factoryDeliveryLocation",
  "factoryDeliveryMapUrl",
  "factoryReceiverPhone",
  "factoryTempDocumentRef",
  "factoryReservationRef",
  "shuttlePassengers",
  "shuttleTravelDate",
  "shuttleTravelTime",
  "shuttleReturnWait",
  "shuttleLocation",
  "shuttleMapUrl",
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
  const clean = sanitizeData(data);
  validateSubmission(clean);

  const sheet = getOrCreateSheet();
  const requestId = "REQ-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  const submittedAt = nowIso();

  const record = {
    requestId: requestId,
    submittedAt: submittedAt,
    requestTopic: clean.requestTopic,
    requestTopicLabel: clean.requestTopicLabel,
    factoryCaseType: clean.factoryCaseType,
    factoryCaseLabel: clean.factoryCaseLabel,
    requesterName: clean.requesterName,
    requesterPhone: clean.requesterPhone,
    contactEmail: clean.contactEmail,
    offcycleStoreName: clean.offcycleStoreName,
    offcycleDeliveryDate: clean.offcycleDeliveryDate,
    offcycleCrates: clean.offcycleCrates,
    offcycleAmount: clean.offcycleAmount,
    factoryJobName: clean.factoryJobName,
    factoryDeliveryDate: clean.factoryDeliveryDate,
    factoryDeliveryTime: clean.factoryDeliveryTime,
    factoryDeliveryLocation: clean.factoryDeliveryLocation,
    factoryDeliveryMapUrl: clean.factoryDeliveryMapUrl,
    factoryReceiverPhone: clean.factoryReceiverPhone,
    factoryTempDocumentRef: clean.factoryTempDocumentRef,
    factoryReservationRef: clean.factoryReservationRef,
    shuttlePassengers: clean.shuttlePassengers,
    shuttleTravelDate: clean.shuttleTravelDate,
    shuttleTravelTime: clean.shuttleTravelTime,
    shuttleReturnWait: clean.shuttleReturnWait,
    shuttleLocation: clean.shuttleLocation,
    shuttleMapUrl: clean.shuttleMapUrl,
    status: "PENDING",
    approvedAt: "",
    approvedBy: "",
    emailStatus: "",
  };

  const row = HEADERS.map(function (header) {
    return record[header] || "";
  });
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
    if (String(values[i][headerIndex("requestId")]) === requestId) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) {
    throw new Error("Request not found: " + requestId);
  }

  const statusCol = headerIndex("status") + 1;
  const approvedAtCol = headerIndex("approvedAt") + 1;
  const approvedByCol = headerIndex("approvedBy") + 1;
  const emailStatusCol = headerIndex("emailStatus") + 1;

  const statusCell = sheet.getRange(foundRow, statusCol);
  const currentStatus = String(statusCell.getValue() || "");
  if (currentStatus.toUpperCase() === "APPROVED") {
    return jsonResponse({
      ok: true,
      message: "Request already approved.",
      emailStatus: sheet.getRange(foundRow, emailStatusCol).getValue(),
    });
  }

  const approvedAt = nowIso();
  statusCell.setValue("APPROVED");
  sheet.getRange(foundRow, approvedAtCol).setValue(approvedAt);
  sheet.getRange(foundRow, approvedByCol).setValue(approvedBy);

  const requestRow = sheet.getRange(foundRow, 1, 1, HEADERS.length).getValues()[0];
  const request = rowToRequest(requestRow);
  request.approvedAt = approvedAt;
  request.approvedBy = approvedBy;
  const emailResult = sendApprovalEmail(request);
  const emailStatusText = emailResult.ok ? "SENT" : "FAILED: " + emailResult.error;
  sheet.getRange(foundRow, emailStatusCol).setValue(emailStatusText);

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
  const recipient = safeText(request.contactEmail);
  if (!recipient) {
    return { ok: false, error: "Missing recipient email." };
  }

  const subject = "อนุมัติคำขอใช้รถแล้ว: " + request.requestId;
  const lines = [
    "คำขอใช้รถของคุณได้รับการอนุมัติแล้ว",
    "",
    "Request ID: " + request.requestId,
    "หัวข้อ: " + request.requestTopicLabel,
  ];

  if (request.requestTopic === TOPIC_OFF_CYCLE) {
    lines.push("ชื่อห้าง: " + request.offcycleStoreName);
    lines.push("วันที่ต้องการจัดส่ง: " + request.offcycleDeliveryDate);
    lines.push("จำนวนลังสินค้า: " + request.offcycleCrates);
    lines.push("ยอดเงิน: " + request.offcycleAmount);
  }

  if (request.requestTopic === TOPIC_FACTORY) {
    lines.push("กรณี: " + request.factoryCaseLabel);
    if (request.factoryCaseType === CASE_FACTORY_DELIVERY) {
      lines.push("ชื่องาน: " + request.factoryJobName);
      lines.push("วันที่จัดส่ง: " + request.factoryDeliveryDate);
      lines.push("เวลาจัดส่ง: " + request.factoryDeliveryTime);
      lines.push("สถานที่จัดส่ง: " + request.factoryDeliveryLocation);
      lines.push("Google Map: " + request.factoryDeliveryMapUrl);
      lines.push("เบอร์ติดต่อหน้างาน: " + request.factoryReceiverPhone);
      lines.push("เอกสารชั่วคราว: " + request.factoryTempDocumentRef);
      lines.push("ข้อมูลเบิกของ/Reservation: " + request.factoryReservationRef);
    }
    if (request.factoryCaseType === CASE_FACTORY_SHUTTLE) {
      lines.push("จำนวนผู้โดยสาร: " + request.shuttlePassengers);
      lines.push("วันที่เดินทาง: " + request.shuttleTravelDate);
      lines.push("เวลาเดินทาง: " + request.shuttleTravelTime);
      lines.push("รอรับกลับ: " + shuttleWaitLabel(request.shuttleReturnWait));
      lines.push("สถานที่: " + request.shuttleLocation);
      lines.push("Google Map: " + request.shuttleMapUrl);
    }
  }

  lines.push("ผู้ขอใช้: " + request.requesterName);
  lines.push("เบอร์ติดต่อผู้ขอ: " + request.requesterPhone);
  lines.push("อีเมลติดต่อ: " + request.contactEmail);
  lines.push("อนุมัติเมื่อ: " + request.approvedAt);
  lines.push("อนุมัติโดย: " + request.approvedBy);
  lines.push("");
  lines.push("This is an automated email from " + CONFIG.APP_NAME + ".");

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

function shuttleWaitLabel(value) {
  if (value === "WAIT_RETURN") return "รอรับกลับ";
  if (value === "NO_RETURN") return "ไม่ต้องรอรับกลับ";
  return value;
}

function sanitizeData(data) {
  const output = {};
  Object.keys(data || {}).forEach(function (key) {
    output[key] = safeText(data[key]);
  });
  return output;
}

function validateSubmission(data) {
  const allowedTopics = [TOPIC_OFF_CYCLE, TOPIC_FACTORY];
  if (allowedTopics.indexOf(data.requestTopic) === -1) {
    throw new Error("Invalid request topic.");
  }

  requireField(data, "requesterName");
  requireField(data, "requesterPhone");
  requireField(data, "contactEmail");
  if (!isValidPhone(data.requesterPhone)) {
    throw new Error("Invalid requesterPhone.");
  }
  if (!isValidEmail(data.contactEmail)) {
    throw new Error("Invalid contactEmail.");
  }
  if (data.policyAgree !== "YES") {
    throw new Error("Policy agreement is required.");
  }

  if (data.requestTopic === TOPIC_OFF_CYCLE) {
    requireField(data, "offcycleStoreName");
    requireField(data, "offcycleDeliveryDate");
    requireField(data, "offcycleCrates");
    requireField(data, "offcycleAmount");
    if (!isIntAtLeast(data.offcycleCrates, 1)) {
      throw new Error("offcycleCrates must be >= 1.");
    }
    if (!isNumberAtLeast(data.offcycleAmount, 0)) {
      throw new Error("offcycleAmount must be >= 0.");
    }
  }

  if (data.requestTopic === TOPIC_FACTORY) {
    requireField(data, "factoryCaseType");
    if (data.factoryCaseType !== CASE_FACTORY_DELIVERY && data.factoryCaseType !== CASE_FACTORY_SHUTTLE) {
      throw new Error("Invalid factoryCaseType.");
    }

    if (data.factoryCaseType === CASE_FACTORY_DELIVERY) {
      requireField(data, "factoryJobName");
      requireField(data, "factoryDeliveryDate");
      requireField(data, "factoryDeliveryTime");
      requireField(data, "factoryDeliveryLocation");
      requireField(data, "factoryDeliveryMapUrl");
      requireField(data, "factoryReceiverPhone");
      requireField(data, "factoryTempDocumentRef");
      requireField(data, "factoryReservationRef");
      if (!isValidUrl(data.factoryDeliveryMapUrl)) {
        throw new Error("Invalid factoryDeliveryMapUrl.");
      }
      if (!isValidPhone(data.factoryReceiverPhone)) {
        throw new Error("Invalid factoryReceiverPhone.");
      }
    }

    if (data.factoryCaseType === CASE_FACTORY_SHUTTLE) {
      requireField(data, "shuttlePassengers");
      requireField(data, "shuttleTravelDate");
      requireField(data, "shuttleTravelTime");
      requireField(data, "shuttleReturnWait");
      requireField(data, "shuttleLocation");
      requireField(data, "shuttleMapUrl");
      if (!isIntAtLeast(data.shuttlePassengers, 6)) {
        throw new Error("shuttlePassengers must be >= 6.");
      }
      if (data.shuttleReturnWait !== "WAIT_RETURN" && data.shuttleReturnWait !== "NO_RETURN") {
        throw new Error("Invalid shuttleReturnWait.");
      }
      if (!isValidUrl(data.shuttleMapUrl)) {
        throw new Error("Invalid shuttleMapUrl.");
      }
    }
  }
}

function requireField(data, field) {
  if (!safeText(data[field])) {
    throw new Error("Missing required field: " + field);
  }
}

function isValidPhone(phone) {
  const digits = safeText(phone).replace(/\D/g, "");
  return digits.length >= 8 && digits.length <= 15;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(safeText(email));
}

function isIntAtLeast(value, min) {
  const num = Number(value);
  return Number.isInteger(num) && num >= min;
}

function isNumberAtLeast(value, min) {
  const num = Number(value);
  return !isNaN(num) && num >= min;
}

function isValidUrl(url) {
  const text = safeText(url);
  return /^https?:\/\/.+/i.test(text);
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

function headerIndex(field) {
  const index = HEADERS.indexOf(field);
  if (index === -1) {
    throw new Error("Header not found: " + field);
  }
  return index;
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
