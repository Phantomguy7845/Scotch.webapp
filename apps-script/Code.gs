const CONFIG = {
  APP_NAME: "Scotch webapps1",
  TIME_ZONE: "Asia/Bangkok",
  SHEET_NAME: "Requests",
  SPREADSHEET_ID: "10b7BtWAWiyZ-AqYFmwuzQQLPsyjDOEr642Ml6dJ-Mx8",
  ADMIN_KEY: "CHANGE_ME_TO_A_STRONG_ADMIN_KEY",
  UPLOAD_FOLDER_ID: "1Bm1WDqIHaxMLzBmzU5iyRvxSlLL99PGS",
  UPLOAD_FOLDER_NAME: "ScotchWebapps1_Uploads",
  MAKE_UPLOAD_PUBLIC: true,
  MAX_UPLOAD_BYTES: 3 * 1024 * 1024,
};

const TOPIC_OFF_CYCLE = "OFF_CYCLE_DELIVERY";
const TOPIC_FACTORY = "FACTORY_CAR";
const CASE_FACTORY_DELIVERY = "FACTORY_DELIVERY_DOCS";
const CASE_FACTORY_SHUTTLE = "FACTORY_STAFF_SHUTTLE";
const ALLOWED_IMAGE_MIME_TYPES = ["image/jpeg", "image/png", "image/webp"];

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
  "factoryTempDocumentImageName",
  "factoryTempDocumentImageUrl",
  "factoryTempDocumentImageDownloadUrl",
  "factoryTempDocumentImageFileId",
  "factoryReservationImageName",
  "factoryReservationImageUrl",
  "factoryReservationImageDownloadUrl",
  "factoryReservationImageFileId",
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
    return jsonResponse({ ok: false, error: normalizeErrorMessage(error) });
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
    return jsonResponse({ ok: false, error: normalizeErrorMessage(error) });
  }
}

function submitRequest(data) {
  const clean = sanitizeTextData(data);
  const attachments = sanitizeAttachments(data);
  validateSubmission(clean, attachments);

  const sheet = getOrCreateSheet();
  const requestId = "REQ-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  const submittedAt = nowIso();

  let tempDocUpload = emptyUploadMeta();
  let reservationUpload = emptyUploadMeta();

  if (clean.requestTopic === TOPIC_FACTORY && clean.factoryCaseType === CASE_FACTORY_DELIVERY) {
    tempDocUpload = saveAttachmentToDrive(attachments.factoryTempDocumentImage, requestId, "temp_doc");
    reservationUpload = saveAttachmentToDrive(attachments.factoryReservationImage, requestId, "reservation");
  }

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
    factoryTempDocumentImageName: tempDocUpload.name,
    factoryTempDocumentImageUrl: tempDocUpload.url,
    factoryTempDocumentImageDownloadUrl: tempDocUpload.downloadUrl,
    factoryTempDocumentImageFileId: tempDocUpload.fileId,
    factoryReservationImageName: reservationUpload.name,
    factoryReservationImageUrl: reservationUpload.url,
    factoryReservationImageDownloadUrl: reservationUpload.downloadUrl,
    factoryReservationImageFileId: reservationUpload.fileId,
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
  if (!requestId) throw new Error("requestId is required.");

  const approvedBy = safeText(payload.approvedBy || "Fleet Admin");
  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) throw new Error("No requests found.");

  let foundRow = -1;
  for (let i = 1; i < values.length; i += 1) {
    if (String(values[i][headerIndex("requestId")]) === requestId) {
      foundRow = i + 1;
      break;
    }
  }
  if (foundRow === -1) throw new Error("Request not found: " + requestId);

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
  if (!recipient) return { ok: false, error: "Missing recipient email." };

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
      lines.push("รายละเอียดเอกสารชั่วคราว: " + request.factoryTempDocumentRef);
      lines.push("รายละเอียด Reservation: " + request.factoryReservationRef);
      lines.push("ลิงก์รูปเอกสารชั่วคราว: " + request.factoryTempDocumentImageUrl);
      lines.push("ลิงก์ดาวน์โหลดเอกสารชั่วคราว: " + request.factoryTempDocumentImageDownloadUrl);
      lines.push("ลิงก์รูป Reservation: " + request.factoryReservationImageUrl);
      lines.push("ลิงก์ดาวน์โหลดรูป Reservation: " + request.factoryReservationImageDownloadUrl);
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

function sanitizeTextData(data) {
  return {
    requestTopic: safeText(data.requestTopic),
    requestTopicLabel: safeText(data.requestTopicLabel),
    factoryCaseType: safeText(data.factoryCaseType),
    factoryCaseLabel: safeText(data.factoryCaseLabel),
    requesterName: safeText(data.requesterName),
    requesterPhone: safeText(data.requesterPhone),
    contactEmail: safeText(data.contactEmail),
    offcycleStoreName: safeText(data.offcycleStoreName),
    offcycleDeliveryDate: safeText(data.offcycleDeliveryDate),
    offcycleCrates: safeText(data.offcycleCrates),
    offcycleAmount: safeText(data.offcycleAmount),
    factoryJobName: safeText(data.factoryJobName),
    factoryDeliveryDate: safeText(data.factoryDeliveryDate),
    factoryDeliveryTime: safeText(data.factoryDeliveryTime),
    factoryDeliveryLocation: safeText(data.factoryDeliveryLocation),
    factoryDeliveryMapUrl: safeText(data.factoryDeliveryMapUrl),
    factoryReceiverPhone: safeText(data.factoryReceiverPhone),
    factoryTempDocumentRef: safeText(data.factoryTempDocumentRef),
    factoryReservationRef: safeText(data.factoryReservationRef),
    shuttlePassengers: safeText(data.shuttlePassengers),
    shuttleTravelDate: safeText(data.shuttleTravelDate),
    shuttleTravelTime: safeText(data.shuttleTravelTime),
    shuttleReturnWait: safeText(data.shuttleReturnWait),
    shuttleLocation: safeText(data.shuttleLocation),
    shuttleMapUrl: safeText(data.shuttleMapUrl),
    policyAgree: safeText(data.policyAgree),
  };
}

function sanitizeAttachments(data) {
  return {
    factoryTempDocumentImage: normalizeAttachment(data.factoryTempDocumentImage),
    factoryReservationImage: normalizeAttachment(data.factoryReservationImage),
  };
}

function normalizeAttachment(rawAttachment) {
  if (!rawAttachment || typeof rawAttachment !== "object") return emptyAttachment();
  return {
    name: safeText(rawAttachment.name),
    mimeType: safeText(rawAttachment.mimeType).toLowerCase(),
    base64: safeBase64(rawAttachment.base64),
    size: Number(rawAttachment.size || 0),
  };
}

function emptyAttachment() {
  return { name: "", mimeType: "", base64: "", size: 0 };
}

function emptyUploadMeta() {
  return { name: "", url: "", downloadUrl: "", fileId: "" };
}

function safeBase64(value) {
  const text = safeText(value);
  if (!text) return "";
  const commaIndex = text.indexOf(",");
  if (commaIndex > -1) return safeText(text.slice(commaIndex + 1));
  return text;
}

function validateSubmission(data, attachments) {
  const allowedTopics = [TOPIC_OFF_CYCLE, TOPIC_FACTORY];
  if (allowedTopics.indexOf(data.requestTopic) === -1) throw new Error("Invalid request topic.");

  requireField(data, "requesterName");
  requireField(data, "requesterPhone");
  requireField(data, "contactEmail");
  if (!isValidPhone(data.requesterPhone)) throw new Error("Invalid requesterPhone.");
  if (!isValidEmail(data.contactEmail)) throw new Error("Invalid contactEmail.");
  if (data.policyAgree !== "YES") throw new Error("Policy agreement is required.");

  if (data.requestTopic === TOPIC_OFF_CYCLE) {
    requireField(data, "offcycleStoreName");
    requireField(data, "offcycleDeliveryDate");
    requireField(data, "offcycleCrates");
    requireField(data, "offcycleAmount");
    if (!isIntAtLeast(data.offcycleCrates, 1)) throw new Error("offcycleCrates must be >= 1.");
    if (!isNumberAtLeast(data.offcycleAmount, 0)) throw new Error("offcycleAmount must be >= 0.");
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
      if (!isValidUrl(data.factoryDeliveryMapUrl)) throw new Error("Invalid factoryDeliveryMapUrl.");
      if (!isValidPhone(data.factoryReceiverPhone)) throw new Error("Invalid factoryReceiverPhone.");
      assertValidAttachment(attachments.factoryTempDocumentImage, "factoryTempDocumentImage");
      assertValidAttachment(attachments.factoryReservationImage, "factoryReservationImage");
    }

    if (data.factoryCaseType === CASE_FACTORY_SHUTTLE) {
      requireField(data, "shuttlePassengers");
      requireField(data, "shuttleTravelDate");
      requireField(data, "shuttleTravelTime");
      requireField(data, "shuttleReturnWait");
      requireField(data, "shuttleLocation");
      requireField(data, "shuttleMapUrl");
      if (!isIntAtLeast(data.shuttlePassengers, 6)) throw new Error("shuttlePassengers must be >= 6.");
      if (data.shuttleReturnWait !== "WAIT_RETURN" && data.shuttleReturnWait !== "NO_RETURN") {
        throw new Error("Invalid shuttleReturnWait.");
      }
      if (!isValidUrl(data.shuttleMapUrl)) throw new Error("Invalid shuttleMapUrl.");
    }
  }
}

function assertValidAttachment(attachment, fieldName) {
  if (!attachment || !attachment.base64) throw new Error("Missing required attachment: " + fieldName);
  if (ALLOWED_IMAGE_MIME_TYPES.indexOf(attachment.mimeType) === -1) {
    throw new Error("Invalid attachment mime type: " + fieldName);
  }
  const decodedBytes = estimateDecodedByteSize(attachment.base64);
  if (decodedBytes <= 0) throw new Error("Invalid attachment content: " + fieldName);
  if (decodedBytes > CONFIG.MAX_UPLOAD_BYTES) throw new Error("Attachment is too large: " + fieldName);
}

function estimateDecodedByteSize(base64) {
  const value = safeText(base64);
  if (!value) return 0;
  const paddingMatches = value.match(/=*$/);
  const paddingCount = paddingMatches ? paddingMatches[0].length : 0;
  return Math.floor((value.length * 3) / 4) - paddingCount;
}

function saveAttachmentToDrive(attachment, requestId, label) {
  const bytes = Utilities.base64Decode(attachment.base64);
  if (!bytes || !bytes.length) throw new Error("Attachment is empty: " + label);
  if (bytes.length > CONFIG.MAX_UPLOAD_BYTES) throw new Error("Attachment exceeds limit: " + label);

  const fileName = buildAttachmentFileName(attachment.name, requestId, label, attachment.mimeType);
  const blob = Utilities.newBlob(bytes, attachment.mimeType, fileName);
  const folder = getUploadFolder();
  const file = folder.createFile(blob);

  if (CONFIG.MAKE_UPLOAD_PUBLIC) {
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (error) {
      // Keep private when org policy blocks public sharing.
    }
  }

  return {
    name: file.getName(),
    url: file.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId(),
    fileId: file.getId(),
  };
}

function buildAttachmentFileName(originalName, requestId, label, mimeType) {
  const rawName = sanitizeFileName(safeText(originalName));
  const ext = rawName && rawName.indexOf(".") > -1 ? "" : "." + extensionFromMimeType(mimeType);
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, "yyyyMMdd_HHmmss");
  const baseName = rawName || label + ext;
  return requestId + "_" + label + "_" + timestamp + "_" + baseName;
}

function extensionFromMimeType(mimeType) {
  if (mimeType === "image/jpeg") return "jpg";
  if (mimeType === "image/png") return "png";
  if (mimeType === "image/webp") return "webp";
  return "bin";
}

function sanitizeFileName(fileName) {
  return safeText(fileName).replace(/[^\w.\-]+/g, "_").slice(0, 120);
}

function getUploadFolder() {
  if (safeText(CONFIG.UPLOAD_FOLDER_ID)) {
    return DriveApp.getFolderById(CONFIG.UPLOAD_FOLDER_ID);
  }
  const folders = DriveApp.getFoldersByName(CONFIG.UPLOAD_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.UPLOAD_FOLDER_NAME);
}

function getSpreadsheet() {
  if (safeText(CONFIG.SPREADSHEET_ID)) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) return activeSpreadsheet;
  throw new Error("Spreadsheet not found. Set CONFIG.SPREADSHEET_ID.");
}

function getOrCreateSheet() {
  const spreadsheet = getSpreadsheet();
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
  if (index === -1) throw new Error("Header not found: " + field);
  return index;
}

function parsePayload(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  if (e && e.postData && e.postData.contents) {
    const raw = safeText(e.postData.contents);
    if (raw) {
      if (raw.charAt(0) === "{") return JSON.parse(raw);
      if (raw.indexOf("payload=") === 0) {
        const decoded = decodeURIComponent(raw.replace(/^payload=/, ""));
        return JSON.parse(decoded);
      }
    }
  }
  throw new Error("Missing payload.");
}

function verifyAdminKey(inputKey) {
  if (safeText(inputKey) !== CONFIG.ADMIN_KEY) {
    throw new Error("Invalid admin key.");
  }
}

function requireField(data, field) {
  if (!safeText(data[field])) throw new Error("Missing required field: " + field);
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

function authorizeDriveAccess() {
  const spreadsheet = getSpreadsheet();
  const folder = getUploadFolder();
  return (
    "Authorization OK. Spreadsheet: " +
    spreadsheet.getId() +
    " | Folder: " +
    folder.getId()
  );
}

function normalizeErrorMessage(error) {
  const raw = safeText(error && error.message ? error.message : error);
  const isDriveAuthError =
    raw.indexOf("DriveApp.getFolderById") > -1 ||
    raw.indexOf("https://www.googleapis.com/auth/drive") > -1 ||
    raw.indexOf("authorization") > -1 && raw.indexOf("DriveApp") > -1;

  if (isDriveAuthError) {
    return (
      "Apps Script ยังไม่ได้อนุญาตสิทธิ์ Google Drive. " +
      "ให้ Run ฟังก์ชัน authorizeDriveAccess() 1 ครั้ง แล้ว Deploy Web App เวอร์ชันใหม่ (Execute as: Me)."
    );
  }

  return raw || "Unexpected error.";
}
