const CONFIG = {
  APP_NAME: "Scotch webapps1",
  TIME_ZONE: "Asia/Bangkok",
  SHEET_NAME: "Requests",
  SPREADSHEET_ID: "10b7BtWAWiyZ-AqYFmwuzQQLPsyjDOEr642Ml6dJ-Mx8",
  ADMIN_KEY: "CHANGE_ME_TO_A_STRONG_ADMIN_KEY",
  UPLOAD_FOLDER_ID: "1LDNViyPm75ByXeJAng-F99VESi-vqrQg",
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
  "offcycleHasPo",
  "offcyclePoImageName",
  "offcyclePoImageUrl",
  "offcyclePoImageDownloadUrl",
  "offcyclePoImageFileId",
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
  "rejectedAt",
  "rejectedBy",
  "decisionRemark",
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

    if (action === "rejectrequest") {
      return rejectRequest(payload);
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (error) {
    return jsonResponse({ ok: false, error: normalizeErrorMessage(error) });
  }
}

function submitRequest(data) {
  if (!data || typeof data !== "object" || Array.isArray(data) || Object.keys(data).length === 0) {
    throw new Error(
      "submitRequest requires request data from web form. Use deployed Web App URL instead of clicking Run on submitRequest().",
    );
  }

  const clean = sanitizeTextData(data);
  const attachments = sanitizeAttachments(data);
  validateSubmission(clean, attachments);

  const sheet = getOrCreateSheet();
  const requestId = "REQ-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  const submittedAt = nowIso();

  let tempDocUpload = emptyUploadMeta();
  let reservationUpload = emptyUploadMeta();
  let offcyclePoUpload = emptyUploadMeta();

  if (clean.requestTopic === TOPIC_OFF_CYCLE && clean.offcycleHasPo === "YES") {
    offcyclePoUpload = saveAttachmentToDrive(attachments.offcyclePoImage, requestId, "offcycle_po");
  }

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
    offcycleHasPo: clean.offcycleHasPo,
    offcyclePoImageName: offcyclePoUpload.name,
    offcyclePoImageUrl: offcyclePoUpload.url,
    offcyclePoImageDownloadUrl: offcyclePoUpload.downloadUrl,
    offcyclePoImageFileId: offcyclePoUpload.fileId,
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
    rejectedAt: "",
    rejectedBy: "",
    decisionRemark: "",
  };

  const row = HEADERS.map(function (header) {
    return record[header] || "";
  });
  sheet.appendRow(row);

  return jsonResponse({
    ok: true,
    requestId: requestId,
    submittedAt: submittedAt,
    message: "Request submitted successfully.",
  });
}

function approveRequest(payload) {
  return handleDecision(payload, "APPROVED");
}

function rejectRequest(payload) {
  return handleDecision(payload, "REJECTED");
}

function handleDecision(payload, nextStatus) {
  verifyAdminKey(payload.adminKey);
  const requestId = safeText(payload.requestId);
  if (!requestId) throw new Error("requestId is required.");

  const decidedBy = safeText(payload.approvedBy || payload.decidedBy || "Fleet Admin");
  const remark = safeText(payload.remark);
  if (!remark) throw new Error("remark is required.");

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
  const rejectedAtCol = headerIndex("rejectedAt") + 1;
  const rejectedByCol = headerIndex("rejectedBy") + 1;
  const decisionRemarkCol = headerIndex("decisionRemark") + 1;
  const emailStatusCol = headerIndex("emailStatus") + 1;

  const statusCell = sheet.getRange(foundRow, statusCol);
  const currentStatus = String(statusCell.getValue() || "").toUpperCase();

  if (currentStatus === nextStatus) {
    return jsonResponse({
      ok: true,
      message: nextStatus === "APPROVED" ? "Request already approved." : "Request already rejected.",
      status: currentStatus,
      approvedAt: String(sheet.getRange(foundRow, approvedAtCol).getValue() || ""),
      approvedBy: String(sheet.getRange(foundRow, approvedByCol).getValue() || ""),
      rejectedAt: String(sheet.getRange(foundRow, rejectedAtCol).getValue() || ""),
      rejectedBy: String(sheet.getRange(foundRow, rejectedByCol).getValue() || ""),
      decisionRemark: String(sheet.getRange(foundRow, decisionRemarkCol).getValue() || ""),
      emailStatus: String(sheet.getRange(foundRow, emailStatusCol).getValue() || ""),
    });
  }

  if (currentStatus === "APPROVED" || currentStatus === "REJECTED") {
    throw new Error("Request already decided: " + currentStatus + ".");
  }

  const decidedAt = nowIso();
  statusCell.setValue(nextStatus);
  sheet.getRange(foundRow, decisionRemarkCol).setValue(remark);

  if (nextStatus === "APPROVED") {
    sheet.getRange(foundRow, approvedAtCol).setValue(decidedAt);
    sheet.getRange(foundRow, approvedByCol).setValue(decidedBy);
    sheet.getRange(foundRow, rejectedAtCol).setValue("");
    sheet.getRange(foundRow, rejectedByCol).setValue("");
  } else {
    sheet.getRange(foundRow, rejectedAtCol).setValue(decidedAt);
    sheet.getRange(foundRow, rejectedByCol).setValue(decidedBy);
    sheet.getRange(foundRow, approvedAtCol).setValue("");
    sheet.getRange(foundRow, approvedByCol).setValue("");
  }

  const requestRow = sheet.getRange(foundRow, 1, 1, HEADERS.length).getValues()[0];
  const request = rowToRequest(requestRow);
  request.status = nextStatus;
  request.decisionRemark = remark;
  if (nextStatus === "APPROVED") {
    request.approvedAt = decidedAt;
    request.approvedBy = decidedBy;
    request.rejectedAt = "";
    request.rejectedBy = "";
  } else {
    request.rejectedAt = decidedAt;
    request.rejectedBy = decidedBy;
    request.approvedAt = "";
    request.approvedBy = "";
  }

  const emailResult = sendDecisionEmail(request);
  const emailStatusText = emailResult.ok
    ? emailResult.warning
      ? "SENT_WITH_WARNING: " + emailResult.warning
      : "SENT"
    : "FAILED: " + emailResult.error;
  sheet.getRange(foundRow, emailStatusCol).setValue(emailStatusText);

  return jsonResponse({
    ok: true,
    requestId: requestId,
    status: nextStatus,
    approvedAt: request.approvedAt,
    approvedBy: request.approvedBy,
    rejectedAt: request.rejectedAt,
    rejectedBy: request.rejectedBy,
    decisionRemark: remark,
    emailStatus: emailStatusText,
    message: nextStatus === "APPROVED" ? "Request approved." : "Request rejected.",
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

function sendDecisionEmail(request) {
  const recipient = safeText(request.contactEmail);
  if (!recipient) return { ok: false, error: "Missing recipient email." };

  const isApproved = request.status === "APPROVED";
  const decisionLabel = isApproved ? "อนุมัติ" : "ไม่อนุมัติ";
  const decisionBadge = isApproved ? "APPROVED" : "REJECTED";
  const decisionColor = isApproved ? "#0E7A42" : "#B42318";
  const decidedAt = safeText(isApproved ? request.approvedAt : request.rejectedAt);
  const decidedBy = safeText(isApproved ? request.approvedBy : request.rejectedBy);
  const decisionRemark = safeText(request.decisionRemark);

  const subject =
    "[Scotch Industrial Fleet] ผลการพิจารณาคำขอ " + safeText(request.requestId) + " : " + decisionLabel;
  const details = buildDecisionDetailItems(request);
  const detailText = details.map(function (item) {
    return item.label + ": " + item.value;
  });

  const textLines = [
    "ผลการพิจารณาคำขอใช้รถบริษัท",
    "",
    "สถานะ: " + decisionLabel + " (" + decisionBadge + ")",
    "Request ID: " + safeText(request.requestId),
    "พิจารณาเมื่อ: " + decidedAt,
    "พิจารณาโดย: " + decidedBy,
    "หมายเหตุ: " + (decisionRemark || "-"),
    "",
    "รายละเอียดคำขอ",
  ]
    .concat(detailText)
    .concat([
      "",
      "อีเมลฉบับนี้ถูกส่งโดยระบบอัตโนมัติ กรุณาอย่าตอบกลับ (no-reply).",
      "Scotch Industrial Fleet Service",
    ]);

  const detailRowsHtml = details
    .map(function (item) {
      return (
        '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #E5E7EB;color:#4B5563;width:32%;font-weight:600;">' +
        escapeHtml(item.label) +
        "</td>" +
        '<td style="padding:8px 10px;border-bottom:1px solid #E5E7EB;color:#111827;">' +
        formatDecisionValueHtml(item.value) +
        "</td>" +
        "</tr>"
      );
    })
    .join("");

  const htmlBody =
    '<div style="margin:0;padding:24px;background:#F4F6FB;font-family:Arial,sans-serif;color:#111827;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="max-width:760px;margin:0 auto;background:#FFFFFF;border-radius:12px;border:1px solid #E5E7EB;overflow:hidden;">' +
    "<tr>" +
    '<td style="padding:16px 20px;background:' +
    decisionColor +
    ';color:#FFFFFF;">' +
    '<div style="font-size:12px;letter-spacing:.04em;text-transform:uppercase;opacity:.9;">Scotch Industrial Fleet Service</div>' +
    '<div style="margin-top:6px;font-size:22px;font-weight:700;">แจ้งผลการพิจารณาคำขอใช้รถ</div>' +
    '<div style="margin-top:8px;font-size:14px;">สถานะ: <strong>' +
    decisionLabel +
    " (" +
    decisionBadge +
    ")</strong></div>" +
    "</td>" +
    "</tr>" +
    "<tr>" +
    '<td style="padding:20px 20px 8px 20px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;border:1px solid #E5E7EB;border-radius:10px;overflow:hidden;">' +
    '<tr><td style="padding:8px 10px;border-bottom:1px solid #E5E7EB;color:#4B5563;width:32%;font-weight:600;">Request ID</td><td style="padding:8px 10px;border-bottom:1px solid #E5E7EB;">' +
    escapeHtml(request.requestId) +
    "</td></tr>" +
    '<tr><td style="padding:8px 10px;border-bottom:1px solid #E5E7EB;color:#4B5563;width:32%;font-weight:600;">พิจารณาเมื่อ</td><td style="padding:8px 10px;border-bottom:1px solid #E5E7EB;">' +
    escapeHtml(decidedAt || "-") +
    "</td></tr>" +
    '<tr><td style="padding:8px 10px;color:#4B5563;width:32%;font-weight:600;">พิจารณาโดย</td><td style="padding:8px 10px;">' +
    escapeHtml(decidedBy || "-") +
    "</td></tr>" +
    "</table>" +
    "</td>" +
    "</tr>" +
    "<tr>" +
    '<td style="padding:8px 20px 12px 20px;">' +
    '<div style="border:1px solid #E5E7EB;border-radius:10px;padding:12px;background:#FAFAFA;">' +
    '<div style="font-size:13px;color:#4B5563;font-weight:600;">Remark จากผู้พิจารณา</div>' +
    '<div style="margin-top:6px;font-size:14px;line-height:1.6;color:#111827;white-space:pre-wrap;">' +
    escapeHtml(decisionRemark || "-") +
    "</div>" +
    "</div>" +
    "</td>" +
    "</tr>" +
    "<tr>" +
    '<td style="padding:0 20px 12px 20px;">' +
    '<div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:8px;">รายละเอียดคำขอ</div>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;border:1px solid #E5E7EB;border-radius:10px;overflow:hidden;">' +
    detailRowsHtml +
    "</table>" +
    "</td>" +
    "</tr>" +
    "<tr>" +
    '<td style="padding:14px 20px 20px 20px;font-size:12px;color:#6B7280;line-height:1.6;">' +
    "อีเมลฉบับนี้ส่งจากระบบอัตโนมัติ กรุณาอย่าตอบกลับ (no-reply).<br/>" +
    "This message was generated automatically by " +
    escapeHtml(CONFIG.APP_NAME) +
    "." +
    "</td>" +
    "</tr>" +
    "</table>" +
    "</div>";

  const emailOptions = {
    to: recipient,
    subject: subject,
    body: textLines.join("\n"),
    htmlBody: htmlBody,
    name: "Scotch Industrial Fleet Notification",
  };
  if (canUseNoReply()) {
    emailOptions.noReply = true;
  }

  try {
    MailApp.sendEmail(emailOptions);
    return { ok: true };
  } catch (error) {
    if (emailOptions.noReply && isNoReplyUnsupportedError(error)) {
      delete emailOptions.noReply;
      MailApp.sendEmail(emailOptions);
      return { ok: true, warning: "NO_REPLY_UNSUPPORTED_FOR_ACCOUNT" };
    }
    return { ok: false, error: error.message };
  }
}

function buildDecisionDetailItems(request) {
  const items = [];
  pushDecisionItem(items, "หัวข้อ", request.requestTopicLabel || request.requestTopic);
  pushDecisionItem(items, "กรณี", request.factoryCaseLabel);
  pushDecisionItem(items, "ผู้ขอใช้", request.requesterName);
  pushDecisionItem(items, "เบอร์ติดต่อผู้ขอ", request.requesterPhone);
  pushDecisionItem(items, "อีเมลติดต่อ", request.contactEmail);
  pushDecisionItem(items, "ส่งคำขอเมื่อ", request.submittedAt);

  if (request.requestTopic === TOPIC_OFF_CYCLE) {
    pushDecisionItem(items, "ชื่อห้าง", request.offcycleStoreName);
    pushDecisionItem(items, "วันที่ต้องการจัดส่ง", request.offcycleDeliveryDate);
    pushDecisionItem(items, "จำนวนลังสินค้า", request.offcycleCrates);
    pushDecisionItem(items, "ยอดเงิน", request.offcycleAmount);
    pushDecisionItem(items, "มี PO", request.offcycleHasPo === "YES" ? "มี" : "ไม่มี");
    if (request.offcycleHasPo === "YES") {
      pushDecisionItem(items, "ลิงก์รูป PO", request.offcyclePoImageUrl);
      pushDecisionItem(items, "ลิงก์ดาวน์โหลดรูป PO", request.offcyclePoImageDownloadUrl);
    }
  }

  if (request.requestTopic === TOPIC_FACTORY) {
    if (request.factoryCaseType === CASE_FACTORY_DELIVERY) {
      pushDecisionItem(items, "ชื่องาน", request.factoryJobName);
      pushDecisionItem(items, "วันที่จัดส่ง", request.factoryDeliveryDate);
      pushDecisionItem(items, "เวลาจัดส่ง", request.factoryDeliveryTime);
      pushDecisionItem(items, "สถานที่จัดส่ง", request.factoryDeliveryLocation);
      pushDecisionItem(items, "Google Map", request.factoryDeliveryMapUrl);
      pushDecisionItem(items, "เบอร์ติดต่อหน้างาน", request.factoryReceiverPhone);
      pushDecisionItem(items, "รายละเอียดเอกสารชั่วคราว", request.factoryTempDocumentRef);
      pushDecisionItem(items, "รายละเอียด Reservation", request.factoryReservationRef);
      pushDecisionItem(items, "ลิงก์รูปเอกสารชั่วคราว", request.factoryTempDocumentImageUrl);
      pushDecisionItem(
        items,
        "ลิงก์ดาวน์โหลดเอกสารชั่วคราว",
        request.factoryTempDocumentImageDownloadUrl,
      );
      pushDecisionItem(items, "ลิงก์รูป Reservation", request.factoryReservationImageUrl);
      pushDecisionItem(items, "ลิงก์ดาวน์โหลดรูป Reservation", request.factoryReservationImageDownloadUrl);
    }
    if (request.factoryCaseType === CASE_FACTORY_SHUTTLE) {
      pushDecisionItem(items, "จำนวนผู้โดยสาร", request.shuttlePassengers);
      pushDecisionItem(items, "วันที่เดินทาง", request.shuttleTravelDate);
      pushDecisionItem(items, "เวลาเดินทาง", request.shuttleTravelTime);
      pushDecisionItem(items, "รอรับกลับ", shuttleWaitLabel(request.shuttleReturnWait));
      pushDecisionItem(items, "สถานที่", request.shuttleLocation);
      pushDecisionItem(items, "Google Map", request.shuttleMapUrl);
    }
  }

  return items;
}

function pushDecisionItem(items, label, value) {
  const text = safeText(value);
  if (!text) return;
  items.push({ label: label, value: text });
}

function formatDecisionValueHtml(value) {
  const text = safeText(value);
  if (!text) return "-";
  if (isValidUrl(text)) {
    const escapedUrl = escapeHtml(text);
    return '<a href="' + escapedUrl + '" target="_blank" rel="noopener noreferrer">' + escapedUrl + "</a>";
  }
  return escapeHtml(text);
}

function shuttleWaitLabel(value) {
  if (value === "WAIT_RETURN") return "รอรับกลับ";
  if (value === "NO_RETURN") return "ไม่ต้องรอรับกลับ";
  return value;
}

function sanitizeTextData(data) {
  const source = data && typeof data === "object" ? data : {};
  return {
    requestTopic: safeText(source.requestTopic),
    requestTopicLabel: safeText(source.requestTopicLabel),
    factoryCaseType: safeText(source.factoryCaseType),
    factoryCaseLabel: safeText(source.factoryCaseLabel),
    requesterName: safeText(source.requesterName),
    requesterPhone: safeText(source.requesterPhone),
    contactEmail: safeText(source.contactEmail),
    offcycleStoreName: safeText(source.offcycleStoreName),
    offcycleDeliveryDate: safeText(source.offcycleDeliveryDate),
    offcycleCrates: safeText(source.offcycleCrates),
    offcycleAmount: safeText(source.offcycleAmount),
    offcycleHasPo: safeText(source.offcycleHasPo) === "YES" ? "YES" : "NO",
    factoryJobName: safeText(source.factoryJobName),
    factoryDeliveryDate: safeText(source.factoryDeliveryDate),
    factoryDeliveryTime: safeText(source.factoryDeliveryTime),
    factoryDeliveryLocation: safeText(source.factoryDeliveryLocation),
    factoryDeliveryMapUrl: safeText(source.factoryDeliveryMapUrl),
    factoryReceiverPhone: safeText(source.factoryReceiverPhone),
    factoryTempDocumentRef: safeText(source.factoryTempDocumentRef),
    factoryReservationRef: safeText(source.factoryReservationRef),
    shuttlePassengers: safeText(source.shuttlePassengers),
    shuttleTravelDate: safeText(source.shuttleTravelDate),
    shuttleTravelTime: safeText(source.shuttleTravelTime),
    shuttleReturnWait: safeText(source.shuttleReturnWait),
    shuttleLocation: safeText(source.shuttleLocation),
    shuttleMapUrl: safeText(source.shuttleMapUrl),
    policyAgree: safeText(source.policyAgree),
  };
}

function sanitizeAttachments(data) {
  const source = data && typeof data === "object" ? data : {};
  return {
    offcyclePoImage: normalizeAttachment(source.offcyclePoImage),
    factoryTempDocumentImage: normalizeAttachment(source.factoryTempDocumentImage),
    factoryReservationImage: normalizeAttachment(source.factoryReservationImage),
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
    if (data.offcycleHasPo !== "YES" && data.offcycleHasPo !== "NO") {
      throw new Error("Invalid offcycleHasPo.");
    }
    if (data.offcycleHasPo === "YES") {
      assertValidAttachment(attachments.offcyclePoImage, "offcyclePoImage");
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

function canUseNoReply() {
  try {
    const effectiveEmail = safeText(Session.getEffectiveUser().getEmail());
    return effectiveEmail && !/@gmail\.com$/i.test(effectiveEmail);
  } catch (error) {
    return false;
  }
}

function isNoReplyUnsupportedError(error) {
  const text = safeText(error && error.message ? error.message : error).toLowerCase();
  return text.indexOf("noreply") > -1 || text.indexOf("no reply") > -1;
}

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
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
  const isDirectSubmitRunError =
    raw.indexOf("submitRequest requires request data from web form") > -1 ||
    (raw.indexOf("Cannot read properties of undefined") > -1 && raw.indexOf("requestTopic") > -1);
  const isDriveAuthError =
    raw.indexOf("DriveApp.getFolderById") > -1 ||
    raw.indexOf("https://www.googleapis.com/auth/drive") > -1 ||
    raw.indexOf("authorization") > -1 && raw.indexOf("DriveApp") > -1;

  if (isDirectSubmitRunError) {
    return (
      "ห้ามกด Run ฟังก์ชัน submitRequest() ตรงๆ ใน Apps Script. " +
      "ให้ทดสอบโดยส่งฟอร์มจากหน้าเว็บ index.html หรือเรียก Web App URL ด้วย action=submitRequest ผ่าน POST เท่านั้น."
    );
  }

  if (isDriveAuthError) {
    return (
      "Apps Script ยังไม่ได้อนุญาตสิทธิ์ Google Drive. " +
      "ให้ Run ฟังก์ชัน authorizeDriveAccess() 1 ครั้ง แล้ว Deploy Web App เวอร์ชันใหม่ (Execute as: Me)."
    );
  }

  return raw || "Unexpected error.";
}
