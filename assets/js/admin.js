(function () {
  const config = window.APP_CONFIG || {};
  const STORAGE_KEY = "scotch_webapps1_admin_key";
  const listEl = document.getElementById("requestList");
  const countEl = document.getElementById("requestCount");
  const messageEl = document.getElementById("adminMessage");
  const adminKeyInput = document.getElementById("adminKeyInput");
  const approvedByInput = document.getElementById("approvedByInput");
  const saveAdminKeyBtn = document.getElementById("saveAdminKeyBtn");
  const refreshBtn = document.getElementById("refreshBtn");
  const logoutAdminBtn = document.getElementById("logoutAdminBtn");

  if (!listEl) return;

  adminKeyInput.value = localStorage.getItem(STORAGE_KEY) || config.adminKey || "";
  preventAdminKeyCopy(adminKeyInput);
  adminKeyInput.addEventListener("input", persistAdminKey);

  saveAdminKeyBtn.addEventListener("click", () => {
    localStorage.setItem(STORAGE_KEY, adminKeyInput.value.trim());
    showMessage("Admin key saved in browser storage.", "success");
  });

  refreshBtn.addEventListener("click", loadRequests);
  if (logoutAdminBtn) logoutAdminBtn.addEventListener("click", exitAdminMode);
  loadRequests();

  async function loadRequests() {
    clearMessage();

    const adminKey = getAdminKey();
    if (!isApiConfigured()) {
      renderEmpty("Set assets/js/config.js with API URL before use.");
      showMessage("API URL is not configured.", "error");
      return;
    }
    if (!adminKey) {
      renderEmpty("Enter and save admin key, then refresh.");
      showMessage("Admin key is required.", "error");
      return;
    }

    try {
      refreshBtn.disabled = true;
      refreshBtn.textContent = "Loading...";

      const url = new URL(config.apiBaseUrl);
      url.searchParams.set("action", "listRequests");
      url.searchParams.set("adminKey", adminKey);

      const response = await fetch(url.toString());
      const data = await response.json();
      if (!data.ok) throw new Error(data.error || "Could not load requests.");

      renderRequests(data.requests || []);
      showMessage("Requests loaded.", "success");
    } catch (error) {
      renderEmpty("Failed to load requests.");
      showMessage(error.message || "Unexpected error while loading.", "error");
    } finally {
      refreshBtn.disabled = false;
      refreshBtn.textContent = "Refresh requests";
    }
  }

  function renderRequests(requests) {
    const sorted = [...requests].sort((a, b) => (b.submittedAt || "").localeCompare(a.submittedAt || ""));
    countEl.textContent = sorted.length + " records";

    if (!sorted.length) {
      renderEmpty("No records found.");
      return;
    }

    listEl.innerHTML = "";
    sorted.forEach((request) => {
      const item = document.createElement("article");
      item.className = "request-item";

      const status = (request.status || "PENDING").toUpperCase();
      let statusClass = "pending";
      if (status === "APPROVED") statusClass = "approved";
      if (status === "REJECTED") statusClass = "rejected";

      item.innerHTML = `
        <div class="request-item-head">
          <h3 class="request-id">${escapeHtml(request.requestId || "-")}</h3>
          <span class="status-pill ${statusClass}">${escapeHtml(status)}</span>
        </div>
        <div class="request-meta">
          ${renderCommonDetails(request)}
          ${renderCaseDetails(request)}
        </div>
      `;

      if (status === "PENDING") {
        const actions = document.createElement("div");
        actions.className = "request-actions";

        const approveBtn = document.createElement("button");
        approveBtn.type = "button";
        approveBtn.textContent = "อนุมัติ + ส่งเมล";

        const rejectBtn = document.createElement("button");
        rejectBtn.type = "button";
        rejectBtn.className = "button-danger";
        rejectBtn.textContent = "ไม่อนุมัติ + ส่งเมล";

        approveBtn.addEventListener("click", () =>
          approveRequest(request.requestId, approveBtn, rejectBtn),
        );
        rejectBtn.addEventListener("click", () => rejectRequest(request.requestId, approveBtn, rejectBtn));

        actions.appendChild(approveBtn);
        actions.appendChild(rejectBtn);
        item.appendChild(actions);
      }

      listEl.appendChild(item);
    });
  }

  function renderCommonDetails(request) {
    return [
      detailRow("หัวข้อ", request.requestTopicLabel || request.requestTopic || "-"),
      detailRow("กรณี", request.factoryCaseLabel || "-"),
      detailRow("ผู้ขอใช้", request.requesterName || "-"),
      detailRow("เบอร์ติดต่อผู้ขอ", request.requesterPhone || "-"),
      detailRow("อีเมลติดต่อ", request.contactEmail || "-"),
      detailRow("ส่งคำขอเมื่อ", request.submittedAt || "-"),
      detailRow("อนุมัติเมื่อ", request.approvedAt || "-"),
      detailRow("อนุมัติโดย", request.approvedBy || "-"),
      detailRow("ไม่อนุมัติเมื่อ", request.rejectedAt || "-"),
      detailRow("ไม่อนุมัติโดย", request.rejectedBy || "-"),
      detailRow("Remark", request.decisionRemark || "-"),
      detailRow("สถานะอีเมล", request.emailStatus || "-"),
    ].join("");
  }

  function renderCaseDetails(request) {
    if (request.requestTopic === "OFF_CYCLE_DELIVERY") {
      return [
        detailRow("ชื่อห้าง", request.offcycleStoreName || "-"),
        detailRow("วันที่ต้องการจัดส่ง", request.offcycleDeliveryDate || "-"),
        detailRow("จำนวนลังสินค้า", request.offcycleCrates || "-"),
        detailRow("ยอดเงิน", request.offcycleAmount || "-"),
        detailRow("มี PO", request.offcycleHasPo === "YES" ? "มี" : "ไม่มี"),
        attachmentRow(
          "รูป PO",
          request.offcyclePoImageUrl,
          request.offcyclePoImageDownloadUrl,
          request.offcyclePoImageName,
        ),
      ].join("");
    }

    if (request.factoryCaseType === "FACTORY_DELIVERY_DOCS") {
      return [
        detailRow("ชื่องาน", request.factoryJobName || "-"),
        detailRow("วันที่จัดส่ง", request.factoryDeliveryDate || "-"),
        detailRow("เวลาจัดส่ง", request.factoryDeliveryTime || "-"),
        detailRow("สถานที่จัดส่ง", request.factoryDeliveryLocation || "-"),
        detailRow("Google Map", request.factoryDeliveryMapUrl || "-"),
        detailRow("เบอร์ติดต่อผู้รับ/หน้างาน", request.factoryReceiverPhone || "-"),
        detailRow("รายละเอียดเอกสารชั่วคราว", request.factoryTempDocumentRef || "-"),
        attachmentRow(
          "รูปเอกสารชั่วคราว",
          request.factoryTempDocumentImageUrl,
          request.factoryTempDocumentImageDownloadUrl,
          request.factoryTempDocumentImageName,
        ),
        detailRow("รายละเอียดเบิกของ/Reservation", request.factoryReservationRef || "-"),
        attachmentRow(
          "รูป Reservation",
          request.factoryReservationImageUrl,
          request.factoryReservationImageDownloadUrl,
          request.factoryReservationImageName,
        ),
      ].join("");
    }

    if (request.factoryCaseType === "FACTORY_STAFF_SHUTTLE") {
      return [
        detailRow("จำนวนผู้โดยสาร", request.shuttlePassengers || "-"),
        detailRow("วันที่เดินทาง", request.shuttleTravelDate || "-"),
        detailRow("เวลาเดินทาง", request.shuttleTravelTime || "-"),
        detailRow("รอรับกลับ", shuttleWaitLabel(request.shuttleReturnWait)),
        detailRow("สถานที่", request.shuttleLocation || "-"),
        detailRow("Google Map", request.shuttleMapUrl || "-"),
      ].join("");
    }

    return "";
  }

  function detailRow(label, value) {
    return `<div><strong>${escapeHtml(label)}:</strong> ${escapeHtml(value || "-")}</div>`;
  }

  function attachmentRow(label, viewUrl, downloadUrl, fileName) {
    if (!viewUrl && !downloadUrl) {
      return detailRow(label, "-");
    }

    const links = [];
    if (viewUrl) {
      links.push(
        `<a class="inline-action" href="${escapeAttr(viewUrl)}" target="_blank" rel="noopener noreferrer">เปิดดู</a>`,
      );
    }
    if (downloadUrl) {
      links.push(
        `<a class="inline-action" href="${escapeAttr(downloadUrl)}" target="_blank" rel="noopener noreferrer" download>ดาวน์โหลด</a>`,
      );
    }

    const namePart = fileName
      ? `<span class="file-meta">${escapeHtml(fileName)}</span>`
      : '<span class="file-meta">แนบไฟล์แล้ว</span>';
    return `<div><strong>${escapeHtml(label)}:</strong> ${links.join("")} ${namePart}</div>`;
  }

  function shuttleWaitLabel(value) {
    if (value === "WAIT_RETURN") return "รอรับกลับ";
    if (value === "NO_RETURN") return "ไม่ต้องรอรับกลับ";
    return value || "-";
  }

  async function approveRequest(requestId, approveBtn, rejectBtn) {
    await handleDecisionAction({
      action: "approveRequest",
      requestId,
      decisionText: "อนุมัติ",
      processingText: "กำลังอนุมัติ...",
      doneText: "อนุมัติคำขอสำเร็จ",
      triggerBtn: approveBtn,
      peerBtn: rejectBtn,
    });
  }

  async function rejectRequest(requestId, approveBtn, rejectBtn) {
    await handleDecisionAction({
      action: "rejectRequest",
      requestId,
      decisionText: "ไม่อนุมัติ",
      processingText: "กำลังไม่อนุมัติ...",
      doneText: "ไม่อนุมัติคำขอสำเร็จ",
      triggerBtn: rejectBtn,
      peerBtn: approveBtn,
    });
  }

  async function handleDecisionAction(options) {
    const adminKey = getAdminKey();
    if (!adminKey) {
      showMessage("Admin key is required.", "error");
      return;
    }

    const remark = requestDecisionRemark(options.decisionText, options.requestId);
    if (remark === null) return;
    if (!remark) {
      showMessage("กรุณาระบุ Remark ก่อนส่งผลการพิจารณา", "error");
      return;
    }

    const approvedBy = approvedByInput.value.trim() || "Fleet Admin";
    if (!window.confirm(options.decisionText + "คำขอ " + options.requestId + " และส่งอีเมลแจ้งผลใช่หรือไม่?")) {
      return;
    }

    const buttons = [options.triggerBtn, options.peerBtn].filter(Boolean);
    const originalLabels = buttons.map((btn) => btn.textContent);
    let shouldRestoreButtons = true;

    try {
      buttons.forEach((btn) => {
        btn.disabled = true;
      });
      if (options.triggerBtn) {
        options.triggerBtn.textContent = options.processingText;
      }

      const response = await postToApi({
        action: options.action,
        adminKey,
        requestId: options.requestId,
        approvedBy,
        remark,
      });

      if (!response.ok) {
        throw new Error(response.error || "Could not update this request.");
      }

      const decisionAt = response.approvedAt || response.rejectedAt || "";
      const decisionAtText = decisionAt ? " เวลา: " + decisionAt : "";
      const emailStatus = response.emailStatus ? " Email: " + response.emailStatus : "";
      showMessage(options.doneText + " (" + options.requestId + ")." + decisionAtText + emailStatus, "success");
      await loadRequests();
      shouldRestoreButtons = false;
    } catch (error) {
      showMessage(error.message || "Unexpected error while updating request.", "error");
    } finally {
      if (shouldRestoreButtons) {
        buttons.forEach((btn, index) => {
          btn.disabled = false;
          btn.textContent = originalLabels[index];
        });
      }
    }
  }

  function requestDecisionRemark(decisionText, requestId) {
    const value = window.prompt(
      decisionText + "คำขอ " + requestId + "\nกรุณาระบุ Remark ว่า " + decisionText + " เนื่องจากอะไร",
      "",
    );
    if (value === null) return null;
    return value.trim();
  }

  function renderEmpty(text) {
    listEl.innerHTML = `<p class="empty-state">${escapeHtml(text)}</p>`;
    countEl.textContent = "0 records";
  }

  function getAdminKey() {
    return adminKeyInput.value.trim();
  }

  function exitAdminMode() {
    localStorage.removeItem(STORAGE_KEY);
    adminKeyInput.value = "";
    renderEmpty("Logged out from admin mode. Enter admin key to continue.");
    showMessage("Admin key removed from this browser.", "success");
  }

  function persistAdminKey() {
    const value = adminKeyInput.value.trim();
    if (!value) {
      localStorage.removeItem(STORAGE_KEY);
      return;
    }
    localStorage.setItem(STORAGE_KEY, value);
  }

  function preventAdminKeyCopy(inputEl) {
    if (!inputEl) return;
    const blockedEvents = ["copy", "cut", "contextmenu", "dragstart"];
    blockedEvents.forEach((eventName) => {
      inputEl.addEventListener(eventName, (event) => {
        event.preventDefault();
      });
    });
    inputEl.addEventListener("keydown", (event) => {
      if ((event.ctrlKey || event.metaKey) && (event.key === "c" || event.key === "x")) {
        event.preventDefault();
      }
    });
  }

  function isApiConfigured() {
    return Boolean(config.apiBaseUrl && !config.apiBaseUrl.includes("PASTE_YOUR_APPS_SCRIPT"));
  }

  async function postToApi(payload) {
    const body = new URLSearchParams({ payload: JSON.stringify(payload) });
    const response = await fetch(config.apiBaseUrl, {
      method: "POST",
      body,
    });
    return response.json();
  }

  function clearMessage() {
    messageEl.textContent = "";
    messageEl.classList.remove("success", "error");
  }

  function showMessage(text, type) {
    messageEl.textContent = text;
    messageEl.classList.remove("success", "error");
    if (type) messageEl.classList.add(type);
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function escapeAttr(value) {
    return escapeHtml(value).replaceAll("`", "&#96;");
  }
})();
