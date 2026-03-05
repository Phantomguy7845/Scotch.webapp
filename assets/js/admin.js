(function () {
  const config = window.APP_CONFIG || {};
  const AUTH_STORAGE_KEY = "scotch_webapps1_admin_auth";

  const listEl = document.getElementById("requestList");
  const countEl = document.getElementById("requestCount");
  const messageEl = document.getElementById("adminMessage");
  const adminNameLabel = document.getElementById("adminNameLabel");
  const refreshBtn = document.getElementById("refreshBtn");
  const logoutAdminBtn = document.getElementById("logoutAdminBtn");

  const loginOverlay = document.getElementById("adminLoginOverlay");
  const loginNameInput = document.getElementById("adminLoginNameInput");
  const loginPassInput = document.getElementById("adminLoginPassInput");
  const loginBtn = document.getElementById("adminLoginBtn");
  const loginMessageEl = document.getElementById("adminLoginMessage");

  let authSession = null;

  if (!listEl) return;

  if (refreshBtn) refreshBtn.addEventListener("click", loadRequests);
  if (logoutAdminBtn) logoutAdminBtn.addEventListener("click", exitAdminMode);
  if (loginBtn) loginBtn.addEventListener("click", loginWithForm);
  if (loginNameInput) {
    loginNameInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        loginWithForm();
      }
    });
  }
  if (loginPassInput) {
    loginPassInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        loginWithForm();
      }
    });
  }

  bootstrap();

  async function bootstrap() {
    setAdminNameDisplay("");

    if (!isApiConfigured()) {
      setAuthenticated(false);
      renderEmpty("Set assets/js/config.js with API URL before use.");
      showMessage("API URL is not configured.", "error");
      setLoginMessage("API URL is not configured.", "error");
      return;
    }

    const storedAuth = readStoredAuth();
    if (!storedAuth) {
      setAuthenticated(false);
      renderEmpty("Please sign in as admin.");
      return;
    }

    if (loginNameInput) loginNameInput.value = storedAuth.adminName;
    if (loginPassInput) loginPassInput.value = storedAuth.adminPass;
    await tryLogin(storedAuth, true);
  }

  async function loginWithForm() {
    const credentials = {
      adminName: loginNameInput ? loginNameInput.value.trim() : "",
      adminPass: loginPassInput ? loginPassInput.value.trim() : "",
    };
    if (!credentials.adminName || !credentials.adminPass) {
      setLoginMessage("กรุณากรอก Admin name และ Password", "error");
      return;
    }
    await tryLogin(credentials, false);
  }

  async function tryLogin(credentials, silent) {
    setLoginBusy(true);
    clearMessage();
    if (!silent) setLoginMessage("", "");

    try {
      const validation = await validateAdmin(credentials);
      if (!validation.ok) {
        throw new Error(validation.error || "Invalid admin login.");
      }

      authSession = {
        adminName: credentials.adminName.trim(),
        adminPass: credentials.adminPass.trim(),
        adminId: validation.adminId || "",
      };
      writeStoredAuth(authSession);
      setAuthenticated(true);
      setAdminNameDisplay(authSession.adminName);

      setLoginMessage("", "");
      if (!silent) {
        showMessage("Login success: " + authSession.adminName, "success");
      }
      await loadRequests();
      return true;
    } catch (error) {
      authSession = null;
      clearStoredAuth();
      setAdminNameDisplay("");
      setAuthenticated(false);
      renderEmpty("Please sign in as admin.");
      if (!silent) {
        setLoginMessage(error.message || "Login failed.", "error");
      } else {
        setLoginMessage("", "");
      }
      return false;
    } finally {
      setLoginBusy(false);
    }
  }

  async function validateAdmin(credentials) {
    const url = new URL(config.apiBaseUrl);
    url.searchParams.set("action", "validateAdmin");
    url.searchParams.set("adminName", credentials.adminName);
    url.searchParams.set("adminPass", credentials.adminPass);

    const response = await fetch(url.toString());
    return response.json();
  }

  async function loadRequests() {
    clearMessage();

    if (!authSession) {
      setAuthenticated(false);
      renderEmpty("Please sign in as admin.");
      showMessage("Admin login is required.", "error");
      return;
    }

    try {
      if (refreshBtn) {
        refreshBtn.disabled = true;
        refreshBtn.textContent = "Loading...";
      }

      const url = new URL(config.apiBaseUrl);
      url.searchParams.set("action", "listRequests");
      url.searchParams.set("adminName", authSession.adminName);
      url.searchParams.set("adminPass", authSession.adminPass);

      const response = await fetch(url.toString());
      const data = await response.json();
      if (!data.ok) throw new Error(data.error || "Could not load requests.");

      renderRequests(data.requests || []);
      showMessage("Requests loaded.", "success");
    } catch (error) {
      renderEmpty("Failed to load requests.");
      showMessage(error.message || "Unexpected error while loading.", "error");
      if (String(error.message || "").toLowerCase().indexOf("invalid admin") > -1) {
        exitAdminMode();
      }
    } finally {
      if (refreshBtn) {
        refreshBtn.disabled = false;
        refreshBtn.textContent = "Refresh requests";
      }
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
          "ไฟล์ PO",
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
          "ไฟล์เอกสารชั่วคราว",
          request.factoryTempDocumentImageUrl,
          request.factoryTempDocumentImageDownloadUrl,
          request.factoryTempDocumentImageName,
        ),
        detailRow("รายละเอียดเบิกของ/Reservation", request.factoryReservationRef || "-"),
        attachmentRow(
          "ไฟล์ Reservation",
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
    if (!authSession) {
      showMessage("Admin login is required.", "error");
      setAuthenticated(false);
      return;
    }

    const remark = requestDecisionRemark(options.decisionText, options.requestId);
    if (remark === null) return;
    if (!remark) {
      showMessage("กรุณาระบุ Remark ก่อนส่งผลการพิจารณา", "error");
      return;
    }

    const approvedBy = authSession.adminName || "Fleet Admin";
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
        adminName: authSession.adminName,
        adminPass: authSession.adminPass,
        requestId: options.requestId,
        approvedBy: approvedBy,
        remark: remark,
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
      if (String(error.message || "").toLowerCase().indexOf("invalid admin") > -1) {
        exitAdminMode();
      }
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

  function setAuthenticated(isAuthenticated) {
    if (loginOverlay) loginOverlay.classList.toggle("is-hidden", isAuthenticated);
    if (refreshBtn) refreshBtn.disabled = !isAuthenticated;
  }

  function setAdminNameDisplay(adminName) {
    if (!adminNameLabel) return;
    adminNameLabel.textContent = adminName || "-";
  }

  function setLoginBusy(isBusy) {
    if (loginBtn) {
      loginBtn.disabled = isBusy;
      loginBtn.textContent = isBusy ? "กำลังตรวจสอบ..." : "เข้าสู่ระบบแอดมิน";
    }
  }

  function setLoginMessage(text, type) {
    if (!loginMessageEl) return;
    loginMessageEl.textContent = text || "";
    loginMessageEl.classList.remove("success", "error");
    if (type) loginMessageEl.classList.add(type);
  }

  function renderEmpty(text) {
    listEl.innerHTML = `<p class="empty-state">${escapeHtml(text)}</p>`;
    countEl.textContent = "0 records";
  }

  function readStoredAuth() {
    try {
      const raw = localStorage.getItem(AUTH_STORAGE_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") return null;
      const adminName = String(parsed.adminName || "").trim();
      const adminPass = String(parsed.adminPass || "").trim();
      if (!adminName || !adminPass) return null;
      return { adminName, adminPass };
    } catch (error) {
      return null;
    }
  }

  function writeStoredAuth(auth) {
    localStorage.setItem(
      AUTH_STORAGE_KEY,
      JSON.stringify({
        adminName: auth.adminName,
        adminPass: auth.adminPass,
      }),
    );
  }

  function clearStoredAuth() {
    localStorage.removeItem(AUTH_STORAGE_KEY);
  }

  function exitAdminMode() {
    authSession = null;
    clearStoredAuth();
    setAdminNameDisplay("");
    if (loginPassInput) loginPassInput.value = "";
    setAuthenticated(false);
    renderEmpty("Logged out from admin mode. Please sign in again.");
    showMessage("Logged out from admin mode.", "success");
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
