(function () {
  const config = window.APP_CONFIG || {};
  const AUTH_STORAGE_KEY = "scotch_webapps1_admin_auth";

  const STATUS_PENDING = "PENDING";
  const STATUS_APPROVED = "APPROVED";
  const STATUS_REJECTED = "REJECTED";

  const VIEW_PENDING = "pending";
  const VIEW_HISTORY = "history";

  const TOPIC_OFF_CYCLE = "OFF_CYCLE_DELIVERY";
  const CASE_FACTORY_DELIVERY = "FACTORY_DELIVERY_DOCS";
  const CASE_FACTORY_SHUTTLE = "FACTORY_STAFF_SHUTTLE";

  const pendingPanel = document.getElementById("pendingPanel");
  const historyPanel = document.getElementById("historyPanel");
  const pendingListEl = document.getElementById("pendingList");
  const historyListEl = document.getElementById("historyList");
  const pendingCountEl = document.getElementById("pendingCount");
  const historyCountEl = document.getElementById("historyCount");

  const pendingTabBtn = document.getElementById("pendingTabBtn");
  const historyTabBtn = document.getElementById("historyTabBtn");

  const pendingSearchInput = document.getElementById("pendingSearchInput");
  const pendingTypeFilter = document.getElementById("pendingTypeFilter");
  const pendingSortOrder = document.getElementById("pendingSortOrder");

  const historySearchInput = document.getElementById("historySearchInput");
  const historyTypeFilter = document.getElementById("historyTypeFilter");
  const historySortBy = document.getElementById("historySortBy");
  const historySortOrder = document.getElementById("historySortOrder");

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
  let allRequests = [];
  let activeView = VIEW_PENDING;

  if (!pendingListEl || !historyListEl) return;

  bindEvents();
  bootstrap();

  function bindEvents() {
    if (refreshBtn) refreshBtn.addEventListener("click", loadRequests);
    if (logoutAdminBtn) logoutAdminBtn.addEventListener("click", exitAdminMode);

    if (pendingTabBtn) pendingTabBtn.addEventListener("click", () => setActiveView(VIEW_PENDING));
    if (historyTabBtn) historyTabBtn.addEventListener("click", () => setActiveView(VIEW_HISTORY));

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

    bindFilterEvents();
  }

  function bindFilterEvents() {
    if (pendingSearchInput) pendingSearchInput.addEventListener("input", renderAllViews);
    if (pendingTypeFilter) pendingTypeFilter.addEventListener("change", renderAllViews);
    if (pendingSortOrder) pendingSortOrder.addEventListener("change", renderAllViews);

    if (historySearchInput) historySearchInput.addEventListener("input", renderAllViews);
    if (historyTypeFilter) historyTypeFilter.addEventListener("change", renderAllViews);
    if (historySortBy) historySortBy.addEventListener("change", renderAllViews);
    if (historySortOrder) historySortOrder.addEventListener("change", renderAllViews);
  }

  async function bootstrap() {
    setActiveView(VIEW_PENDING);
    setAdminNameDisplay("");

    if (!isApiConfigured()) {
      setAuthenticated(false);
      renderSignedOutState("Set assets/js/config.js with API URL before use.");
      showMessage("API URL is not configured.", "error");
      setLoginMessage("API URL is not configured.", "error");
      return;
    }

    const storedAuth = readStoredAuth();
    if (!storedAuth) {
      setAuthenticated(false);
      renderSignedOutState("Please sign in as admin.");
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
      renderSignedOutState("Please sign in as admin.");

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
      renderSignedOutState("Please sign in as admin.");
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

      allRequests = normalizeRequests(data.requests || []);
      renderAllViews();
      showMessage("Requests loaded.", "success");
    } catch (error) {
      allRequests = [];
      renderPanelEmpty(pendingListEl, pendingCountEl, "Failed to load requests.");
      renderPanelEmpty(historyListEl, historyCountEl, "Failed to load requests.");
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

  function normalizeRequests(requests) {
    return (Array.isArray(requests) ? requests : []).map((request) => {
      const item = request && typeof request === "object" ? { ...request } : {};
      item.status = normalizeStatus(item.status);
      return item;
    });
  }

  function renderAllViews() {
    const pendingSource = allRequests.filter((request) => normalizeStatus(request.status) === STATUS_PENDING);
    const historySource = allRequests.filter((request) => {
      const status = normalizeStatus(request.status);
      return status === STATUS_APPROVED || status === STATUS_REJECTED;
    });

    const pendingFiltered = sortRequests(
      applyCommonFilters(
        pendingSource,
        pendingSearchInput ? pendingSearchInput.value : "",
        pendingTypeFilter ? pendingTypeFilter.value : "ALL",
      ),
      pendingSortOrder ? pendingSortOrder.value : "oldest",
      "submitted",
    );

    const historyFiltered = sortRequests(
      applyCommonFilters(
        historySource,
        historySearchInput ? historySearchInput.value : "",
        historyTypeFilter ? historyTypeFilter.value : "ALL",
      ),
      historySortOrder ? historySortOrder.value : "oldest",
      historySortBy ? historySortBy.value : "submitted",
    );

    renderRequestList({
      listEl: pendingListEl,
      countEl: pendingCountEl,
      requests: pendingFiltered,
      totalCount: pendingSource.length,
      emptyText: "No pending requests.",
      historyMode: false,
    });

    renderRequestList({
      listEl: historyListEl,
      countEl: historyCountEl,
      requests: historyFiltered,
      totalCount: historySource.length,
      emptyText: "No approved/rejected history yet.",
      historyMode: true,
    });
  }

  function renderRequestList(options) {
    const listEl = options.listEl;
    const countEl = options.countEl;
    const requests = Array.isArray(options.requests) ? options.requests : [];
    const totalCount = Number(options.totalCount || 0);

    setCountText(countEl, requests.length, totalCount);

    if (!requests.length) {
      listEl.innerHTML = `<p class="empty-state">${escapeHtml(options.emptyText || "No records found.")}</p>`;
      return;
    }

    listEl.innerHTML = "";
    requests.forEach((request) => {
      listEl.appendChild(renderRequestCard(request, options.historyMode));
    });
  }

  function setCountText(countEl, filteredCount, totalCount) {
    if (!countEl) return;
    if (!totalCount) {
      countEl.textContent = "0 records";
      return;
    }

    if (filteredCount === totalCount) {
      countEl.textContent = totalCount + " records";
      return;
    }

    countEl.textContent = filteredCount + " / " + totalCount + " records";
  }

  function renderRequestCard(request, historyMode) {
    const item = document.createElement("article");
    const status = normalizeStatus(request.status);
    const statusClass = statusToClass(status);

    const topicLabel = getTopicLabel(request);
    const caseLabel = getCaseLabel(request);
    const submittedAt = request.submittedAt || "-";

    const decisionAt =
      status === STATUS_APPROVED
        ? request.approvedAt || "-"
        : status === STATUS_REJECTED
          ? request.rejectedAt || "-"
          : "-";

    const decisionBy =
      status === STATUS_APPROVED
        ? request.approvedBy || "-"
        : status === STATUS_REJECTED
          ? request.rejectedBy || "-"
          : "-";

    item.className = "request-item request-item-" + statusClass;

    item.innerHTML =
      '<div class="request-item-head">' +
      '<div class="request-head-main">' +
      '<h3 class="request-id">' + escapeHtml(request.requestId || "-") + "</h3>" +
      '<p class="request-caption">' +
      escapeHtml(topicLabel) +
      (caseLabel ? "<span> | </span>" + escapeHtml(caseLabel) : "") +
      "</p>" +
      "</div>" +
      '<span class="status-pill ' + statusClass + '">' + escapeHtml(status) + "</span>" +
      "</div>" +
      '<div class="request-time-row">' +
      '<span class="time-chip"><strong>ส่งคำขอ:</strong> ' + escapeHtml(submittedAt) + "</span>" +
      (status !== STATUS_PENDING
        ? '<span class="time-chip"><strong>' +
          (status === STATUS_APPROVED ? "อนุมัติเมื่อ" : "ไม่อนุมัติเมื่อ") +
          ":</strong> " +
          escapeHtml(decisionAt) +
          "</span>"
        : "") +
      "</div>" +
      '<div class="request-section-grid">' +
      renderInfoBlock("ผู้ร้องขอ", [
        { label: "ผู้ขอใช้", value: request.requesterName || "-" },
        { label: "เบอร์ติดต่อ", value: request.requesterPhone || "-" },
        { label: "อีเมล", value: request.contactEmail || "-" },
        { label: "หมายเหตุผู้ร้องขอ", value: request.requesterRemark || "-" },
      ]) +
      renderCaseInfoBlock(request) +
      renderAttachmentBlock(request) +
      (historyMode
        ? renderInfoBlock("ผลการพิจารณา", [
            { label: "พิจารณาโดย", value: decisionBy },
            { label: "Remark", value: request.decisionRemark || "-" },
            { label: "สถานะอีเมล", value: request.emailStatus || "-" },
          ])
        : "") +
      "</div>";

    if (!historyMode && status === STATUS_PENDING) {
      const actions = document.createElement("div");
      actions.className = "request-actions";

      const approveBtn = document.createElement("button");
      approveBtn.type = "button";
      approveBtn.textContent = "อนุมัติ + ส่งเมล";

      const rejectBtn = document.createElement("button");
      rejectBtn.type = "button";
      rejectBtn.className = "button-danger";
      rejectBtn.textContent = "ไม่อนุมัติ + ส่งเมล";

      approveBtn.addEventListener("click", () => approveRequest(request.requestId, approveBtn, rejectBtn));
      rejectBtn.addEventListener("click", () => rejectRequest(request.requestId, approveBtn, rejectBtn));

      actions.appendChild(approveBtn);
      actions.appendChild(rejectBtn);
      item.appendChild(actions);
    }

    return item;
  }

  function renderCaseInfoBlock(request) {
    if (request.requestTopic === TOPIC_OFF_CYCLE) {
      return renderInfoBlock("รายละเอียดคำขอ", [
        { label: "ชื่อห้าง", value: request.offcycleStoreName || "-" },
        { label: "วันที่ต้องการจัดส่ง", value: request.offcycleDeliveryDate || "-" },
        { label: "จำนวนลังสินค้า", value: request.offcycleCrates || "-" },
        { label: "ยอดเงิน", value: request.offcycleAmount || "-" },
        { label: "มี PO", value: request.offcycleHasPo === "YES" ? "มี" : "ไม่มี" },
      ]);
    }

    if (request.factoryCaseType === CASE_FACTORY_DELIVERY) {
      return renderInfoBlock("รายละเอียดคำขอ", [
        { label: "ชื่องาน", value: request.factoryJobName || "-" },
        { label: "วันที่จัดส่ง", value: request.factoryDeliveryDate || "-" },
        { label: "เวลาจัดส่ง", value: request.factoryDeliveryTime || "-" },
        { label: "สถานที่จัดส่ง", value: request.factoryDeliveryLocation || "-" },
        { label: "Google Map", value: request.factoryDeliveryMapUrl || "-", allowUrl: true },
        { label: "เบอร์ติดต่อหน้างาน", value: request.factoryReceiverPhone || "-" },
        { label: "รายละเอียดเอกสาร", value: request.factoryTempDocumentRef || "-" },
        { label: "รายละเอียด Reservation", value: request.factoryReservationRef || "-" },
      ]);
    }

    if (request.factoryCaseType === CASE_FACTORY_SHUTTLE) {
      return renderInfoBlock("รายละเอียดคำขอ", [
        { label: "จำนวนผู้โดยสาร", value: request.shuttlePassengers || "-" },
        { label: "วันที่เดินทาง", value: request.shuttleTravelDate || "-" },
        { label: "เวลาเดินทาง", value: request.shuttleTravelTime || "-" },
        { label: "รอรับกลับ", value: shuttleWaitLabel(request.shuttleReturnWait) },
        { label: "สถานที่", value: request.shuttleLocation || "-" },
        { label: "Google Map", value: request.shuttleMapUrl || "-", allowUrl: true },
      ]);
    }

    return renderInfoBlock("รายละเอียดคำขอ", [{ label: "ข้อมูล", value: "-" }]);
  }

  function renderAttachmentBlock(request) {
    const pairs = [];

    if (request.requestTopic === TOPIC_OFF_CYCLE) {
      pairs.push(
        attachmentPair(
          "ไฟล์ PO",
          request.offcyclePoImageUrl,
          request.offcyclePoImageDownloadUrl,
          request.offcyclePoImageName,
        ),
      );
    }

    if (request.factoryCaseType === CASE_FACTORY_DELIVERY) {
      pairs.push(
        attachmentPair(
          "ไฟล์เอกสารชั่วคราว",
          request.factoryTempDocumentImageUrl,
          request.factoryTempDocumentImageDownloadUrl,
          request.factoryTempDocumentImageName,
        ),
      );
      pairs.push(
        attachmentPair(
          "ไฟล์ Reservation",
          request.factoryReservationImageUrl,
          request.factoryReservationImageDownloadUrl,
          request.factoryReservationImageName,
        ),
      );
    }

    if (!pairs.length) {
      return renderInfoBlock("ไฟล์แนบ", [{ label: "ไฟล์", value: "-" }]);
    }

    return renderInfoBlock("ไฟล์แนบ", pairs);
  }

  function renderInfoBlock(title, pairs) {
    const rows = (Array.isArray(pairs) ? pairs : [])
      .map((pair) => {
        const label = escapeHtml(pair.label || "-");
        const valueHtml = renderPairValue(pair);
        return (
          '<div class="kv-item">' +
          '<span class="kv-label">' +
          label +
          "</span>" +
          '<div class="kv-value">' +
          valueHtml +
          "</div>" +
          "</div>"
        );
      })
      .join("");

    return (
      '<section class="request-info-block">' +
      "<h4>" +
      escapeHtml(title || "รายละเอียด") +
      "</h4>" +
      '<div class="kv-grid">' +
      rows +
      "</div>" +
      "</section>"
    );
  }

  function renderPairValue(pair) {
    if (pair && pair.asHtml) {
      return pair.value || "-";
    }

    const text = String((pair && pair.value) || "-");
    if (pair && pair.allowUrl && isValidUrl(text)) {
      const safeUrl = escapeAttr(text);
      return '<a class="inline-action" href="' + safeUrl + '" target="_blank" rel="noopener noreferrer">เปิดลิงก์</a>';
    }

    return escapeHtml(text);
  }

  function attachmentPair(label, viewUrl, downloadUrl, fileName) {
    if (!viewUrl && !downloadUrl) {
      return { label: label, value: "-" };
    }

    const links = [];
    if (viewUrl) {
      links.push(
        '<a class="inline-action" href="' +
          escapeAttr(viewUrl) +
          '" target="_blank" rel="noopener noreferrer">เปิดดู</a>',
      );
    }
    if (downloadUrl) {
      links.push(
        '<a class="inline-action" href="' +
          escapeAttr(downloadUrl) +
          '" target="_blank" rel="noopener noreferrer" download>ดาวน์โหลด</a>',
      );
    }

    const meta = fileName
      ? '<span class="file-meta">' + escapeHtml(fileName) + "</span>"
      : '<span class="file-meta">แนบไฟล์แล้ว</span>';

    return {
      label: label,
      value: '<div class="attachment-links">' + links.join("") + meta + "</div>",
      asHtml: true,
    };
  }

  function applyCommonFilters(requests, searchText, typeFilter) {
    const query = String(searchText || "").trim().toLowerCase();
    const type = String(typeFilter || "ALL").trim();

    return (Array.isArray(requests) ? requests : []).filter((request) => {
      if (!matchTypeFilter(request, type)) return false;
      if (!query) return true;
      return matchSearchQuery(request, query);
    });
  }

  function matchTypeFilter(request, typeFilter) {
    if (!typeFilter || typeFilter === "ALL") return true;
    if (typeFilter === TOPIC_OFF_CYCLE) return request.requestTopic === TOPIC_OFF_CYCLE;
    if (typeFilter === CASE_FACTORY_DELIVERY) return request.factoryCaseType === CASE_FACTORY_DELIVERY;
    if (typeFilter === CASE_FACTORY_SHUTTLE) return request.factoryCaseType === CASE_FACTORY_SHUTTLE;
    return true;
  }

  function matchSearchQuery(request, query) {
    const buffer = [
      request.requestId,
      getTopicLabel(request),
      getCaseLabel(request),
      request.requesterName,
      request.requesterPhone,
      request.contactEmail,
      request.requesterRemark,
      request.offcycleStoreName,
      request.factoryJobName,
      request.factoryDeliveryLocation,
      request.shuttleLocation,
      request.factoryTempDocumentRef,
      request.factoryReservationRef,
      request.decisionRemark,
      request.submittedAt,
      request.approvedAt,
      request.rejectedAt,
    ]
      .map((value) => String(value || ""))
      .join(" ")
      .toLowerCase();

    return buffer.indexOf(query) > -1;
  }

  function sortRequests(requests, sortOrder, sortBy) {
    const order = sortOrder === "latest" ? "latest" : "oldest";
    const basis = sortBy === "decision" ? "decision" : "submitted";

    return [...requests].sort((a, b) => {
      const aPrimary = getSortTimestamp(a, basis);
      const bPrimary = getSortTimestamp(b, basis);

      if (aPrimary !== bPrimary) {
        return order === "latest" ? bPrimary - aPrimary : aPrimary - bPrimary;
      }

      const aFallback = getSortTimestamp(a, "submitted");
      const bFallback = getSortTimestamp(b, "submitted");
      if (aFallback !== bFallback) {
        return order === "latest" ? bFallback - aFallback : aFallback - bFallback;
      }

      return String(a.requestId || "").localeCompare(String(b.requestId || ""));
    });
  }

  function getSortTimestamp(request, sortBy) {
    if (sortBy === "decision") {
      const status = normalizeStatus(request.status);
      if (status === STATUS_APPROVED) {
        return parseDateTimeValue(request.approvedAt) || parseDateTimeValue(request.submittedAt);
      }
      if (status === STATUS_REJECTED) {
        return parseDateTimeValue(request.rejectedAt) || parseDateTimeValue(request.submittedAt);
      }
    }

    return parseDateTimeValue(request.submittedAt);
  }

  function parseDateTimeValue(value) {
    const text = String(value || "").trim();
    if (!text) return 0;

    const normalized = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(text)
      ? text.replace(" ", "T")
      : /^\d{4}-\d{2}-\d{2}$/.test(text)
        ? text + "T00:00:00"
        : text;

    const timestamp = Date.parse(normalized);
    if (Number.isNaN(timestamp)) return 0;
    return timestamp;
  }

  function statusToClass(status) {
    if (status === STATUS_APPROVED) return "approved";
    if (status === STATUS_REJECTED) return "rejected";
    return "pending";
  }

  function normalizeStatus(status) {
    return String(status || STATUS_PENDING)
      .trim()
      .toUpperCase();
  }

  function getTopicLabel(request) {
    return request.requestTopicLabel || request.requestTopic || "-";
  }

  function getCaseLabel(request) {
    return request.factoryCaseLabel || "";
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

  function setActiveView(view) {
    activeView = view === VIEW_HISTORY ? VIEW_HISTORY : VIEW_PENDING;

    if (pendingPanel) pendingPanel.classList.toggle("hidden", activeView !== VIEW_PENDING);
    if (historyPanel) historyPanel.classList.toggle("hidden", activeView !== VIEW_HISTORY);

    if (pendingTabBtn) pendingTabBtn.classList.toggle("is-active", activeView === VIEW_PENDING);
    if (historyTabBtn) historyTabBtn.classList.toggle("is-active", activeView === VIEW_HISTORY);
  }

  function setAuthenticated(isAuthenticated) {
    if (loginOverlay) loginOverlay.classList.toggle("is-hidden", isAuthenticated);
    if (refreshBtn) refreshBtn.disabled = !isAuthenticated;

    const tabs = [pendingTabBtn, historyTabBtn].filter(Boolean);
    tabs.forEach((tab) => {
      tab.disabled = !isAuthenticated;
    });
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

  function renderSignedOutState(message) {
    allRequests = [];
    renderPanelEmpty(pendingListEl, pendingCountEl, message || "Please sign in as admin.");
    renderPanelEmpty(historyListEl, historyCountEl, message || "Please sign in as admin.");
  }

  function renderPanelEmpty(listEl, countEl, text) {
    if (listEl) {
      listEl.innerHTML = `<p class="empty-state">${escapeHtml(text || "No records found.")}</p>`;
    }
    if (countEl) {
      countEl.textContent = "0 records";
    }
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
    setActiveView(VIEW_PENDING);

    if (loginPassInput) loginPassInput.value = "";

    setAuthenticated(false);
    renderSignedOutState("Logged out from admin mode. Please sign in again.");
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

  function isValidUrl(url) {
    const text = String(url || "").trim();
    return /^https?:\/\/.+/i.test(text);
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
