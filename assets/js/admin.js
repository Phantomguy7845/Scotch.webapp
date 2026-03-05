(function () {
  const config = window.APP_CONFIG || {};
  const STORAGE_KEY = "scotch_webapps1_admin_key";
  const listEl = document.getElementById("requestList");
  const countEl = document.getElementById("requestCount");
  const messageEl = document.getElementById("adminMessage");
  const apiUrlInput = document.getElementById("apiUrlInput");
  const adminKeyInput = document.getElementById("adminKeyInput");
  const approvedByInput = document.getElementById("approvedByInput");
  const saveAdminKeyBtn = document.getElementById("saveAdminKeyBtn");
  const refreshBtn = document.getElementById("refreshBtn");

  if (!listEl) return;

  apiUrlInput.value = config.apiBaseUrl || "";
  const savedKey = localStorage.getItem(STORAGE_KEY) || config.adminKey || "";
  adminKeyInput.value = savedKey;

  saveAdminKeyBtn.addEventListener("click", () => {
    const key = adminKeyInput.value.trim();
    localStorage.setItem(STORAGE_KEY, key);
    showMessage("Admin key saved in browser storage.", "success");
  });

  refreshBtn.addEventListener("click", () => {
    loadRequests();
  });

  loadRequests();

  async function loadRequests() {
    clearMessage();

    const adminKey = getAdminKey();
    if (!isApiConfigured()) {
      renderEmpty("Set assets/js/config.js with Apps Script URL before use.");
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

      if (!data.ok) {
        throw new Error(data.error || "Could not load requests.");
      }

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
    const sorted = [...requests].sort((a, b) => {
      return (b.submittedAt || "").localeCompare(a.submittedAt || "");
    });

    countEl.textContent = sorted.length + " records";
    if (sorted.length === 0) {
      renderEmpty("No records found.");
      return;
    }

    listEl.innerHTML = "";
    sorted.forEach((request) => {
      const item = document.createElement("article");
      item.className = "request-item";

      const status = (request.status || "PENDING").toUpperCase();
      const statusClass = status === "APPROVED" ? "approved" : "pending";

      item.innerHTML = `
        <div class="request-item-head">
          <h3 class="request-id">${escapeHtml(request.requestId || "-")}</h3>
          <span class="status-pill ${statusClass}">${status}</span>
        </div>
        <div class="request-meta">
          <div><strong>Name:</strong> ${escapeHtml(request.employeeName || "-")}</div>
          <div><strong>Department:</strong> ${escapeHtml(request.department || "-")}</div>
          <div><strong>Email:</strong> ${escapeHtml(request.email || "-")}</div>
          <div><strong>Phone:</strong> ${escapeHtml(request.phone || "-")}</div>
          <div><strong>Trip:</strong> ${escapeHtml(request.tripDate || "-")} ${escapeHtml(request.startTime || "")} - ${escapeHtml(request.endTime || "")}</div>
          <div><strong>Passengers:</strong> ${escapeHtml(request.passengers || "-")}</div>
          <div><strong>Pickup:</strong> ${escapeHtml(request.pickup || "-")}</div>
          <div><strong>Destination:</strong> ${escapeHtml(request.destination || "-")}</div>
          <div><strong>Purpose:</strong> ${escapeHtml(request.purpose || "-")}</div>
          <div><strong>Submitted:</strong> ${escapeHtml(request.submittedAt || "-")}</div>
        </div>
      `;

      if (status !== "APPROVED") {
        const actions = document.createElement("div");
        actions.className = "request-actions";
        const approveBtn = document.createElement("button");
        approveBtn.textContent = "Approve and notify";
        approveBtn.addEventListener("click", () => approveRequest(request.requestId, approveBtn));
        actions.appendChild(approveBtn);
        item.appendChild(actions);
      }

      listEl.appendChild(item);
    });
  }

  async function approveRequest(requestId, buttonEl) {
    const adminKey = getAdminKey();
    if (!adminKey) {
      showMessage("Admin key is required for approval.", "error");
      return;
    }

    const approvedBy = approvedByInput.value.trim() || "Fleet Admin";
    const confirmed = window.confirm("Approve request " + requestId + " and send email notification?");
    if (!confirmed) return;

    try {
      buttonEl.disabled = true;
      buttonEl.textContent = "Approving...";

      const response = await postToApi({
        action: "approveRequest",
        adminKey,
        requestId,
        approvedBy,
      });

      if (!response.ok) {
        throw new Error(response.error || "Could not approve this request.");
      }

      const emailStatus = response.emailStatus ? " Email: " + response.emailStatus : "";
      showMessage("Request " + requestId + " approved." + emailStatus, "success");
      await loadRequests();
    } catch (error) {
      showMessage(error.message || "Unexpected error while approving.", "error");
      buttonEl.disabled = false;
      buttonEl.textContent = "Approve and notify";
    }
  }

  function renderEmpty(text) {
    listEl.innerHTML = `<p class="empty-state">${escapeHtml(text)}</p>`;
    countEl.textContent = "0 records";
  }

  function getAdminKey() {
    return adminKeyInput.value.trim();
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
})();
