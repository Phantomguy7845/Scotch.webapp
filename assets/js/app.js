(function () {
  const config = window.APP_CONFIG || {};
  const form = document.getElementById("requestForm");
  const submitBtn = document.getElementById("submitBtn");
  const messageEl = document.getElementById("formMessage");

  if (!form) return;

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearMessage();

    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());

    if (!validateTime(payload.startTime, payload.endTime)) {
      showMessage("End time must be later than start time.", "error");
      return;
    }

    if (!isApiConfigured()) {
      showMessage("Set assets/js/config.js with a valid Apps Script URL first.", "error");
      return;
    }

    try {
      submitBtn.disabled = true;
      submitBtn.textContent = "Submitting...";

      const response = await postToApi({
        action: "submitRequest",
        data: payload,
      });

      if (!response.ok) {
        throw new Error(response.error || "Unable to submit request.");
      }

      showMessage(
        "Request submitted successfully. Request ID: " + response.requestId,
        "success",
      );
      form.reset();
    } catch (error) {
      showMessage(error.message || "Unexpected error while submitting.", "error");
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = "Submit request";
    }
  });

  function validateTime(startTime, endTime) {
    if (!startTime || !endTime) return true;
    return startTime < endTime;
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
})();
