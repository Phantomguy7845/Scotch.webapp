(function () {
  const config = window.APP_CONFIG || {};
  const DRAFT_KEY = "scotch_webapps1_request_draft_v2";
  const MAX_FILE_SIZE_BYTES = 3 * 1024 * 1024;
  const FILE_FIELDS = ["factoryTempDocumentImage", "factoryReservationImage"];
  const ALLOWED_FILE_TYPES = ["image/jpeg", "image/png", "image/webp"];
  const form = document.getElementById("requestForm");
  const submitBtn = document.getElementById("submitBtn");
  const clearDraftBtn = document.getElementById("clearDraftBtn");
  const messageEl = document.getElementById("formMessage");
  const offCycleBlock = document.getElementById("offCycleBlock");
  const factoryCaseBlock = document.getElementById("factoryCaseBlock");
  const factoryDeliveryBlock = document.getElementById("factoryDeliveryBlock");
  const factoryShuttleBlock = document.getElementById("factoryShuttleBlock");

  let draftTimer;

  if (!form) return;

  init();

  function init() {
    setMinDate();
    restoreDraft();
    updateConditionalBlocks();
    updateChoiceCards();

    form.addEventListener("change", onAnyFieldChanged);
    form.addEventListener("input", onAnyFieldChanged);
    form.addEventListener("submit", onSubmit);

    if (clearDraftBtn) {
      clearDraftBtn.addEventListener("click", () => {
        form.reset();
        clearAllFieldErrors();
        clearMessage();
        clearDraft();
        updateConditionalBlocks();
        updateChoiceCards();
      });
    }
  }

  async function onSubmit(event) {
    event.preventDefault();
    clearMessage();
    clearAllFieldErrors();

    const rawData = collectPayload();
    const validation = validatePayload(rawData);
    if (validation) {
      showMessage(validation.message, "error");
      return;
    }

    if (!isApiConfigured()) {
      showMessage("Set assets/js/config.js with a valid Apps Script URL first.", "error");
      return;
    }

    try {
      submitBtn.disabled = true;
      submitBtn.textContent = "Submitting...";

      const payload = await buildPayload(rawData);
      const response = await postToApi({
        action: "submitRequest",
        data: payload,
      });

      if (!response.ok) {
        throw new Error(response.error || "Unable to submit request.");
      }

      showMessage("ส่งคำร้องสำเร็จ รหัสคำขอ: " + response.requestId, "success");
      form.reset();
      clearDraft();
      updateConditionalBlocks();
      updateChoiceCards();
    } catch (error) {
      showMessage(error.message || "เกิดข้อผิดพลาดระหว่างส่งคำร้อง", "error");
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = "ส่งคำร้อง";
    }
  }

  function onAnyFieldChanged(event) {
    if (event && event.target && event.target.name) {
      clearFieldError(event.target.name);
    }
    updateConditionalBlocks();
    updateChoiceCards();
    scheduleDraftSave();
  }

  function collectPayload() {
    const data = {};
    const formData = new FormData(form);
    for (const [key, value] of formData.entries()) {
      if (value instanceof File) {
        data[key] = value.size > 0 ? value : null;
      } else {
        data[key] = String(value || "").trim();
      }
    }

    data.policyAgree = form.elements.policyAgree.checked ? "YES" : "NO";
    data.requestTopicLabel = getTopicLabel(data.requestTopic);
    data.factoryCaseLabel = getFactoryCaseLabel(data.factoryCaseType);

    if (data.requestTopic !== "OFF_CYCLE_DELIVERY") {
      clearKeys(data, [
        "offcycleStoreName",
        "offcycleDeliveryDate",
        "offcycleCrates",
        "offcycleAmount",
      ]);
    }

    if (data.requestTopic !== "FACTORY_CAR") {
      clearKeys(data, [
        "factoryCaseType",
        "factoryCaseLabel",
        "factoryJobName",
        "factoryDeliveryDate",
        "factoryDeliveryTime",
        "factoryDeliveryLocation",
        "factoryDeliveryMapUrl",
        "factoryReceiverPhone",
        "factoryTempDocumentImage",
        "factoryTempDocumentRef",
        "factoryReservationImage",
        "factoryReservationRef",
        "shuttlePassengers",
        "shuttleTravelDate",
        "shuttleTravelTime",
        "shuttleReturnWait",
        "shuttleLocation",
        "shuttleMapUrl",
      ]);
    }

    if (data.factoryCaseType !== "FACTORY_DELIVERY_DOCS") {
      clearKeys(data, [
        "factoryJobName",
        "factoryDeliveryDate",
        "factoryDeliveryTime",
        "factoryDeliveryLocation",
        "factoryDeliveryMapUrl",
        "factoryReceiverPhone",
        "factoryTempDocumentImage",
        "factoryTempDocumentRef",
        "factoryReservationImage",
        "factoryReservationRef",
      ]);
    }

    if (data.factoryCaseType !== "FACTORY_STAFF_SHUTTLE") {
      clearKeys(data, [
        "shuttlePassengers",
        "shuttleTravelDate",
        "shuttleTravelTime",
        "shuttleReturnWait",
        "shuttleLocation",
        "shuttleMapUrl",
      ]);
    }

    return data;
  }

  async function buildPayload(rawData) {
    const payload = { ...rawData };
    for (const fieldName of FILE_FIELDS) {
      payload[fieldName] = await toAttachmentPayload(rawData[fieldName]);
    }
    return payload;
  }

  function validatePayload(data) {
    if (!data.requestTopic) {
      setFieldError("requestTopic", "กรุณาเลือกหัวข้อคำขอ");
      return { message: "กรุณาเลือกหัวข้อคำขอ" };
    }

    if (!data.requesterName) {
      setFieldError("requesterName", "กรุณากรอกชื่อผู้ขอใช้");
      return { message: "กรุณากรอกชื่อผู้ขอใช้" };
    }

    if (!data.requesterPhone) {
      setFieldError("requesterPhone", "กรุณากรอกเบอร์ติดต่อ");
      return { message: "กรุณากรอกเบอร์ติดต่อผู้ขอใช้" };
    }
    if (!isValidPhone(data.requesterPhone)) {
      setFieldError("requesterPhone", "รูปแบบเบอร์โทรไม่ถูกต้อง");
      return { message: "รูปแบบเบอร์ติดต่อไม่ถูกต้อง" };
    }

    if (!data.contactEmail) {
      setFieldError("contactEmail", "กรุณากรอกอีเมลติดต่อ");
      return { message: "กรุณากรอกอีเมลติดต่อ" };
    }
    if (!isValidEmail(data.contactEmail)) {
      setFieldError("contactEmail", "รูปแบบอีเมลไม่ถูกต้อง");
      return { message: "รูปแบบอีเมลไม่ถูกต้อง" };
    }

    if (data.requestTopic === "OFF_CYCLE_DELIVERY") {
      if (!data.offcycleStoreName) return fieldRequired("offcycleStoreName", "กรุณากรอกชื่อห้าง");
      if (!data.offcycleDeliveryDate)
        return fieldRequired("offcycleDeliveryDate", "กรุณาเลือกวันที่ต้องการจัดส่ง");
      if (isPastDate(data.offcycleDeliveryDate)) {
        setFieldError("offcycleDeliveryDate", "วันที่จัดส่งต้องไม่ย้อนหลัง");
        return { message: "วันที่จัดส่งต้องไม่ย้อนหลัง" };
      }
      if (!data.offcycleCrates) return fieldRequired("offcycleCrates", "กรุณากรอกจำนวนลังสินค้า");
      if (!isIntegerAtLeast(data.offcycleCrates, 1)) {
        setFieldError("offcycleCrates", "จำนวนลังต้องมากกว่า 0");
        return { message: "จำนวนลังสินค้าไม่ถูกต้อง" };
      }
      if (!data.offcycleAmount) return fieldRequired("offcycleAmount", "กรุณากรอกยอดเงิน");
      if (!isNumberAtLeast(data.offcycleAmount, 0)) {
        setFieldError("offcycleAmount", "ยอดเงินต้องไม่ติดลบ");
        return { message: "ยอดเงินไม่ถูกต้อง" };
      }
    }

    if (data.requestTopic === "FACTORY_CAR") {
      if (!data.factoryCaseType) {
        setFieldError("factoryCaseType", "กรุณาเลือกกรณีของรถโรงงาน");
        return { message: "กรุณาเลือกกรณีของรถโรงงาน" };
      }

      if (data.factoryCaseType === "FACTORY_DELIVERY_DOCS") {
        if (!data.factoryJobName) return fieldRequired("factoryJobName", "กรุณากรอกชื่องาน");
        if (!data.factoryDeliveryDate)
          return fieldRequired("factoryDeliveryDate", "กรุณาเลือกวันที่จัดส่ง");
        if (isPastDate(data.factoryDeliveryDate)) {
          setFieldError("factoryDeliveryDate", "วันที่จัดส่งต้องไม่ย้อนหลัง");
          return { message: "วันที่จัดส่งต้องไม่ย้อนหลัง" };
        }
        if (!data.factoryDeliveryTime)
          return fieldRequired("factoryDeliveryTime", "กรุณาเลือกเวลาจัดส่ง");
        if (!data.factoryDeliveryLocation)
          return fieldRequired("factoryDeliveryLocation", "กรุณากรอกสถานที่จัดส่ง");
        if (!data.factoryDeliveryMapUrl)
          return fieldRequired("factoryDeliveryMapUrl", "กรุณากรอก Google Map URL");
        if (!isValidUrl(data.factoryDeliveryMapUrl)) {
          setFieldError("factoryDeliveryMapUrl", "URL ไม่ถูกต้อง");
          return { message: "Google Map URL ไม่ถูกต้อง" };
        }
        if (!data.factoryReceiverPhone)
          return fieldRequired("factoryReceiverPhone", "กรุณากรอกเบอร์ติดต่อผู้รับ/หน้างาน");
        if (!isValidPhone(data.factoryReceiverPhone)) {
          setFieldError("factoryReceiverPhone", "รูปแบบเบอร์โทรไม่ถูกต้อง");
          return { message: "เบอร์ติดต่อผู้รับ/หน้างานไม่ถูกต้อง" };
        }
        if (!isValidUploadFile(data.factoryTempDocumentImage)) {
          setFieldError("factoryTempDocumentImage", "กรุณาแนบรูปเอกสารให้ถูกต้อง (png/jpg/webp ไม่เกิน 3MB)");
          return { message: "กรุณาแนบรูปเอกสารชั่วคราวให้ถูกต้อง" };
        }
        if (!isValidUploadFile(data.factoryReservationImage)) {
          setFieldError("factoryReservationImage", "กรุณาแนบรูป Reservation ให้ถูกต้อง (png/jpg/webp ไม่เกิน 3MB)");
          return { message: "กรุณาแนบรูป Reservation ให้ถูกต้อง" };
        }
      }

      if (data.factoryCaseType === "FACTORY_STAFF_SHUTTLE") {
        if (!data.shuttlePassengers)
          return fieldRequired("shuttlePassengers", "กรุณากรอกจำนวนผู้โดยสาร");
        if (!isIntegerAtLeast(data.shuttlePassengers, 6)) {
          setFieldError("shuttlePassengers", "ต้อง 6 คนขึ้นไป");
          return { message: "กรณีรับ-ส่งพนักงาน ต้องมีผู้โดยสารอย่างน้อย 6 คน" };
        }
        if (!data.shuttleTravelDate)
          return fieldRequired("shuttleTravelDate", "กรุณาเลือกวันที่เดินทาง");
        if (isPastDate(data.shuttleTravelDate)) {
          setFieldError("shuttleTravelDate", "วันที่เดินทางต้องไม่ย้อนหลัง");
          return { message: "วันที่เดินทางต้องไม่ย้อนหลัง" };
        }
        if (!data.shuttleTravelTime)
          return fieldRequired("shuttleTravelTime", "กรุณาเลือกเวลาเดินทาง");
        if (!data.shuttleReturnWait)
          return fieldRequired("shuttleReturnWait", "กรุณาเลือกรอรับกลับหรือไม่");
        if (!data.shuttleLocation) return fieldRequired("shuttleLocation", "กรุณากรอกสถานที่");
        if (!data.shuttleMapUrl) return fieldRequired("shuttleMapUrl", "กรุณากรอก Google Map URL");
        if (!isValidUrl(data.shuttleMapUrl)) {
          setFieldError("shuttleMapUrl", "URL ไม่ถูกต้อง");
          return { message: "Google Map URL ไม่ถูกต้อง" };
        }
      }
    }

    if (data.policyAgree !== "YES") {
      setFieldError("policyAgree", "กรุณายืนยันข้อมูลก่อนส่งคำร้อง");
      return { message: "กรุณายืนยันข้อมูลก่อนส่งคำร้อง" };
    }

    return null;
  }

  function updateConditionalBlocks() {
    const topic = form.elements.requestTopic.value;
    const caseType = form.elements.factoryCaseType.value;

    setBlockVisible(factoryCaseBlock, topic === "FACTORY_CAR");
    setBlockVisible(offCycleBlock, topic === "OFF_CYCLE_DELIVERY");
    setBlockVisible(factoryDeliveryBlock, topic === "FACTORY_CAR" && caseType === "FACTORY_DELIVERY_DOCS");
    setBlockVisible(factoryShuttleBlock, topic === "FACTORY_CAR" && caseType === "FACTORY_STAFF_SHUTTLE");
  }

  function setBlockVisible(blockEl, visible) {
    if (!blockEl) return;
    blockEl.classList.toggle("hidden", !visible);
  }

  function updateChoiceCards() {
    const cards = form.querySelectorAll(".choice-card");
    cards.forEach((card) => {
      const input = card.querySelector("input[type='radio']");
      if (!input) return;
      card.classList.toggle("active", input.checked);
    });
  }

  function setFieldError(name, text) {
    const errorEl = form.querySelector('[data-error-for="' + name + '"]');
    if (errorEl) errorEl.textContent = text;

    const controls = form.querySelectorAll('[name="' + name + '"]');
    if (!controls.length) return;

    controls.forEach((control) => {
      control.classList.add("input-invalid");
      control.setAttribute("aria-invalid", "true");
    });

    if (controls[0] && typeof controls[0].focus === "function") {
      controls[0].focus();
    }
  }

  function clearFieldError(name) {
    const errorEl = form.querySelector('[data-error-for="' + name + '"]');
    if (errorEl) errorEl.textContent = "";

    const controls = form.querySelectorAll('[name="' + name + '"]');
    controls.forEach((control) => {
      control.classList.remove("input-invalid");
      control.removeAttribute("aria-invalid");
    });
  }

  function clearAllFieldErrors() {
    const errorEls = form.querySelectorAll(".field-error");
    errorEls.forEach((el) => {
      el.textContent = "";
    });
    const invalidEls = form.querySelectorAll(".input-invalid");
    invalidEls.forEach((el) => {
      el.classList.remove("input-invalid");
      el.removeAttribute("aria-invalid");
    });
  }

  function fieldRequired(name, message) {
    setFieldError(name, message);
    return { message };
  }

  function clearKeys(data, keys) {
    keys.forEach((key) => {
      data[key] = FILE_FIELDS.includes(key) ? null : "";
    });
  }

  function getTopicLabel(topic) {
    if (topic === "OFF_CYCLE_DELIVERY") return "ขอส่งสินค้าไม่ตรงรอบ";
    if (topic === "FACTORY_CAR") return "ขอใช้รถโรงงาน";
    return "";
  }

  function getFactoryCaseLabel(caseType) {
    if (caseType === "FACTORY_DELIVERY_DOCS") return "ส่งสินค้า/เอกสาร";
    if (caseType === "FACTORY_STAFF_SHUTTLE") return "รับ-ส่งพนักงาน";
    return "";
  }

  function scheduleDraftSave() {
    window.clearTimeout(draftTimer);
    draftTimer = window.setTimeout(saveDraft, 250);
  }

  function saveDraft() {
    const draft = {};
    Array.from(form.elements).forEach((el) => {
      if (!el.name) return;
      if (el.type === "file") return;
      if (el.type === "radio") {
        if (el.checked) draft[el.name] = el.value;
      } else if (el.type === "checkbox") {
        draft[el.name] = el.checked;
      } else {
        draft[el.name] = el.value;
      }
    });
    localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
  }

  function restoreDraft() {
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (!raw) return;
      const draft = JSON.parse(raw);
      Object.keys(draft).forEach((name) => {
        const controls = form.querySelectorAll('[name="' + name + '"]');
        if (!controls.length) return;

        controls.forEach((control) => {
          if (control.type === "radio") {
            control.checked = control.value === draft[name];
          } else if (control.type === "checkbox") {
            control.checked = Boolean(draft[name]);
          } else if (control.type === "file") {
            control.value = "";
          } else {
            control.value = draft[name] || "";
          }
        });
      });
    } catch (error) {
      localStorage.removeItem(DRAFT_KEY);
    }
  }

  function clearDraft() {
    localStorage.removeItem(DRAFT_KEY);
  }

  function setMinDate() {
    const today = todayIso();
    const dateFields = form.querySelectorAll("input[type='date']");
    dateFields.forEach((field) => {
      field.min = today;
    });
  }

  function todayIso() {
    const now = new Date();
    const tzOffset = now.getTimezoneOffset() * 60000;
    return new Date(now.getTime() - tzOffset).toISOString().slice(0, 10);
  }

  function isPastDate(dateText) {
    return dateText < todayIso();
  }

  function isValidEmail(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  }

  function isValidPhone(phone) {
    const digits = String(phone || "").replace(/\D/g, "");
    return digits.length >= 8 && digits.length <= 15;
  }

  function isIntegerAtLeast(value, min) {
    const num = Number(value);
    return Number.isInteger(num) && num >= min;
  }

  function isNumberAtLeast(value, min) {
    const num = Number(value);
    return Number.isFinite(num) && num >= min;
  }

  function isValidUrl(url) {
    try {
      const parsed = new URL(url);
      return parsed.protocol === "http:" || parsed.protocol === "https:";
    } catch (error) {
      return false;
    }
  }

  function isValidUploadFile(file) {
    if (!(file instanceof File)) return false;
    if (!ALLOWED_FILE_TYPES.includes(file.type)) return false;
    if (file.size <= 0 || file.size > MAX_FILE_SIZE_BYTES) return false;
    return true;
  }

  async function toAttachmentPayload(file) {
    if (!(file instanceof File) || file.size === 0) return null;
    const base64 = await readFileAsBase64(file);
    return {
      name: file.name,
      mimeType: file.type,
      base64: base64,
      size: file.size,
    };
  }

  function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = String(reader.result || "");
        const base64 = result.split(",")[1] || "";
        if (!base64) {
          reject(new Error("Cannot read selected file."));
          return;
        }
        resolve(base64);
      };
      reader.onerror = () => reject(new Error("Cannot read selected file."));
      reader.readAsDataURL(file);
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
})();
