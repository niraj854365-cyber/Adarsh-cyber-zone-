const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxFlPsO4tMJML8eCipgbHC2HcHgpT3oHakE1d4k7_I_51L6f1AgYz9kF-N2mQKLpNQ5/exec";

const FALLBACK_LOGO = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'%3E%3Ccircle cx='50' cy='50' r='46' fill='%23ffb703'/%3E%3Ctext x='50' y='64' font-size='38' text-anchor='middle' fill='%23064b85' font-family='Arial' font-weight='700'%3EACZ%3C/text%3E%3C/svg%3E";

function $(id) {
    return document.getElementById(id);
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function initCommonUi() {
    document.querySelectorAll(".top-header img").forEach((img) => {
        img.addEventListener("error", () => {
            img.src = FALLBACK_LOGO;
        });
    });

    const current = location.pathname.split("/").pop() || "index.html";
    document.querySelectorAll("nav a[href]").forEach((link) => {
        if (link.getAttribute("href") === current) link.classList.add("active");
    });
}

function showMessage(elementId, message, isError = true, timeout = 4500) {
    const msgDiv = $(elementId);
    if (!msgDiv) return;
    msgDiv.textContent = message;
    msgDiv.style.display = "block";
    msgDiv.style.opacity = "1";
    msgDiv.className = isError ? "error-msg" : "success-msg";
    if (timeout) {
        window.clearTimeout(msgDiv._hideTimer);
        msgDiv._hideTimer = window.setTimeout(() => {
            msgDiv.style.opacity = "0";
            window.setTimeout(() => {
                msgDiv.style.display = "none";
                msgDiv.style.opacity = "1";
            }, 240);
        }, timeout);
    }
}

async function sendRequest(data) {
    try {
        const response = await fetch(GOOGLE_SCRIPT_URL, {
            method: "POST",
            body: JSON.stringify(data)
        });

        if (!response.ok) {
            throw new Error(`Server responded ${response.status}`);
        }

        return await response.json();
    } catch (error) {
        console.error("Request failed:", error);
        return { success: false, message: error.message || "Network error" };
    }
}

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

function setButtonLoading(button, loadingText) {
    if (!button) return () => {};
    const originalHtml = button.innerHTML;
    button.disabled = true;
    button.innerHTML = `<span class="loading"></span>${escapeHtml(loadingText)}`;
    return () => {
        button.disabled = false;
        button.innerHTML = originalHtml;
    };
}

function regNoFormat(fieldId) {
    const field = $(fieldId);
    if (!field) return;
    field.addEventListener("input", function () {
        let val = this.value.toUpperCase().replace(/[^ACZ0-9]/g, "");
        if (val.startsWith("ACZ")) {
            val = "ACZ" + val.slice(3).replace(/\D/g, "").slice(0, 7);
        } else {
            val = "ACZ" + val.replace(/\D/g, "").slice(0, 7);
        }
        this.value = val;
    });
}

function autoDOB(fieldId) {
    const field = $(fieldId);
    if (!field) return;
    field.addEventListener("input", function () {
        let val = this.value.replace(/\D/g, "");
        if (val.length >= 2) val = `${val.slice(0, 2)}-${val.slice(2)}`;
        if (val.length >= 5) val = `${val.slice(0, 5)}-${val.slice(5, 9)}`;
        this.value = val.slice(0, 10);
    });
}

function onlyCapsAndSpace(fieldId) {
    const field = $(fieldId);
    if (!field) return;
    field.addEventListener("input", function () {
        this.value = this.value.toUpperCase().replace(/[^A-Z\s]/g, "");
    });
}

function onlyDigits(fieldId, maxLength) {
    const field = $(fieldId);
    if (!field) return;
    field.addEventListener("input", function () {
        this.value = this.value.replace(/\D/g, "").slice(0, maxLength);
    });
}

function validateDob(value) {
    if (!/^\d{2}-\d{2}-\d{4}$/.test(value)) return false;
    const [dd, mm, yyyy] = value.split("-").map(Number);
    const date = new Date(yyyy, mm - 1, dd);
    return date.getFullYear() === yyyy && date.getMonth() === mm - 1 && date.getDate() === dd;
}

function formatDate(dateStr) {
    if (!dateStr) return "-";
    const formatDateObject = (date) => {
        const parts = new Intl.DateTimeFormat("en-GB", {
            timeZone: "Asia/Kolkata",
            day: "2-digit",
            month: "2-digit",
            year: "numeric"
        }).formatToParts(date);
        const get = (type) => parts.find((part) => part.type === type)?.value || "";
        return `${get("day")}-${get("month")}-${get("year")}`;
    };
    if (dateStr instanceof Date && !Number.isNaN(dateStr.getTime())) {
        return formatDateObject(dateStr);
    }
    const text = String(dateStr).trim();
    const ddmmyyyy = text.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})(?:\s+.*)?$/);
    if (ddmmyyyy) return `${ddmmyyyy[1]}-${ddmmyyyy[2]}-${ddmmyyyy[3]}`;
    const isoDateTime = text.match(/^(\d{4})-(\d{2})-(\d{2})T/);
    if (isoDateTime) {
        const parsedIso = new Date(text);
        if (!Number.isNaN(parsedIso.getTime())) {
            return formatDateObject(parsedIso);
        }
        return `${isoDateTime[3]}-${isoDateTime[2]}-${isoDateTime[1]}`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
        const [year, month, day] = text.split("-");
        return `${day}-${month}-${year}`;
    }
    const parsed = new Date(text);
    if (!Number.isNaN(parsed.getTime())) {
        return formatDateObject(parsed);
    }
    return text;
}

function formatDateTime(value) {
    if (!value) return "-";
    const formatDateObject = (date) => {
        const parts = new Intl.DateTimeFormat("en-GB", {
            timeZone: "Asia/Kolkata",
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
            hour12: false
        }).formatToParts(date);
        const get = (type) => parts.find((part) => part.type === type)?.value || "";
        return `${get("day")}-${get("month")}-${get("year")} ${get("hour")}:${get("minute")}:${get("second")}`;
    };
    if (value instanceof Date && !Number.isNaN(value.getTime())) return formatDateObject(value);
    const text = String(value).trim();
    if (/^\d{2}[-\/]\d{2}[-\/]\d{4}/.test(text)) return text.replaceAll("/", "-");
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? text : formatDateObject(parsed);
}

function getDriveFileId(url) {
    const text = String(url || "");
    return text.match(/\/d\/([^/]+)/)?.[1] || text.match(/[?&]id=([^&]+)/)?.[1] || "";
}

function driveDownloadUrl(url) {
    const id = getDriveFileId(url);
    return id ? `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}` : String(url || "");
}

function safeFileName(value, fallback = "document") {
    return String(value || fallback).replace(/[<>:"/\\|?*\x00-\x1F]+/g, "_").trim() || fallback;
}

function triggerDownload(url, fileName = "") {
    if (!url) return;
    const anchor = document.createElement("a");
    anchor.href = url;
    if (fileName) anchor.download = safeFileName(fileName);
    anchor.rel = "noopener";
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
}

function downloadDriveFile(url, fileName = "") {
    triggerDownload(driveDownloadUrl(url), fileName);
}

function downloadEncodedDriveFile(encodedUrl, fileName = "") {
    downloadDriveFile(decodeURIComponent(encodedUrl || ""), fileName);
}

async function downloadDriveFilePdf(url, fileName = "document.pdf", messageElementId) {
    const result = await sendRequest({
        action: "createFilePdf",
        url,
        fileName: safeFileName(fileName, "document.pdf")
    });
    if (result.success && (result.downloadUrl || result.link)) {
        triggerDownload(result.downloadUrl || driveDownloadUrl(result.link), result.fileName || fileName);
        return true;
    }
    if (messageElementId) showMessage(messageElementId, result.message || "Document PDF download failed.", true);
    return false;
}

function downloadEncodedDriveFilePdf(encodedUrl, fileName = "document.pdf", messageElementId) {
    downloadDriveFilePdf(decodeURIComponent(encodedUrl || ""), fileName, messageElementId);
}

async function downloadReceiptPdf(title, fileName, rows, messageElementId) {
    const result = await sendRequest({
        action: "createReceiptPdf",
        title,
        fileName: safeFileName(fileName, "receipt.pdf"),
        rows
    });
    if (result.success && (result.downloadUrl || result.link)) {
        triggerDownload(result.downloadUrl || driveDownloadUrl(result.link), result.fileName || fileName);
        return true;
    }
    if (messageElementId) showMessage(messageElementId, result.message || "PDF download failed.", true);
    return false;
}

let confirmCallback = null;

function openModal() {
    const modal = $("confirmModal");
    if (modal) modal.style.display = "flex";
}

function closeModal() {
    const modal = $("confirmModal");
    if (modal) modal.style.display = "none";
}

function confirmAction(message, callback) {
    if (window.confirm(message) && typeof callback === "function") callback();
}

function confirmYes() {
    const callback = confirmCallback;
    confirmCallback = null;
    closeModal();
    if (typeof callback === "function") callback();
}

function confirmNo() {
    confirmCallback = null;
    closeModal();
}

async function knowRegNo(event) {
    event?.preventDefault();
    const name = $("knowName")?.value.trim().toUpperCase();
    const father = $("knowFather")?.value.trim().toUpperCase();
    const dob = $("knowDob")?.value.trim();

    if (!name || !father || !dob) {
        showMessage("knowMsg", "कृपया नाम, पिता का नाम और जन्म तिथि भरें।", true);
        return;
    }

    if (!validateDob(dob)) {
        showMessage("knowMsg", "DOB सही DD-MM-YYYY फॉर्मेट में भरें।", true);
        return;
    }

    const restore = setButtonLoading(event?.target, "खोजा जा रहा है...");
    const result = await sendRequest({
        action: "knowRegNo",
        name,
        fatherName: father,
        dob
    });
    restore();

    if (result.success && (result.regNo || result.data?.regNo)) {
        showMessage("knowMsg", `आपका Registration No: ${result.regNo || result.data.regNo}`, false, 0);
    } else {
        showMessage("knowMsg", result.message || "रिकॉर्ड नहीं मिला।", true);
    }
}

document.addEventListener("DOMContentLoaded", initCommonUi);
