const SHEETS = {
  USERS: 'Users',
  CERTIFICATES: 'Certificates',
  ADMINS: 'Admins',
  PAYMENT: 'Payment',
  USER_REQUEST: 'UserRequest'
};

const USER_HEADERS = [
  'RegNo', 'ApplicantNmaeEng', 'ApplicantNmaeHindi', 'FatherNameEng', 'FatherNameHindi',
  'MotherNameEng', 'MotherNameHindi', 'HusbandNameEng', 'HusbandNameHindi', 'MaritalStatus',
  'DOB', 'Gender', 'CasteCategory', 'CasteName', 'Profession', 'TotalAnnualIncome',
  'MobileNumber', 'Email', 'FullPresentAddress', 'FullPermanentAddress', 'PhotoLink',
  'SignatureLink', 'AadhaarLink', 'OtherDocumentLink', 'RegistrationSlipLink', 'Status', 'DateOfRegistration'
];

const CERT_HEADERS = ['RegNo', 'CertificateName', 'CertificateType', 'UploadDate', 'CertificateLink'];
const ADMIN_HEADERS = ['AdminID', 'Password'];
const PAYMENT_HEADERS = ['RegNo', 'ApplicantNmae', 'TransactionId', 'PaymentSlipLink', 'PaymentDate'];
const REQUEST_HEADERS = ['RegNo', 'ApplicantNmae', 'DOB', 'MobileNumber', 'RequestFor', 'RequestDate', 'Status'];

function doGet() {
  return jsonResponse({ success: true, message: 'Adarsh Cyber Zone API is running.' });
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = data.action;
    setupSheets();

    switch (action) {
      case 'register': return jsonResponse(registerUser(data));
      case 'getUser': return jsonResponse(getUser(data));
      case 'getAllUsers': return jsonResponse(getAllUsers());
      case 'knowRegNo': return jsonResponse(knowRegNo(data));
      case 'updateProfile': return jsonResponse(updateUser(data));
      case 'adminUpdateUser': return jsonResponse(updateUser(data));
      case 'updateUserDocuments': return jsonResponse(updateUserDocuments(data));
      case 'adminLogin': return jsonResponse(adminLogin(data));
      case 'getAdmins': return jsonResponse(getAdmins());
      case 'updateAdmin': return jsonResponse(updateAdmin(data));
      case 'uploadCertificate': return jsonResponse(uploadCertificate(data));
      case 'getCertificates': return jsonResponse(getCertificates(data));
      case 'deleteCertificate': return jsonResponse(deleteByRegAndLink(SHEETS.CERTIFICATES, data));
      case 'submitPayment': return jsonResponse(submitPayment(data));
      case 'getPayments': return jsonResponse(getPayments(data));
      case 'deletePayment': return jsonResponse(deleteByRegAndLink(SHEETS.PAYMENT, data));
      case 'submitUserRequest': return jsonResponse(submitUserRequest(data));
      case 'getUserRequests': return jsonResponse(getUserRequests(data));
      case 'updateUserRequestStatus': return jsonResponse(updateUserRequestStatus(data));
      case 'updateStatus': return jsonResponse(updateStatus(data));
      case 'createReceiptPdf': return jsonResponse(createReceiptPdf(data));
      case 'createFilePdf': return jsonResponse(createFilePdf(data));
      case 'deleteUser':
      case 'removeUser': return jsonResponse(deleteUser(data));
      default: return jsonResponse({ success: false, message: 'Invalid action.' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function setupSheets() {
  ensureSheet(SHEETS.USERS, USER_HEADERS);
  ensureSheet(SHEETS.CERTIFICATES, CERT_HEADERS);
  ensureSheet(SHEETS.ADMINS, ADMIN_HEADERS);
  ensureSheet(SHEETS.PAYMENT, PAYMENT_HEADERS);
  ensureSheet(SHEETS.USER_REQUEST, REQUEST_HEADERS);
}

function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  const existing = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].filter(String);
  headers.forEach(h => {
    if (existing.indexOf(h) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      existing.push(h);
    }
  });
}

function registerUser(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet(SHEETS.USERS);
   let regNo;

do {
  regNo = nextRegNo();
} while (findRowByValue(sheet, 'RegNo', regNo).row > 0);
   

    const row = {};
    USER_HEADERS.forEach(h => row[h] = data[h] || '');
    row.RegNo = regNo;
    row.ApplicantNmaeEng = data.ApplicantNmaeEng || data.ApplicantNameEng || '';
    row.ApplicantNmaeHindi = data.ApplicantNmaeHindi || data.ApplicantNameHindi || '';
    row.Status = data.Status || 'PENDING';
    row.DateOfRegistration = formatDateOnly(new Date());

    row.PhotoLink = saveBase64File(data.photoBase64, data.photoFileName, regNo, 'Photo') || '';
    row.SignatureLink = saveBase64File(data.signatureBase64, data.signatureFileName, regNo, 'Signature') || '';
    row.AadhaarLink = saveBase64File(data.aadhaarBase64, data.aadhaarFileName, regNo, 'Aadhaar') || '';
    row.OtherDocumentLink = saveBase64File(data.otherDocumentBase64, data.otherDocumentFileName, regNo, 'OtherDocument') || '';
    row.RegistrationSlipLink = createRegistrationSlip(row);

    appendObject(sheet, USER_HEADERS, row);
    return { success: true, regNo, RegNo: regNo, data: row };
  } finally {
    lock.releaseLock();
  }
}

function getUser(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  const dob = normalizeDob(data.DOB || data.dob);
  const found = findUser(regNo);
  if (!found) return { success: false, message: 'User not found.' };
  if (dob && normalizeDob(found.data.DOB) !== dob) return { success: false, message: 'DOB does not match.' };
  found.data.regNo = found.data.RegNo;
  found.data.DOB = normalizeDob(found.data.DOB);
  return { success: true, data: found.data };
}

function getAllUsers() {
  const users = getObjects(getSheet(SHEETS.USERS)).map(u => {
    u.regNo = u.RegNo;
    u.createdDate = '';
    u.DOB = normalizeDob(u.DOB);
    return u;
  });
  return { success: true, users };
}

function knowRegNo(data) {
  const name = String(data.name || data.ApplicantNmaeEng || '').trim().toUpperCase();
  const father = String(data.fatherName || data.FatherNameEng || '').trim().toUpperCase();
  const dob = normalizeDob(data.dob || data.DOB);
  const users = getObjects(getSheet(SHEETS.USERS));
  const user = users.find(u =>
    String(u.ApplicantNmaeEng || '').trim().toUpperCase() === name &&
    String(u.FatherNameEng || '').trim().toUpperCase() === father &&
    normalizeDob(u.DOB) === dob
  );
  if (user) user.DOB = normalizeDob(user.DOB);
  return user ? { success: true, regNo: user.RegNo, data: user } : { success: false, message: 'Record not found.' };
}

function updateUser(data) {
  const found = findUser(data.RegNo || data.regNo);
  if (!found) return { success: false, message: 'User not found.' };
  const updates = {};
  USER_HEADERS.forEach(h => {
    if (h !== 'RegNo' && data[h] !== undefined) updates[h] = data[h];
  });
  if (data.ApplicantNameEng !== undefined) updates.ApplicantNmaeEng = data.ApplicantNameEng;
  if (updates.DOB !== undefined) updates.DOB = normalizeDob(updates.DOB);
  updateRow(getSheet(SHEETS.USERS), found.row, updates);
  const updated = findUser(data.RegNo || data.regNo).data;
  updated.regNo = updated.RegNo;
  updated.DOB = normalizeDob(updated.DOB);
  return { success: true, data: updated };
}

function updateUserDocuments(data) {
  const found = findUser(data.RegNo || data.regNo);
  if (!found) return { success: false, message: 'User not found.' };
  const regNo = found.data.RegNo;
  const updates = {};
  const photo = saveBase64File(data.photoBase64, data.photoFileName, regNo, 'Photo');
  const sign = saveBase64File(data.signatureBase64, data.signatureFileName, regNo, 'Signature');
  const aadhaar = saveBase64File(data.aadhaarBase64, data.aadhaarFileName, regNo, 'Aadhaar');
  const other = saveBase64File(data.otherDocumentBase64, data.otherDocumentFileName, regNo, 'OtherDocument');
  if (photo) updates.PhotoLink = photo;
  if (sign) updates.SignatureLink = sign;
  if (aadhaar) updates.AadhaarLink = aadhaar;
  if (other) updates.OtherDocumentLink = other;
  updateRow(getSheet(SHEETS.USERS), found.row, updates);
  return { success: true, links: updates };
}

function adminLogin(data) {
  const adminID = String(data.AdminID || data.adminID || '').trim();
  const password = String(data.Password || data.password || '').trim();
  const admins = getObjects(getSheet(SHEETS.ADMINS));
  const ok = admins.some(a => String(a.AdminID).trim() === adminID && String(a.Password).trim() === password);
  return ok ? { success: true } : { success: false, message: 'Invalid Admin ID or Password.' };
}

function getAdmins() {
  return { success: true, admins: getObjects(getSheet(SHEETS.ADMINS)) };
}

function updateAdmin(data) {
  const adminID = String(data.AdminID || data.adminID || '').trim();
  const password = String(data.Password || data.password || '').trim();
  if (!adminID || !password) return { success: false, message: 'Admin ID and Password are required.' };
  const sheet = getSheet(SHEETS.ADMINS);
  if (sheet.getLastRow() < 2) {
    appendObject(sheet, ADMIN_HEADERS, { AdminID: adminID, Password: password });
  } else {
    sheet.getRange(2, 1).setValue(adminID);
    sheet.getRange(2, 2).setValue(password);
  }
  return { success: true };
}

function uploadCertificate(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  if (!findUser(regNo)) return { success: false, message: 'User not found.' };
  const link = saveBase64File(data.fileBase64, data.fileName, regNo, 'Certificate');
  const row = {
    RegNo: regNo,
    CertificateName: data.CertificateName || data.certificateName || 'Document',
    CertificateType: data.CertificateType || data.certificateType || 'सर्टिफिकेट',
    UploadDate: formatDateOnly(new Date()),
    CertificateLink: link
  };
  appendObject(getSheet(SHEETS.CERTIFICATES), CERT_HEADERS, row);
  return { success: true, certificate: row };
}

function getCertificates(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  const certificates = getObjects(getSheet(SHEETS.CERTIFICATES))
    .filter(c => String(c.RegNo) === regNo)
    .map(c => {
      c.UploadDate = normalizeDateTime(c.UploadDate);
      return c;
    });
  return { success: true, certificates };
}

function submitPayment(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  if (!findUser(regNo)) return { success: false, message: 'User not found.' };
  const link = saveBase64File(data.paymentSlipBase64 || data.fileBase64, data.paymentSlipFileName || data.fileName, regNo, 'PaymentSlip');
  const row = {
    RegNo: regNo,
    ApplicantNmae: data.ApplicantNmae || data.applicantName || '',
    TransactionId: data.TransactionId || data.transactionId || '',
    PaymentSlipLink: link || data.PaymentSlipLink || '',
    PaymentDate: formatDateOnly(new Date())
  };
  appendObject(getSheet(SHEETS.PAYMENT), PAYMENT_HEADERS, row);
  return { success: true, payment: row };
}

function getPayments(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  const payments = getObjects(getSheet(SHEETS.PAYMENT)).filter(p => String(p.RegNo) === regNo);
  return { success: true, payments };
}

function updateStatus(data) {
  const found = findUser(data.RegNo || data.regNo);
  if (!found) return { success: false, message: 'User not found.' };
  updateRow(getSheet(SHEETS.USERS), found.row, { Status: data.Status || data.status || 'PENDING' });
  return { success: true };
}

function submitUserRequest(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  const name = String(data.ApplicantNmae || data.applicantName || '').trim().toUpperCase();
  const dob = normalizeDob(data.DOB || data.dob);
  const mobile = String(data.MobileNumber || data.mobileNumber || '').trim();
  const requestFor = String(data.RequestFor || data.requestFor || '').trim();
  if (!regNo || !name || !dob || !mobile || !requestFor) {
    return { success: false, message: 'All request fields are required.' };
  }
  const found = findUser(regNo);
  if (!found) return { success: false, message: 'Registration number Users sheet me nahi mila.' };
  const user = found.data;
  const userName = String(user.ApplicantNmaeEng || user.ApplicantNameEng || '').trim().toUpperCase();
  if (userName !== name || normalizeDob(user.DOB) !== dob || String(user.MobileNumber || '').trim() !== mobile) {
    return { success: false, message: 'User details Users sheet se match nahi kar raha hai.' };
  }
  const payments = getObjects(getSheet(SHEETS.PAYMENT));
  const hasPayment = payments.some(p =>
    normalizeRegNo(p.RegNo) === regNo &&
    String(p.ApplicantNmae || '').trim().toUpperCase() === name
  );
  if (!hasPayment) return { success: false, message: 'Payment sheet me Reg No aur name match nahi mila.' };
  const row = {
    RegNo: regNo,
    ApplicantNmae: user.ApplicantNmaeEng || data.ApplicantNmae || '',
    DOB: dob,
    MobileNumber: mobile,
    RequestFor: requestFor,
    RequestDate: formatDateOnly(new Date()),
    Status: 'PENDING'
  };
  appendObject(getSheet(SHEETS.USER_REQUEST), REQUEST_HEADERS, row);
  return { success: true, request: row };
}

function getUserRequests(data) {
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  const requests = getObjectsWithRow(getSheet(SHEETS.USER_REQUEST))
    .filter(r => !regNo || normalizeRegNo(r.RegNo) === regNo)
    .map(r => {
      r.DOB = normalizeDob(r.DOB);
      r.RequestDate = normalizeDateOnly(r.RequestDate);
      return r;
    });
  return { success: true, requests };
}

function updateUserRequestStatus(data) {
  const rowNumber = Number(data.rowNumber || data.RowNumber || data._row);
  const status = String(data.Status || data.status || '').trim().toUpperCase();
  if (!rowNumber || rowNumber < 2) return { success: false, message: 'Request row not found.' };
  if (!['PENDING', 'ASSIGNED', 'COMPLETED'].includes(status)) return { success: false, message: 'Invalid request status.' };
  const sheet = getSheet(SHEETS.USER_REQUEST);
  if (rowNumber > sheet.getLastRow()) return { success: false, message: 'Request row not found.' };
  updateRow(sheet, rowNumber, { Status: status });
  return { success: true };
}

function deleteUser(data) {
  const found = findUser(data.RegNo || data.regNo);
  if (!found) return { success: false, message: 'User not found.' };
  getSheet(SHEETS.USERS).deleteRow(found.row);
  return { success: true };
}

function deleteByRegAndLink(sheetName, data) {
  const sheet = getSheet(sheetName);
  const regNo = normalizeRegNo(data.RegNo || data.regNo);
  const link = data.link || data.CertificateLink || data.PaymentSlipLink || '';
  const objects = getObjects(sheet);
  for (let i = 0; i < objects.length; i++) {
    if (String(objects[i].RegNo) === regNo && (!link || Object.values(objects[i]).indexOf(link) !== -1)) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, message: 'Record not found.' };
}

function findUser(regNo) {
  const sheet = getSheet(SHEETS.USERS);
  const found = findRowByValue(sheet, 'RegNo', normalizeRegNo(regNo));
  if (found.row < 1) return null;
  return { row: found.row, data: rowToObject(sheet, found.row) };
}

function nextRegNo() {
  const sheet = getSheet(SHEETS.USERS);

  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();

  let maxNo = 0;

  data.forEach(row => {
    const reg = String(row[0] || "").trim();
    const match = reg.match(/^ACZ(\d+)$/);

    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNo) {
        maxNo = num;
      }
    }
  });

  return "ACZ" + String(maxNo + 1).padStart(7, "0");
}

function normalizeRegNo(value) {
  return String(value || '').trim().toUpperCase();
}

function normalizeDob(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Kolkata', 'dd-MM-yyyy');
  }
  const text = String(value).trim();
  let match = text.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})$/);
  if (match) return `${match[1]}-${match[2]}-${match[3]}`;
  match = text.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})/);
  if (match) return `${match[3]}-${match[2]}-${match[1]}`;
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'Asia/Kolkata', 'dd-MM-yyyy');
  }
  return text;
}

function normalizeDateTime(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Kolkata', 'dd-MM-yyyy HH:mm:ss');
  }
  const text = String(value).trim();
  if (/^\d{2}[-\/]\d{2}[-\/]\d{4}/.test(text)) return text.replace(/\//g, '-');
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, 'Asia/Kolkata', 'dd-MM-yyyy HH:mm:ss');
  }
  return text;
}

function formatDateOnly(value) {
  return Utilities.formatDate(value || new Date(), 'Asia/Kolkata', 'dd-MM-yyyy');
}

function normalizeDateOnly(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDateOnly(value);
  }
  const text = String(value).trim();
  if (/^\d{2}[-\/]\d{2}[-\/]\d{4}/.test(text)) return text.slice(0, 10).replace(/\//g, '-');
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return formatDateOnly(parsed);
  return text;
}

function saveBase64File(base64, fileName, regNo, type) {
  if (!base64) return '';
  const folder = getUploadFolder();
  const match = String(base64).match(/^data:([^;]+);base64,(.+)$/);
  const mime = match ? match[1] : 'application/octet-stream';
  const bytes = Utilities.base64Decode(match ? match[2] : base64);
  const safeName = `${regNo}_${type}_${Date.now()}_${fileName || 'file'}`.replace(/[^\w.\-]+/g, '_');
  const file = folder.createFile(Utilities.newBlob(bytes, mime, safeName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function createRegistrationSlip(user) {
  const html = [
    '<h1>Adarsh Cyber Zone</h1>',
    '<h2>Registration Slip</h2>',
    `<p><b>Reg No:</b> ${escapeHtmlGs(user.RegNo)}</p>`,
    `<p><b>Name:</b> ${escapeHtmlGs(user.ApplicantNmaeEng)}</p>`,
    `<p><b>DOB:</b> ${escapeHtmlGs(user.DOB)}</p>`,
    `<p><b>Mobile:</b> ${escapeHtmlGs(user.MobileNumber)}</p>`,
    `<p><b>Status:</b> ${escapeHtmlGs(user.Status)}</p>`
  ].join('');
  const file = getUploadFolder().createFile(`${user.RegNo}_RegistrationSlip.html`, html, MimeType.HTML);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function createReceiptPdf(data) {
  const title = String(data.title || 'Receipt').trim();
  const rows = Array.isArray(data.rows) ? data.rows : [];
  const requestedName = String(data.fileName || `${title}.pdf`).replace(/\.pdf$/i, '');
  const safeName = requestedName.replace(/[<>:"\/\\|?*\x00-\x1F]+/g, '_') || 'receipt';
  const tableRows = rows.map(row => {
    const label = Array.isArray(row) ? row[0] : row.label;
    const value = Array.isArray(row) ? row[1] : row.value;
    return `<tr><td><strong>${escapeHtmlGs(label)}</strong></td><td>${escapeHtmlGs(value || '-')}</td></tr>`;
  }).join('');
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial, sans-serif; color: #111827; padding: 24px; }
          h1 { margin: 0 0 6px; color: #064b85; }
          h2 { margin: 0 0 14px; }
          table { width: 100%; border-collapse: collapse; }
          td { border: 1px solid #cbd5e1; padding: 8px; vertical-align: top; }
          td:first-child { width: 34%; background: #f8fafc; }
        </style>
      </head>
      <body>
        <h1>Adarsh Cyber Zone</h1>
        <h2>${escapeHtmlGs(title)}</h2>
        <p><strong>Date:</strong> ${Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MM-yyyy HH:mm:ss')}</p>
        <table>${tableRows}</table>
      </body>
    </html>`;
  const pdfBlob = Utilities.newBlob(html, MimeType.HTML, `${safeName}.html`).getAs(MimeType.PDF).setName(`${safeName}.pdf`);
  const file = getUploadFolder().createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return {
    success: true,
    fileName: `${safeName}.pdf`,
    link: file.getUrl(),
    downloadUrl: `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}`
  };
}

function createFilePdf(data) {
  const sourceUrl = String(data.url || data.link || '').trim();
  const id = driveFileIdFromUrl(sourceUrl);
  if (!id) return { success: false, message: 'File ID not found.' };
  const sourceFile = DriveApp.getFileById(id);
  const requestedName = String(data.fileName || sourceFile.getName() || 'document.pdf').replace(/\.pdf$/i, '');
  const safeName = requestedName.replace(/[<>:"\/\\|?*\x00-\x1F]+/g, '_') || 'document';
  const sourceBlob = sourceFile.getBlob();
  const pdfBlob = sourceBlob.getContentType() === MimeType.PDF
    ? sourceBlob.setName(`${safeName}.pdf`)
    : sourceBlob.getAs(MimeType.PDF).setName(`${safeName}.pdf`);
  const file = getUploadFolder().createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfId = file.getId();
  return {
    success: true,
    fileName: `${safeName}.pdf`,
    link: file.getUrl(),
    downloadUrl: `https://drive.google.com/uc?export=download&id=${encodeURIComponent(pdfId)}`
  };
}

function driveFileIdFromUrl(url) {
  const text = String(url || '');
  let match = text.match(/\/d\/([^\/?#]+)/);
  if (match) return match[1];
  match = text.match(/[?&]id=([^&]+)/);
  if (match) return match[1];
  return '';
}

function getUploadFolder() {
  const name = 'Adarsh_Cyber_Zone_Uploads';
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getObjects(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const headers = getHeaders(sheet);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    .map(row => headers.reduce((obj, h, i) => (obj[h] = row[i], obj), {}));
}

function getObjectsWithRow(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const headers = getHeaders(sheet);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
    .map((row, index) => {
      const obj = headers.reduce((out, h, i) => (out[h] = row[i], out), {});
      obj._row = index + 2;
      return obj;
    });
}

function appendObject(sheet, headers, obj) {
  sheet.appendRow(headers.map(h => obj[h] || ''));
}

function rowToObject(sheet, rowNumber) {
  const headers = getHeaders(sheet);
  const row = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.reduce((obj, h, i) => (obj[h] = row[i], obj), {});
}

function updateRow(sheet, rowNumber, updates) {
  const headers = getHeaders(sheet);
  Object.keys(updates).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (col > 0) sheet.getRange(rowNumber, col).setValue(updates[key]);
  });
}

function findRowByValue(sheet, header, value) {
  const headers = getHeaders(sheet);
  const col = headers.indexOf(header) + 1;
  if (col < 1 || sheet.getLastRow() < 2) return { row: -1 };
  const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === String(value).trim().toUpperCase()) return { row: i + 2 };
  }
  return { row: -1 };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function escapeHtmlGs(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
 function setInitialRegNo() {
  PropertiesService.getScriptProperties()
    .setProperty("LAST_REG_NO", "3");
}
