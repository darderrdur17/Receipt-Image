const CONFIG = {
  FORM_SHEET_NAME: "Form Responses 1",
  OUTPUT_SHEET_NAME: "Astra Expenses",
  UPLOAD_QUESTION_TITLE: "Upload Invoice",
  DEFAULT_ENTERED_VALUE: "BILL",
  SUPPLIER_NAME_CASE: "title",
  REF_NUMBER_MAX_LENGTH: 20,
  TRANS_NUMBER_MAX_LENGTH: 20,
  GEMINI_MODEL: "gemini-2.5-flash",
  CSV_EXPORT_FOLDER_ID: ""
};

const OUTPUT_HEADERS = [
  "Date",
  "Entered by",
  "Ref. No.",
  "Suppliers",
  "Sub-Total",
  "GST, if any",
  "Payable",
  "Total Payable"
];

function setupOutputSheet() {
  const sheet = getOrCreateOutputSheet_();
  trimOutputColumns_(sheet);
  const headerRange = sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length);
  headerRange.setValues([OUTPUT_HEADERS]);
  headerRange.setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function onFormSubmit(e) {
  if (!e) {
    throw new Error("onFormSubmit requires an event object.");
  }

  const fileId = extractFileIdFromEvent_(e);
  if (!fileId) {
    throw new Error("No file ID found in the form submission.");
  }

  const submittedAt = extractSubmittedAt_(e);
  const formData = extractFormData_(e);
  const receiptData = analyzeReceipt_(fileId);

  const outputSheet = getOrCreateOutputSheet_();
  trimOutputColumns_(outputSheet);
  const rowValues = buildRowValues_(receiptData, submittedAt, formData, outputSheet);
  outputSheet.appendRow(rowValues);
  updateAllTotalPayable_(outputSheet);

  exportOutputToCsv_(outputSheet);
}

function processLatestResponse() {
  const formSheet = getFormSheet_();
  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No form responses found.");
  }

  const fileId = extractFileIdFromRow_(formSheet, lastRow);
  if (!fileId) {
    throw new Error("No file ID found in the latest response row.");
  }

  const submittedAt = extractSubmittedAtFromRow_(formSheet, lastRow);
  const receiptData = analyzeReceipt_(fileId);

  const outputSheet = getOrCreateOutputSheet_();
  trimOutputColumns_(outputSheet);
  const rowValues = buildRowValues_(receiptData, submittedAt, {}, outputSheet);
  outputSheet.appendRow(rowValues);
  updateAllTotalPayable_(outputSheet);

  exportOutputToCsv_(outputSheet);
}

function backfillAllResponses() {
  const formSheet = getFormSheet_();
  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No form responses found.");
  }

  const outputSheet = getOrCreateOutputSheet_();
  trimOutputColumns_(outputSheet);

  for (let row = 2; row <= lastRow; row += 1) {
    const fileId = extractFileIdFromRow_(formSheet, row);
    if (!fileId) {
      continue;
    }

    const submittedAt = extractSubmittedAtFromRow_(formSheet, row);
    const receiptData = analyzeReceipt_(fileId);
    const rowValues = buildRowValues_(receiptData, submittedAt, {}, outputSheet);
    outputSheet.appendRow(rowValues);
  }

  updateAllTotalPayable_(outputSheet);
}

function standardizeOutputSheet() {
  const sheet = getOrCreateOutputSheet_();
  standardizeOutputSheet_(sheet);
}

function analyzeReceipt_(fileId) {
  const apiKey = getRequiredScriptProperty_("GEMINI_API_KEY");
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  const prompt = [
    "You are extracting data from a receipt or tax invoice image.",
    "Return ONLY a JSON object with these keys:",
    "vendor, invoice_number, transaction_number, auth_code, date, subtotal, tax, total, currency, payment_method, card_reference, cash_reference.",
    "invoice_number, transaction_number, card_reference, and cash_reference must be strings and must preserve leading zeros as printed.",
    "If you see TRANS. NUMBER / TRANSACTION NO / TRANS NO, put it in transaction_number.",
    "If you see AUTH CODE, put it in auth_code (do NOT use it as invoice_number).",
    "If you see REF. NO (or reference no.) on card receipts, put it in card_reference.",
    "If you see POS/TR/ID line on cash receipts (e.g., POS1 TR#### ID########), put it in cash_reference.",
    "If only one reference exists, also set invoice_number to that value (but never use auth_code).",
    "Format date as DD/MM/YYYY when possible.",
    "Use numbers for subtotal, tax, and total without currency symbols.",
    "If a field is missing, set it to null."
  ].join("\n");

  const requestBody = {
    contents: [
      {
        role: "user",
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: blob.getContentType(),
              data: base64
            }
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.2
    }
  };

  const url = "https://generativelanguage.googleapis.com/v1beta/models/" +
    CONFIG.GEMINI_MODEL + ":generateContent?key=" + apiKey;

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  const responseText = response.getContentText();
  if (response.getResponseCode() >= 300) {
    throw new Error("Gemini request failed: " + responseText);
  }

  const parsed = JSON.parse(responseText);
  const contentParts = (parsed.candidates || [])
    .map((candidate) => candidate.content && candidate.content.parts)
    .filter(Boolean)
    .flat();
  const text = (contentParts || [])
    .map((part) => part.text || "")
    .join("\n")
    .trim();

  return parseReceiptJson_(text);
}

function parseReceiptJson_(text) {
  if (!text) {
    return {};
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    const match = text.match(/\{[\s\S]*\}/);
    if (match) {
      return JSON.parse(match[0]);
    }
    return {};
  }
}

function buildRowValues_(receiptData, submittedAt, formData, sheet) {
  const normalized = normalizeReceiptData_(receiptData, submittedAt, formData);

  return [
    normalized.date,
    CONFIG.DEFAULT_ENTERED_VALUE,
    normalized.invoiceNumber,
    normalized.vendor,
    normalized.subtotal,
    normalized.tax,
    normalized.total,
    ""
  ];
}

function normalizeReceiptData_(receiptData, submittedAt, formData) {
  // Use AI-extracted data as primary, form data as fallback
  const vendor = normalizeSupplierName_(receiptData.vendor || formData.vendorName || "");
  const invoiceNumber = selectReferenceNumber_(receiptData);
  const dateValue = receiptData.date || formData.dateOfPurchase || "";

  const subtotal = parseNumber_(receiptData.subtotal);
  const tax = parseNumber_(receiptData.tax || receiptData.gst);
  const total = parseNumber_(receiptData.total || receiptData.payable);

  let resolvedSubtotal = subtotal;
  let resolvedTax = tax;
  let resolvedTotal = total;

  if (!resolvedTotal && resolvedSubtotal && resolvedTax) {
    resolvedTotal = resolvedSubtotal + resolvedTax;
  }
  if (!resolvedTax && resolvedTotal && resolvedSubtotal) {
    resolvedTax = resolvedTotal - resolvedSubtotal;
  }
  if (!resolvedSubtotal && resolvedTotal && resolvedTax) {
    resolvedSubtotal = resolvedTotal - resolvedTax;
  }

  return {
    vendor: vendor,
    invoiceNumber: invoiceNumber,
    date: resolveDate_(dateValue, submittedAt),
    subtotal: resolvedSubtotal || "",
    tax: resolvedTax || "",
    total: resolvedTotal || ""
  };
}

function resolveDate_(value, submittedAt) {
  if (value) {
    return value;
  }
  if (!submittedAt) {
    return "";
  }
  const timeZone = Session.getScriptTimeZone();
  return Utilities.formatDate(submittedAt, timeZone, "dd/MM/yyyy");
}

function parseNumber_(value) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  if (typeof value === "number") {
    return value;
  }
  const cleaned = String(value)
    .replace(/[^0-9.\-]/g, "")
    .trim();
  if (!cleaned) {
    return "";
  }
  const parsed = Number(cleaned);
  return Number.isNaN(parsed) ? "" : parsed;
}

function selectReferenceNumber_(receiptData) {
  const paymentMethod = getPaymentMethod_(receiptData);
  const cardReference = firstNonEmpty_([
    receiptData.card_reference,
    receiptData.card_ref,
    receiptData.card_ref_no,
    receiptData.ref_no,
    receiptData.reference_number,
    receiptData.reference_no,
    receiptData.ref_number
  ]);
  const cashReference = firstNonEmpty_([
    receiptData.cash_reference,
    receiptData.cash_ref,
    receiptData.pos_reference,
    receiptData.pos_ref,
    receiptData.pos_id,
    receiptData.transaction_reference
  ]);
  const transactionNumber = firstNonEmpty_([
    receiptData.transaction_number,
    receiptData.trans_number,
    receiptData.trans_no,
    receiptData.trans_id
  ]);
  const authCode = firstNonEmpty_([
    receiptData.auth_code,
    receiptData.authorization_code,
    receiptData.auth,
    receiptData.authCode
  ]);
  const invoiceNumber = firstNonEmpty_([
    receiptData.invoice_number,
    receiptData.invoice_no,
    receiptData.invoice
  ]);

  let selected = "";
  let maxLength = CONFIG.REF_NUMBER_MAX_LENGTH;

  if (transactionNumber) {
    selected = transactionNumber;
    maxLength = CONFIG.TRANS_NUMBER_MAX_LENGTH;
  } else if (isCashPayment_(paymentMethod)) {
    selected = cashReference || invoiceNumber || cardReference || "";
  } else if (isCardPayment_(paymentMethod)) {
    selected = cardReference || invoiceNumber || cashReference || "";
  } else {
    selected = invoiceNumber || cardReference || cashReference || "";
  }

  if (authCode && selected === authCode) {
    selected = transactionNumber || cashReference || cardReference || invoiceNumber || "";
    maxLength = selected === transactionNumber ? CONFIG.TRANS_NUMBER_MAX_LENGTH : maxLength;
  }

  return formatReferenceNumber_(selected, maxLength);
}

function firstNonEmpty_(values) {
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value === 0) {
      return "0";
    }
    if (value) {
      const text = String(value).trim();
      if (text) {
        return text;
      }
    }
  }
  return "";
}

function getPaymentMethod_(receiptData) {
  const value = receiptData.payment_method ||
    receiptData.paymentMethod ||
    receiptData.payment_type ||
    receiptData.paymentType ||
    receiptData.tender_type ||
    receiptData.tender ||
    "";
  return String(value).trim().toLowerCase();
}

function isCashPayment_(paymentMethod) {
  return paymentMethod.includes("cash");
}

function isCardPayment_(paymentMethod) {
  const cardHints = [
    "visa",
    "mastercard",
    "master card",
    "amex",
    "american express",
    "card",
    "credit",
    "debit",
    "paywave",
    "contactless",
    "nets",
    "unionpay"
  ];
  return cardHints.some((hint) => paymentMethod.includes(hint));
}

function formatReferenceNumber_(value, maxLength) {
  if (value === null || value === undefined || value === "") {
    return "";
  }
  const raw = String(value).trim();
  if (!raw) {
    return "";
  }

  let normalized = raw;
  const shouldTrim = typeof maxLength === "number" && maxLength > 0;
  if (shouldTrim && /^\d+$/.test(normalized) && normalized.length > maxLength) {
    normalized = normalized.slice(0, maxLength);
  }

  if (normalized.startsWith("0")) {
    return "'" + normalized;
  }
  return normalized;
}

function normalizeSupplierName_(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  const normalized = raw.replace(/\s+/g, " ");
  const mode = String(CONFIG.SUPPLIER_NAME_CASE || "title").toLowerCase();
  if (mode === "upper") {
    return normalized.toUpperCase();
  }
  if (mode === "lower") {
    return normalized.toLowerCase();
  }
  return toTitleCase_(normalized.toLowerCase());
}

function toTitleCase_(value) {
  return value.split(" ").map(titleCaseToken_).join(" ");
}

function titleCaseToken_(token) {
  return token.replace(/[a-zA-Z]+/g, (word) => {
    if (word.length <= 2) {
      return word.toUpperCase();
    }
    return word.charAt(0).toUpperCase() + word.slice(1);
  });
}

function extractFileIdFromEvent_(e) {
  const namedValues = e.namedValues || {};
  const uploadValues = namedValues[CONFIG.UPLOAD_QUESTION_TITLE];
  return extractFirstFileId_(uploadValues);
}

function extractSubmittedAt_(e) {
  const namedValues = e.namedValues || {};
  const timestampValues = namedValues.Timestamp || namedValues["Timestamp"];
  if (timestampValues && timestampValues[0]) {
    const parsed = new Date(timestampValues[0]);
    if (!Number.isNaN(parsed.valueOf())) {
      return parsed;
    }
  }
  return null;
}

function extractFormData_(e) {
  const namedValues = e.namedValues || {};

  return {
    dateOfPurchase: (namedValues["Date of Purchase"] || namedValues["Date of Purchase"] || [""])[0] || "",
    vendorName: (namedValues["Vendor Name"] || namedValues["Vendor Name"] || [""])[0] || "",
    category: (namedValues["Category"] || namedValues["Category"] || [""])[0] || "",
    notes: (namedValues["Notes"] || namedValues["Notes"] || [""])[0] || ""
  };
}

function extractSubmittedAtFromRow_(sheet, row) {
  const headerMap = getHeaderMap_(sheet);
  const timestampColumn = headerMap.Timestamp || headerMap["Timestamp"];
  if (!timestampColumn) {
    return null;
  }
  const value = sheet.getRange(row, timestampColumn).getValue();
  if (value instanceof Date) {
    return value;
  }
  const parsed = new Date(value);
  return Number.isNaN(parsed.valueOf()) ? null : parsed;
}

function extractFileIdFromRow_(sheet, row) {
  const headerMap = getHeaderMap_(sheet);
  const uploadColumn = headerMap[CONFIG.UPLOAD_QUESTION_TITLE];
  if (!uploadColumn) {
    throw new Error("Upload column not found. Check UPLOAD_QUESTION_TITLE.");
  }
  const cellValue = sheet.getRange(row, uploadColumn).getValue();
  return extractFirstFileId_(cellValue);
}

function extractFirstFileId_(value) {
  if (!value) {
    return "";
  }
  const raw = Array.isArray(value) ? value[0] : value;
  if (!raw) {
    return "";
  }
  const stringValue = String(raw);
  const idMatch = stringValue.match(/[-\w]{25,}/);
  return idMatch ? idMatch[0] : "";
}

function getFormSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.FORM_SHEET_NAME);
  if (!sheet) {
    throw new Error("Form sheet not found: " + CONFIG.FORM_SHEET_NAME);
  }
  return sheet;
}

function getOrCreateOutputSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.OUTPUT_SHEET_NAME);
    sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length).setValues([OUTPUT_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function trimOutputColumns_(sheet) {
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns > OUTPUT_HEADERS.length) {
    sheet.deleteColumns(OUTPUT_HEADERS.length + 1, maxColumns - OUTPUT_HEADERS.length);
  }
}

function updateAllTotalPayable_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const payableColumn = 7;
  const totalPayableColumn = 8;

  if (lastRow > 2) {
    sheet.getRange(2, totalPayableColumn, lastRow - 2, 1).clearContent();
  }

  const payableValues = sheet.getRange(2, payableColumn, lastRow - 1, 1).getValues();

  let grandTotalPayable = 0;
  for (let i = 0; i < payableValues.length; i += 1) {
    const payableValue = parseNumber_(payableValues[i][0]);
    if (payableValue !== "") {
      grandTotalPayable += payableValue;
    }
  }

  if (grandTotalPayable > 0) {
    sheet.getRange(lastRow, totalPayableColumn).setValue(grandTotalPayable);
  }
}

function standardizeOutputSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  const range = sheet.getRange(2, 1, lastRow - 1, OUTPUT_HEADERS.length);
  const values = range.getValues();

  let grandTotalPayable = 0;

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    row[2] = formatReferenceNumber_(row[2], CONFIG.REF_NUMBER_MAX_LENGTH);
    row[3] = normalizeSupplierName_(row[3]);

    const payableValue = parseNumber_(row[6]);
    if (payableValue !== "") {
      grandTotalPayable += payableValue;
    }

    row[7] = "";
    values[rowIndex] = row;
  }

  range.setValues(values);

  if (grandTotalPayable > 0) {
    sheet.getRange(lastRow, 8).setValue(grandTotalPayable);
  }
}

function getHeaderMap_(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headerRow.forEach((header, index) => {
    if (header) {
      map[header] = index + 1;
    }
  });
  return map;
}

function exportOutputToCsv_(sheet) {
  if (!CONFIG.CSV_EXPORT_FOLDER_ID) {
    return;
  }
  const folder = DriveApp.getFolderById(CONFIG.CSV_EXPORT_FOLDER_ID);
  const csv = sheetToCsv_(sheet);
  const timeZone = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd-HHmmss");
  const fileName = "astra-expenses-" + timestamp + ".csv";
  folder.createFile(fileName, csv, MimeType.CSV);
}

function sheetToCsv_(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  return data
    .map((row) => row.map(escapeCsvValue_).join(","))
    .join("\n");
}

function escapeCsvValue_(value) {
  const stringValue = String(value === null || value === undefined ? "" : value);
  if (stringValue.includes(",") || stringValue.includes("\"") || stringValue.includes("\n")) {
    return "\"" + stringValue.replace(/"/g, "\"\"") + "\"";
  }
  return stringValue;
}

function getRequiredScriptProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error("Missing script property: " + key);
  }
  return value;
}
