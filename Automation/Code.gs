const CONFIG = {
  FORM_SHEET_NAME: "Form Responses 1",
  OUTPUT_SHEET_NAME: "Astra Expenses",
  UPLOAD_QUESTION_TITLE: "Upload Invoice",
  DEFAULT_ENTERED_VALUE: "BILL",
  CURRENCY_SYMBOL: "",
  GEMINI_MODEL: "gemini-1.5-flash",
  CSV_EXPORT_FOLDER_ID: ""
};

const OUTPUT_HEADERS = [
  "Date",
  "Entered",
  "Ref. No.",
  "Suppliers",
  "#",
  "Sub-Total",
  "#",
  "GST, if any",
  "Payable"
];

function setupOutputSheet() {
  const sheet = getOrCreateOutputSheet_();
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
  const receiptData = analyzeReceipt_(fileId);
  const rowValues = buildRowValues_(receiptData, submittedAt);

  const outputSheet = getOrCreateOutputSheet_();
  outputSheet.appendRow(rowValues);

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
  const rowValues = buildRowValues_(receiptData, submittedAt);

  const outputSheet = getOrCreateOutputSheet_();
  outputSheet.appendRow(rowValues);

  exportOutputToCsv_(outputSheet);
}

function backfillAllResponses() {
  const formSheet = getFormSheet_();
  const lastRow = formSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No form responses found.");
  }

  for (let row = 2; row <= lastRow; row += 1) {
    const fileId = extractFileIdFromRow_(formSheet, row);
    if (!fileId) {
      continue;
    }

    const submittedAt = extractSubmittedAtFromRow_(formSheet, row);
    const receiptData = analyzeReceipt_(fileId);
    const rowValues = buildRowValues_(receiptData, submittedAt);

    const outputSheet = getOrCreateOutputSheet_();
    outputSheet.appendRow(rowValues);
  }
}

function analyzeReceipt_(fileId) {
  const apiKey = getRequiredScriptProperty_("GEMINI_API_KEY");
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const base64 = Utilities.base64Encode(blob.getBytes());

  const prompt = [
    "You are extracting data from a receipt or tax invoice image.",
    "Return ONLY a JSON object with these keys:",
    "vendor, invoice_number, date, subtotal, tax, total, currency.",
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

function buildRowValues_(receiptData, submittedAt) {
  const normalized = normalizeReceiptData_(receiptData, submittedAt);

  return [
    normalized.date,
    CONFIG.DEFAULT_ENTERED_VALUE,
    normalized.invoiceNumber,
    normalized.vendor,
    CONFIG.CURRENCY_SYMBOL,
    normalized.subtotal,
    CONFIG.CURRENCY_SYMBOL,
    normalized.tax,
    normalized.total
  ];
}

function normalizeReceiptData_(receiptData, submittedAt) {
  const vendor = receiptData.vendor || "";
  const invoiceNumber = receiptData.invoice_number || receiptData.reference_number || "";
  const dateValue = receiptData.date || "";

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
