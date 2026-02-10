# LinkedIn Project Description — Receipt Automation

Use the text below when adding this project to your LinkedIn profile (Projects section) or when posting about it.

---

## Short headline (for Projects / post title)

**Receipt Automation: AI-Powered Expense Extraction with Google Forms & Gemini**

---

## One-paragraph summary (for Projects description or post)

I built an automated receipt processing system that turns receipt images into structured expense data. Users upload receipts via a Google Form; Google Apps Script triggers on each submission, sends the image to the **Gemini 2.5 Flash** API for extraction, and writes normalized rows (date, vendor, ref. no., subtotal, GST, payable) into a Google Sheet with a fixed CSV layout. The flow supports AI-only or AI + manual fallback, optional CSV export to Drive, and configurable formatting (e.g. supplier name casing, ref. number handling). It’s designed for expense tracking with minimal manual data entry.

---

## Bullet points (for Projects or post)

- **End-to-end automation**: Google Form → Apps Script → Gemini API → Google Sheet with a consistent CSV-style layout  
- **AI extraction**: Uses Gemini 2.5 Flash to read vendor, date, amounts, tax, and reference numbers from receipt images (PDF/JPG/PNG)  
- **Two workflows**: AI-only (upload only) or AI + optional manual fields as fallback  
- **Structured output**: Fixed columns (Date, Entered by, Ref. No., Suppliers, Sub-Total, GST, Payable, Total Payable) with auto-calculated grand total  
- **Optional CSV export**: Each submission can trigger a CSV export to a Google Drive folder  
- **Configurable**: Supplier casing, ref. number length, model name, and sheet names via a single config object  

---

## Tech stack (for Projects or post)

Google Forms • Google Sheets • Google Apps Script • Google Drive API • Gemini 2.5 Flash (Google AI) • JavaScript

---

## Suggested media for LinkedIn

- **Screenshot 1**: Google Form with “Upload Invoice” and optional fields  
- **Screenshot 2**: “Astra Expenses” sheet with a few sample rows of extracted data  
- **Screenshot 3**: Apps Script editor showing the trigger or a key function (e.g. `onFormSubmit` or `analyzeReceipt_`)  
- **Optional**: Short screen recording of submitting a form and seeing a new row appear in the sheet  

You can use the images in your `Resources/` folder or capture new screens from your Form/Sheet/Script.

---

## Link to use

If you publish the repo (e.g. on GitHub), paste the repo URL in the “Project link” field on LinkedIn.

---

## Example post (copy-paste and adjust)

**Receipt Automation — From image to expense row in one click**

I built a small automation that uses **Google Forms**, **Apps Script**, and **Gemini 2.5 Flash** to turn receipt photos into structured expense rows in a Google Sheet.

Flow: upload receipt → form submit → script runs → Gemini extracts vendor, date, amounts, tax → row appended to sheet with a fixed CSV layout. Supports AI-only or AI + manual fallback, optional CSV export, and configurable formatting.

Tech: Google Forms, Sheets, Apps Script, Drive API, Gemini 2.5 Flash, JavaScript.

[Add your repo link or “DM for details” if you prefer.]

#Automation #GoogleWorkspace #Gemini #ExpenseTracking #AppsScript #AI
