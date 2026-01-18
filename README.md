# Receipt Automation (Google Form -> Sheet -> CSV)

This project sets up an automated flow that reads receipt images from a Google
Form, extracts pricing details with Gemini, and writes clean rows into a
structured sheet that matches your target CSV layout.

## What You Get
- A Google Form for file uploads (receipts/invoices).
- A response sheet that stores the raw form data.
- An output sheet named `Astra Expenses` with the exact columns you requested.
- Optional CSV exports saved to Drive after each submission.

## CSV / Sheet Layout
The output sheet (and CSV export) uses these headers:

```
Date, Entered, Ref. No., Suppliers, #, Sub-Total, #, GST, if any, Payable
```

## Setup Steps
1. **Create the Form**
   - Add a file upload question titled exactly `Upload Invoice`.
   - Point the form responses to a Google Sheet (default sheet is
     `Form Responses 1`).

2. **Open the Script Editor**
   - In the Google Sheet, go to `Extensions` -> `Apps Script`.
   - Create a new project and paste the contents of
     `Automation/Code.gs`.

3. **Set Your Gemini API Key**
   - In Apps Script, go to `Project Settings` -> `Script Properties`.
   - Add `GEMINI_API_KEY` with your API key value.
   - Get a key from https://aistudio.google.com

4. **Adjust Config (if needed)**
   - In `Code.gs`, update:
     - `UPLOAD_QUESTION_TITLE` if your form uses a different label.
     - `FORM_SHEET_NAME` if your responses sheet has a different name.
     - `OUTPUT_SHEET_NAME` if you want a different output sheet name.
     - `CSV_EXPORT_FOLDER_ID` if you want automatic CSV exports.

5. **Create Output Sheet**
   - Run `setupOutputSheet()` once from Apps Script.
   - This creates the `Astra Expenses` sheet with your headers.

6. **Create Trigger**
   - In Apps Script, go to `Triggers`.
   - Add a trigger for `onFormSubmit`.
   - Event source: `From spreadsheet`
   - Event type: `On form submit`

## Test the Flow
1. Submit your form with a receipt image.
2. Check the `Astra Expenses` sheet for the parsed fields.
3. If `CSV_EXPORT_FOLDER_ID` is set, a CSV is saved to Drive.

## Notes
- Best results come from clear, high-resolution receipt photos.
- PDF uploads work best if they are single-page; otherwise upload a JPG/PNG.
- If Gemini misses fields, you can edit the row manually.

## Troubleshooting
- If no file is found, confirm the file upload question title matches
  `UPLOAD_QUESTION_TITLE`.
- If no output sheet is created, run `setupOutputSheet()` manually.
- If Gemini errors, confirm `GEMINI_API_KEY` is set and valid.
