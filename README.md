# Receipt Automation (Google Form → Sheet → CSV)

This project creates a fully automated receipt processing system that extracts pricing details from receipt images using Google Gemini AI and formats them into a structured CSV layout for expense tracking.

## ✨ What You Get

- **Smart Google Form** with file upload and optional manual fields
- **Automated AI Processing** using Gemini 2.5 Flash (free tier available)
- **Structured Output** in "Astra Expenses" sheet with your exact CSV format
- **Two Workflow Options**: AI-only (ideal) or AI + manual fallback
- **Optional CSV Export** to Google Drive after each submission

## 📊 CSV / Sheet Layout

The output uses these exact headers for Astra Expenses:

```
Date, Entered, Ref. No., Suppliers, #, Sub-Total, #, GST, if any, Payable
```

## 🚀 Complete Setup Guide (Step-by-Step)

### Prerequisites
- Google account with access to Google Forms, Sheets, and Drive
- Basic familiarity with Google Workspace

### Step 1: Create the Google Form

1. **Open Google Forms**: Go to [forms.google.com](https://forms.google.com)

2. **Create new form**:
   - Click "+" or select "Blank"
   - Title it "Astra Expense Receipt Submission" (or your preference)

3. **Add required file upload question**:
   - Add question → Change type to "File upload"
   - **Title must be exactly**: `Upload Invoice`
   - Configure: Allow 1 file, max 10MB, accept PDF/JPG/PNG

4. **Add optional manual fields** (for fallback if AI fails):
   - Date of Purchase (Date question)
   - Vendor Name (Short answer)
   - Category (Short answer)
   - Notes (Paragraph)

5. **Form Settings**:
   - Go to Settings gear → Check "Limit to 1 response" if desired
   - Responses → Check "Collect email addresses" if needed

### Step 2: Set Up Response Sheet

1. **Link form to Google Sheet**:
   - In form editor, go to "Responses" tab
   - Click green spreadsheet icon
   - Choose "Create a new spreadsheet"
   - Name it "Astra Expense Responses"
   - Click "Create"

2. **Verify sheet setup**:
   - Sheet should have "Form Responses 1" tab
   - Headers should include: Timestamp, Upload Invoice, Date of Purchase, Vendor Name, Category, Notes

### Step 3: Open Apps Script Editor

1. **In your Google Sheet**, go to menu: `Extensions` → `Apps Script`

2. **Apps Script opens** in new tab with empty `Code.gs` file

3. **Project is automatically linked** to your spreadsheet

### Step 4: Copy the Automation Code

1. **Replace all content** in `Code.gs` with the code from `Automation/Code.gs`

2. **Save the script** (Ctrl+S or click save icon)

3. **Name your project** (optional): Click "Untitled project" → enter "Receipt Automation"

### Step 5: Get Gemini API Key (Free!)

1. **Go to Google AI Studio**: [aistudio.google.com](https://aistudio.google.com)

2. **Sign in** with same Google account

3. **Get API key**:
   - Click profile picture → "API Keys"
   - Click "Create API key"
   - Name it "Receipt Automation"
   - Copy the generated key

### Step 6: Configure API Key in Apps Script

1. **In Apps Script editor**, click gear icon (Project Settings)

2. **Go to "Script Properties" section**

3. **Add property**:
   - Property: `GEMINI_API_KEY`
   - Value: Paste your API key
   - Click "Save script properties"

### Step 7: Run Initial Setup

1. **In Apps Script**, find function dropdown or click in `setupOutputSheet` function

2. **Click "Run"** (▶️ button)

3. **Authorize the script**:
   - Click "Authorize access"
   - Choose your account
   - Click "Allow" for required permissions
   - May need to click "Advanced" → "Go to [Project Name] (unsafe)"

4. **Verify success**: Check your Google Sheet for new "Astra Expenses" tab with headers

### Step 8: Create Automatic Trigger

1. **In Apps Script**, click clock icon (Triggers)

2. **Click "+ Add Trigger"**

3. **Configure trigger**:
   - Function: `onFormSubmit`
   - Deployment: `Head`
   - Source: `From spreadsheet`
   - Type: `On form submit`
   - Settings: `Notify me immediately`

4. **Click "Save"**

## 🎯 How to Use (Two Workflow Options)

### **Option 1: Ideal Workflow (AI-Only)**
1. **Just upload receipt image** (leave other fields empty)
2. **Submit form**
3. **AI automatically extracts**: vendor, date, amounts, tax
4. **Data appears** in "Astra Expenses" sheet

### **Option 2: Fallback Workflow (AI + Manual)**
1. **Upload receipt image**
2. **Fill manual fields** if AI might miss data
3. **Submit form**
4. **AI tries first**, manual fields used as backup
5. **Complete data** in "Astra Expenses" sheet

## 🧪 Test Your System

1. **Submit test form** with a receipt image from `Resources/` folder
2. **Wait 15-60 seconds** for processing
3. **Check "Astra Expenses" sheet** for new row with extracted data
4. **Verify data format** matches your CSV requirements

## ⚙️ Configuration Options

Edit the `CONFIG` object in `Code.gs` to customize:

```javascript
const CONFIG = {
  FORM_SHEET_NAME: "Form Responses 1",        // Your form responses sheet name
  OUTPUT_SHEET_NAME: "Astra Expenses",       // Output sheet name
  UPLOAD_QUESTION_TITLE: "Upload Invoice",   // Exact form question title
  DEFAULT_ENTERED_VALUE: "BILL",             // Value for "Entered" column
  REF_NUMBER_MIN_LENGTH: 20,                 // Pad ref. no. with leading zeros
  TRANS_NUMBER_MIN_LENGTH: 20,               // Short transaction number length
  CURRENCY_SYMBOL: "",                       // Leave empty for your format
  GEMINI_MODEL: "gemini-2.5-flash",          // Current Gemini model
  CSV_EXPORT_FOLDER_ID: ""                   // Optional: Google Drive folder ID
};
```

## 🔧 Troubleshooting

### Common Issues & Solutions

**❌ "Gemini model not found" error**
- ✅ **Fixed**: Code uses `gemini-2.5-flash` (current model)
- Update if new model names are released

**❌ Authorization errors**
- ✅ Run `setupOutputSheet()` first to trigger authorization
- Allow all requested permissions
- May need "Advanced" → "Proceed" for personal scripts

**❌ No data in output sheet**
- ✅ Check trigger is created and active
- Verify API key is properly saved in script properties
- Check Apps Script execution logs for errors

**❌ AI extraction incomplete**
- ✅ Use clear, high-resolution receipt photos
- Manual fields provide fallback data
- Edit rows manually if needed

**❌ Ref. No. loses leading zeros (e.g., 001527 → 1527)**
- ✅ `REF_NUMBER_MIN_LENGTH` pads numeric IDs with leading zeros
- ✅ Ref. No. is forced to text when it starts with 0

**❌ Cash vs card reference number mismatch**
- ✅ Card receipts use `REF. NO` as the reference
- ✅ Cash receipts use the `POS/TR/ID` line as the reference
- ✅ If the wrong ref appears, check that payment method text is visible

**❌ Wrong sheet names**
- ✅ Verify "Form Responses 1" exists
- Update `FORM_SHEET_NAME` in config if different

### Debug Tools
- **Apps Script Logs**: View → Logs (shows processing details)
- **Execution History**: Triggers → Executions (shows success/failures)
- **Manual Functions**:
  - `processLatestResponse()`: Process last form submission
  - `backfillAllResponses()`: Process all existing submissions

## 📋 Additional Features

### CSV Export (Optional)
1. Create Google Drive folder for exports
2. Get folder ID from URL: `https://drive.google.com/drive/folders/[FOLDER_ID]`
3. Add to `CSV_EXPORT_FOLDER_ID` in config
4. CSV saves automatically after each submission

### Manual Processing
- Use `processLatestResponse()` to manually process the last form submission
- Use `backfillAllResponses()` to process all existing form submissions

## 🎉 Success Checklist

- [ ] Google Form created with "Upload Invoice" question
- [ ] Form responses linked to Google Sheet
- [ ] Apps Script code copied and saved
- [ ] Gemini API key configured in script properties
- [ ] `setupOutputSheet()` run successfully
- [ ] Trigger created for automatic processing
- [ ] Test submission processed correctly
- [ ] "Astra Expenses" sheet has properly formatted data

## 💡 Tips for Best Results

- **Image Quality**: Clear, well-lit receipt photos work best
- **PDF Support**: Single-page PDFs work; multi-page may need conversion to JPG
- **Free Tier**: Gemini 2.5 Flash has generous free limits
- **Manual Override**: Always edit rows if AI makes mistakes
- **Backup**: Manual fields ensure data completeness

---

**Your receipt automation system is now ready!** 🤖📄✨
