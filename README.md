# AppSheet + Gemini AI Starter Kit ⚡🤖

> Connect Google AppSheet to Gemini AI in 10 minutes. Classify documents, extract data, and analyze text — all from your AppSheet app.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Apps Script](https://img.shields.io/badge/Google-Apps%20Script-green)](https://script.google.com)
[![Gemini](https://img.shields.io/badge/Google-Gemini%20AI-purple)](https://ai.google.dev)

---

## What This Does

This starter kit connects **AppSheet** to **Gemini AI** through **Apps Script** — giving your no-code apps the power of AI.

**4 ready-to-use AI functions:**

| Function | Input | Output |
|----------|-------|--------|
| 🗂 **Classify** | Any document (PDF, image, Doc) | Type, language, summary, confidence |
| 📋 **Extract** | Any document + field list | Structured data from documents |
| 📝 **Summarize** | Any document | Title, summary, key points |
| 💬 **Analyze** | Text from AppSheet fields | Sentiment, category, tags |

**Two integration methods:**
1. **Webhook** — AppSheet automation calls Apps Script web app
2. **Sheet Trigger** — Apps Script processes new rows automatically

---

## Architecture

```
┌─────────────┐      ┌──────────────────┐      ┌─────────────┐
│             │      │                  │      │             │
│  AppSheet   │─────▶│  Google Apps     │─────▶│  Gemini AI  │
│  App        │      │  Script          │      │  API        │
│             │◀─────│                  │◀─────│             │
└─────────────┘      └──────────────────┘      └─────────────┘
       │                      │
       │              ┌───────┴───────┐
       │              │               │
       ▼              ▼               ▼
┌─────────────┐ ┌──────────┐  ┌──────────┐
│ Google      │ │ Google   │  │ Google   │
│ Drive       │ │ Sheets   │  │ Docs     │
│ (files)     │ │ (data)   │  │ (export) │
└─────────────┘ └──────────┘  └──────────┘
```

---

## Quick Start (10 minutes)

### Step 1: Get a Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Click **Create API Key**
3. Copy the key

### Step 2: Set Up Apps Script

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete the default `Code.gs` content
4. Copy and paste the contents of [`src/Code.js`](src/Code.js)
5. Go to **Project Settings** (⚙️ icon)
6. Under **Script Properties**, add:
   - Key: `GEMINI_API_KEY`
   - Value: `your-api-key-here`

### Step 3: Test the Connection

1. In Apps Script editor, select `testGeminiConnection` from the dropdown
2. Click **▶ Run**
3. Check the **Execution Log** — you should see: `✅ Gemini API connected`

### Step 4: Set Up Your Sheet

Create a sheet named **Documents** with these columns:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| **ID** | **File Link** | **Status** | **Type** | **Language** | **Summary** | **Confidence** |

### Step 5: Connect to AppSheet

**Option A: Sheet Trigger (easiest)**
1. In Apps Script, go to **Triggers** (clock icon)
2. Add trigger:
   - Function: `processNewRows`
   - Event source: Time-driven
   - Interval: Every 1 minute
3. In AppSheet, when user uploads a file → write "New" to Status column
4. Apps Script picks it up automatically and classifies it

**Option B: Webhook**
1. In Apps Script, click **Deploy → New Deployment**
2. Select **Web App** → set access to **Anyone**
3. Copy the web app URL
4. In AppSheet, create a **Bot** with a **Call a webhook** task
5. Set the URL and send JSON:
```json
{
  "action": "classify",
  "fileId": "<<[File ID]>>",
  "sheetName": "Documents",
  "rowId": "<<[ID]>>"
}
```

---

## API Reference

### `callGemini(prompt, options?)`
Send a text prompt to Gemini.

```javascript
const answer = callGemini("What is AppSheet?");
Logger.log(answer);
```

### `callGeminiWithImage(prompt, imageSource, options?)`
Send text + image to Gemini. Accepts a Drive file ID or Blob.

```javascript
const result = callGeminiWithImage(
  "What product is shown in this image?",
  "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs" // Drive file ID
);
```

### `callGeminiJSON(prompt, options?)`
Get structured JSON response from Gemini.

```javascript
const data = callGeminiJSON(
  'Extract {name, email, phone} from this text: "John Smith, john@example.com, 555-0123"'
);
// → { name: "John Smith", email: "john@example.com", phone: "555-0123" }
```

### `classifyDocument(fileId)`
Classify a document by type, language, and content.

```javascript
const result = classifyDocument("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs");
// → { type: "Invoice", language: "EN", summary: "...", confidence: 92 }
```

### `extractData(fileId, fields?)`
Extract specific fields from a document.

```javascript
const result = extractData("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs", [
  "company_name", "invoice_number", "total_amount", "date"
]);
// → { company_name: "Acme Corp", invoice_number: "INV-001", ... }
```

### `summarizeDocument(fileId)`
Generate a summary of any document.

```javascript
const result = summarizeDocument("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs");
// → { title: "...", summary: "...", key_points: [...], word_count: 1250 }
```

### `analyzeText(text)`
Analyze text from AppSheet form inputs.

```javascript
const result = analyzeText("The product quality was excellent and delivery was fast");
// → { sentiment: "Positive", category: "Product Review", tags: [...] }
```

---

## Options

All `callGemini*` functions accept an options object:

| Option | Default | Description |
|--------|---------|-------------|
| `model` | `gemini-2.0-flash` | Gemini model to use |
| `temperature` | `0.2` | Creativity (0 = deterministic, 1 = creative) |
| `maxTokens` | `2048` | Maximum response length |

```javascript
const result = callGemini("Write a creative product description", {
  model: "gemini-2.0-pro",
  temperature: 0.8,
  maxTokens: 1000
});
```

---

## Use Cases

### Invoice Processing
Upload invoices in AppSheet → auto-extract vendor, amount, date, line items.

### Document Management
Drop files in a Drive folder → auto-classify by type, language, and department.

### Quality Inspection
Take photos in AppSheet → Gemini analyzes for defects or compliance.

### Customer Feedback
Submit text feedback in AppSheet → auto-categorize sentiment and topics.

### Product Catalog Matching
Upload product documents → match against your catalog using AI.

---

## Built-In Reliability

- **Retry logic** — auto-retries on rate limits (429) with exponential backoff
- **Error handling** — graceful failures with error messages written to sheet
- **JSON parsing** — handles markdown-wrapped and raw JSON responses
- **File conversion** — auto-converts Google Docs, Sheets, Slides to PDF for AI
- **Batch processing** — processes up to 10 rows per trigger run to stay within limits

---

## Supported File Types

| Type | How it's processed |
|------|--------------------|
| PDF | Sent directly to Gemini |
| Images (PNG, JPG, etc.) | Sent directly to Gemini |
| Google Docs | Exported as PDF → sent to Gemini |
| Google Sheets | Exported as PDF → sent to Gemini |
| Google Slides | Exported as PDF → sent to Gemini |
| Office files (DOCX, XLSX) | Converted via Drive → sent to Gemini |

---

## Project Structure

```
appsheet-gemini-starter/
├── src/
│   ├── Code.js           # Main source — copy this to Apps Script
│   └── appsscript.json   # Apps Script manifest
├── docs/
│   └── SETUP.md          # Detailed setup guide
├── LICENSE
└── README.md
```

---

## FAQ

**Q: Do I need a paid Gemini API plan?**
No. The free tier gives you 15 requests/minute, which is enough for most AppSheet use cases.

**Q: Can I use this with AppSheet Core (free)?**
Yes. The Sheet Trigger method works with any AppSheet edition.

**Q: How much does it cost to run?**
Apps Script is free. Gemini API free tier covers most small-to-medium apps. At scale, Gemini 2.0 Flash costs ~$0.10 per 1M input tokens.

**Q: Can I use a different Gemini model?**
Yes. Pass `{ model: "gemini-2.0-pro" }` as options to any function.

---

## Contributing

Issues and PRs welcome. If you build something cool with this, [open an issue](../../issues) and I'll add it to the use cases.

---

## Author

**Islom Ilkhomov** — Google Workspace & GCP Automation Expert

- [LinkedIn](https://linkedin.com/in/islomilkhomov)
- [YouTube](https://youtube.com/@islomilkhomov)
- [GitHub](https://github.com/IslomIlkhom)

Building production automation systems with AppSheet + Apps Script + Gemini AI.

---

## License

MIT — use it, modify it, build on it. Attribution appreciated.
