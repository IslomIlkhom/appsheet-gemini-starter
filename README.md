# AppSheet + Gemini AI Starter Kit ⚡🤖

> Connect Google AppSheet to Gemini AI via Apps Script. Classify documents, extract data, and analyze text — directly from your AppSheet app.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Apps Script](https://img.shields.io/badge/Google-Apps%20Script-green)](https://script.google.com)
[![Gemini](https://img.shields.io/badge/Google-Gemini%20AI-purple)](https://ai.google.dev)

---

## What This Does

This starter kit gives your **AppSheet** apps the power of **Gemini AI** through **Apps Script** — no external services, no webhooks, no complex setup.

**4 ready-to-use AI functions:**

| Function | Input | Output |
|----------|-------|--------|
| 🗂 `classifyDocument()` | Any file (PDF, image, Doc) | Type, language, summary, confidence |
| 📋 `extractData()` | Any file + field list | Structured data from documents |
| 📝 `summarizeDocument()` | Any file | Title, summary, key points |
| 💬 `analyzeText()` | Text string | Sentiment, category, tags |

---

## Architecture

```
┌─────────────┐         ┌──────────────────┐         ┌─────────────┐
│             │  Call    │                  │  API     │             │
│  AppSheet   │────────▶│  Google Apps      │────────▶│  Gemini AI  │
│  Automation │  Script │  Script          │         │  API        │
│             │◀────────│                  │◀────────│             │
└─────────────┘  Write  └──────────────────┘  JSON   └─────────────┘
                 back           │
                         ┌──────┴──────┐
                         │             │
                         ▼             ▼
                   ┌──────────┐  ┌──────────┐
                   │ Google   │  │ Google   │
                   │ Drive    │  │ Sheets   │
                   │ (files)  │  │ (data)   │
                   └──────────┘  └──────────┘
```

**How it works:**
1. User uploads a file or submits data in AppSheet
2. AppSheet automation triggers **"Call a script"**
3. Apps Script sends the file/text to Gemini AI
4. Results are written back to Google Sheets
5. AppSheet displays the results instantly

---

## Quick Start (10 minutes)

### Step 1: Get a Gemini API Key

1. Go to [Google AI Studio](https://aistudio.google.com/apikey)
2. Click **Create API Key**
3. Copy the key

### Step 2: Set Up Apps Script

1. Open your Google Sheet (the one connected to your AppSheet app)
2. Go to **Extensions → Apps Script**
3. Delete the default `Code.gs` content
4. Copy and paste the contents of [`src/Code.js`](src/Code.js)
5. Go to **Project Settings** (⚙️ icon)
6. Under **Script Properties**, add:
   - Key: `GEMINI_API_KEY`
   - Value: *your-api-key*

### Step 3: Test the Connection

1. In the Apps Script editor, select `testConnection` from the dropdown
2. Click **▶ Run**
3. Check **Execution Log** — you should see: `✅ Connection successful`

### Step 4: Connect AppSheet

1. In AppSheet, go to **Automation → Bots**
2. Create a new Bot with an event (e.g., "When a new row is added")
3. Add a **Task** → choose **"Call a script"**
4. Select your Apps Script project
5. Choose the function (e.g., `classifyDocument`)
6. Pass the file ID from your row

That's it. AppSheet now has AI.

---

## API Reference

### `callGemini(prompt, options?)`
Send a text prompt to Gemini.

```javascript
const answer = callGemini("What is AppSheet?");
Logger.log(answer);
```

### `callGeminiWithFile(prompt, fileSource, options?)`
Send text + file to Gemini. Accepts a Drive file ID or Blob.

```javascript
const result = callGeminiWithFile(
  "What product is shown in this image?",
  "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs"
);
```

### `callGeminiJSON(prompt, fileSource?, options?)`
Get structured JSON response from Gemini. Optionally include a file.

```javascript
// Text only
const data = callGeminiJSON(
  'Extract {name, email} from: "John Smith, john@example.com"'
);
// → { name: "John Smith", email: "john@example.com" }

// With file
const result = callGeminiJSON("Classify this document", "1BxiMVs...");
// → { type: "Invoice", language: "EN", ... }
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
const result = extractData("1BxiMVs...", [
  "company_name", "invoice_number", "total_amount", "date"
]);
// → { company_name: "Acme Corp", invoice_number: "INV-001", ... }
```

### `summarizeDocument(fileId)`
Generate a summary of any document.

```javascript
const result = summarizeDocument("1BxiMVs...");
// → { title: "...", summary: "...", key_points: [...], word_count: 1250 }
```

### `analyzeText(text)`
Analyze text from AppSheet form inputs.

```javascript
const result = analyzeText("The product quality was excellent");
// → { sentiment: "Positive", category: "Product Review", tags: [...] }
```

---

## Options

All `callGemini*` functions accept an options object:

| Option | Default | Description |
|--------|---------|-------------|
| `model` | `gemini-2.0-flash-lite` | Gemini model to use |
| `temperature` | `0.2` | Creativity (0 = deterministic, 1 = creative) |
| `maxTokens` | `2048` | Maximum response length |

```javascript
const result = callGemini("Write a creative product description", {
  model: "gemini-2.0-flash",
  temperature: 0.8,
  maxTokens: 1000
});
```

---

## Use Cases

- **Invoice Processing** — Upload invoices → auto-extract vendor, amount, date
- **Document Management** — Drop files → auto-classify by type and language
- **Quality Inspection** — Take photos in AppSheet → AI analyzes for defects
- **Customer Feedback** — Submit text → auto-categorize sentiment and topics
- **Product Catalog** — Upload spec sheets → match against your catalog

---

## Supported File Types

| Type | How it's processed |
|------|--------------------|
| PDF | Sent directly to Gemini |
| Images (PNG, JPG, etc.) | Sent directly to Gemini |
| Google Docs | Exported as PDF → sent to Gemini |
| Google Sheets | Exported as PDF → sent to Gemini |
| Google Slides | Exported as PDF → sent to Gemini |

---

## Built-In Reliability

- **Retry logic** — auto-retries on rate limits (429) with exponential backoff
- **JSON parsing** — handles markdown-wrapped and raw JSON responses
- **File conversion** — auto-converts Docs, Sheets, Slides to PDF for AI
- **Error handling** — clear error messages for debugging

---

## Project Structure

```
appsheet-gemini-starter/
├── src/
│   ├── Code.js           # Copy this into Apps Script
│   └── appsscript.json   # Apps Script manifest
├── LICENSE
└── README.md
```

---

## FAQ

**Q: Do I need a paid Gemini API plan?**  
No. The free tier gives you 15 requests/minute — enough for most AppSheet apps.

**Q: Can I use this with any AppSheet edition?**  
Yes. "Call a script" works with AppSheet Core and Enterprise.

**Q: Can I use a different Gemini model?**  
Yes. Pass `{ model: "gemini-2.0-flash" }` as options to any function.

---

## Author

**Islom Ilkhomov** — Google Workspace & GCP Automation Expert

Building production automation systems with AppSheet, Apps Script, and Gemini AI.

- [LinkedIn](https://linkedin.com/in/islomilkhomov)
- [GitHub](https://github.com/IslomIlkhom)

---

## License

MIT — use it, modify it, build on it.
