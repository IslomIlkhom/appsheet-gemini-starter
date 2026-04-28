# appsheet-gemini-starter

> Call Gemini AI from Google Apps Script. Built for AppSheet automations.

[![npm: Apps Script](https://img.shields.io/badge/runtime-Apps%20Script%20V8-34A853)](https://script.google.com)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)

📺 [Watch the video walkthrough](https://www.youtube.com/watch?v=eIrrYjY4M1M)

## Features

- **`callGemini(prompt)`** — text prompt, get text back
- **`callGeminiWithFile(prompt, fileId)`** — send a Drive file (PDF, image, Doc) with your prompt
- **`callGeminiJSON(prompt, fileId?)`** — get structured JSON back
- **`classifyDocument(fileId)`** — detect document type, language, summary
- **`extractData(fileId, fields)`** — pull specific fields from any document
- **`summarizeDocument(fileId)`** — title, summary, key points
- **`analyzeText(text)`** — sentiment, category, tags

Handles retries on rate limits, converts Google Docs/Sheets/Slides to PDF automatically, parses JSON even when Gemini wraps it in markdown.

## Install

1. Open your Google Sheet → **Extensions → Apps Script**
2. Paste the contents of [`src/Code.js`](src/Code.js)
3. In **Project Settings → Script Properties**, add:
   - `GEMINI_API_KEY` = your key from [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
4. Run `testConnection` to verify

## Usage

### From Apps Script

```javascript
// Classify a document
var result = classifyDocument('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs');
Logger.log(result);
// {type: "Invoice", language: "EN", summary: "...", confidence: 92}

// Extract specific fields
var data = extractData('1BxiMVs...', ['company_name', 'invoice_number', 'total']);
Logger.log(data.company_name); // "Acme Corp"

// Analyze text
var feedback = analyzeText('The delivery was late and the box was damaged');
Logger.log(feedback.sentiment); // "Negative"

// Raw prompt with file
var answer = callGeminiWithFile('What products are listed in this PDF?', '1BxiMVs...');
Logger.log(answer);
```

### From AppSheet

1. Go to **Automation → Bots**
2. Add a **"Call a script"** task
3. Pick a function (e.g. `classifyDocument`)
4. Pass the file ID from your table

The script writes results to the sheet. AppSheet picks them up automatically.

## Supported files

| Type | Handling |
|------|----------|
| PDF, images | sent directly |
| Google Docs | exported as PDF |
| Google Sheets | exported as PDF |
| Google Slides | exported as PDF |
| Other | sent as blob |

## Configuration

Default model is `gemini-2.0-flash-lite`. Override per call:

```javascript
callGemini('Describe this in detail', {
  model: 'gemini-2.0-flash',
  temperature: 0.8,
  maxTokens: 4096
});
```

| Option | Default | Description |
|--------|---------|-------------|
| `model` | `gemini-2.0-flash-lite` | Gemini model |
| `temperature` | `0.2` | 0 = deterministic, 1 = creative |
| `maxTokens` | `2048` | max response length |

## Background

I extracted this from a production document processing system that classifies thousands of files against a product catalog using Gemini. The full system includes concurrent workers, circuit breakers, and fuzzy matching — this starter is the simplified foundation.

If you're building something similar: [get in touch](https://linkedin.com/in/islomilkhomov).

## License

MIT

---

[YouTube](https://www.youtube.com/@IslomIlkhom) · [LinkedIn](https://linkedin.com/in/islomilkhomov) · [GitHub](https://github.com/IslomIlkhom)
