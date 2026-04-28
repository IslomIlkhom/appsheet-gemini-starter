/**
 * AppSheet + Gemini AI Starter Kit
 * 
 * Connects Google AppSheet to Gemini AI via Apps Script.
 * Classify documents, extract data from images, and analyze text — 
 * all triggered directly from your AppSheet app.
 * 
 * @author Islom Ilkhomov
 * @license MIT
 * @see https://github.com/IslomIlkhom/appsheet-gemini-starter
 */

// ─────────────────────────────────────────────
// CONFIGURATION
// ─────────────────────────────────────────────

/**
 * Get Gemini API key from Script Properties (secure storage).
 * Set via: Extensions → Apps Script → Project Settings → Script Properties
 * Key: GEMINI_API_KEY
 */
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('GEMINI_API_KEY not set in Script Properties');
  return key;
}

/** Gemini API configuration */
const GEMINI_CONFIG = {
  MODEL: 'gemini-2.0-flash',
  API_URL: 'https://generativelanguage.googleapis.com/v1beta/models',
  MAX_RETRIES: 3,
  RETRY_DELAY_MS: 2000
};


// ─────────────────────────────────────────────
// CORE: GEMINI API CLIENT
// ─────────────────────────────────────────────

/**
 * Call Gemini API with text prompt.
 * 
 * @param {string} prompt - Text prompt to send
 * @param {Object} [options] - Optional settings
 * @param {string} [options.model] - Gemini model to use
 * @param {number} [options.temperature] - Response creativity (0-1)
 * @returns {string} Gemini response text
 */
function callGemini(prompt, options = {}) {
  const model = options.model || GEMINI_CONFIG.MODEL;
  const url = `${GEMINI_CONFIG.API_URL}/${model}:generateContent?key=${getApiKey()}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: options.temperature || 0.2,
      maxOutputTokens: options.maxTokens || 2048
    }
  };

  return _sendRequest(url, payload);
}

/**
 * Call Gemini API with text + image (multimodal).
 * Accepts a Google Drive file ID or a Blob.
 * 
 * @param {string} prompt - Text prompt
 * @param {string|Blob} imageSource - Drive file ID or Blob object
 * @param {Object} [options] - Optional settings
 * @returns {string} Gemini response text
 */
function callGeminiWithImage(prompt, imageSource, options = {}) {
  const model = options.model || GEMINI_CONFIG.MODEL;
  const url = `${GEMINI_CONFIG.API_URL}/${model}:generateContent?key=${getApiKey()}`;

  // Get image blob
  let blob;
  if (typeof imageSource === 'string') {
    blob = DriveApp.getFileById(imageSource).getBlob();
  } else {
    blob = imageSource;
  }

  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType: mimeType, data: base64 } }
      ]
    }],
    generationConfig: {
      temperature: options.temperature || 0.2,
      maxOutputTokens: options.maxTokens || 2048
    }
  };

  return _sendRequest(url, payload);
}

/**
 * Call Gemini API and get structured JSON response.
 * 
 * @param {string} prompt - Text prompt (should ask for JSON output)
 * @param {Object} [options] - Optional settings
 * @returns {Object} Parsed JSON response
 */
function callGeminiJSON(prompt, options = {}) {
  const model = options.model || GEMINI_CONFIG.MODEL;
  const url = `${GEMINI_CONFIG.API_URL}/${model}:generateContent?key=${getApiKey()}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: options.temperature || 0.1,
      maxOutputTokens: options.maxTokens || 2048,
      responseMimeType: 'application/json'
    }
  };

  const text = _sendRequest(url, payload);
  try {
    return JSON.parse(text);
  } catch (e) {
    // Try to extract JSON from markdown code blocks
    const match = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (match) return JSON.parse(match[1].trim());
    throw new Error('Failed to parse Gemini JSON response: ' + text.substring(0, 200));
  }
}

/**
 * Internal: Send request to Gemini API with retry logic.
 */
function _sendRequest(url, payload) {
  let lastError;

  for (let attempt = 0; attempt < GEMINI_CONFIG.MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      const body = JSON.parse(response.getContentText());

      if (code === 200) {
        const text = body.candidates?.[0]?.content?.parts?.[0]?.text;
        if (!text) throw new Error('Empty response from Gemini');
        return text.trim();
      }

      // Rate limited — retry with backoff
      if (code === 429) {
        const delay = GEMINI_CONFIG.RETRY_DELAY_MS * Math.pow(2, attempt);
        Logger.log(`Rate limited. Retrying in ${delay}ms (attempt ${attempt + 1})`);
        Utilities.sleep(delay);
        continue;
      }

      throw new Error(`Gemini API error ${code}: ${JSON.stringify(body.error || body)}`);

    } catch (e) {
      lastError = e;
      if (attempt < GEMINI_CONFIG.MAX_RETRIES - 1) {
        Utilities.sleep(GEMINI_CONFIG.RETRY_DELAY_MS * Math.pow(2, attempt));
      }
    }
  }

  throw lastError;
}


// ─────────────────────────────────────────────
// APPSHEET INTEGRATION
// ─────────────────────────────────────────────

/**
 * AppSheet webhook endpoint.
 * 
 * Receives POST requests from AppSheet automations.
 * Routes to the appropriate AI function based on "action" parameter.
 * 
 * Setup: Deploy as Web App → paste URL in AppSheet webhook bot.
 * 
 * @param {Object} e - POST event from AppSheet
 * @returns {Object} JSON response
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;

    switch (action) {
      case 'classify':
        result = classifyDocument(data.fileId);
        break;
      case 'extract':
        result = extractData(data.fileId, data.fields);
        break;
      case 'summarize':
        result = summarizeDocument(data.fileId);
        break;
      case 'analyze':
        result = analyzeText(data.text);
        break;
      default:
        return ContentService.createTextOutput(
          JSON.stringify({ error: 'Unknown action: ' + action })
        ).setMimeType(ContentService.MimeType.JSON);
    }

    // Write result back to sheet if row info provided
    if (data.sheetName && data.rowId) {
      writeResultToSheet(data.sheetName, data.rowId, result);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, result: result })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log('doPost error: ' + error.message);
    return ContentService.createTextOutput(
      JSON.stringify({ error: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}


// ─────────────────────────────────────────────
// AI FUNCTIONS
// ─────────────────────────────────────────────

/**
 * Classify a document by type, language, and topic.
 * Works with PDFs, images, Google Docs, and Office files.
 * 
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} { type, language, summary, confidence }
 */
function classifyDocument(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = _getFileAsPdf(file);

  const prompt = `Analyze this document and return a JSON object with:
- "type": document type (Invoice, Contract, Report, Letter, Specification, Certificate, Marketing, Other)
- "language": 2-letter ISO language code (EN, PL, DE, FR, etc.)
- "summary": one-sentence summary of the document content
- "confidence": your confidence level 0-100

Return ONLY valid JSON, no explanation.`;

  return callGeminiJSON(prompt + '\n\nAnalyze the attached document:', { temperature: 0.1 });
}

/**
 * Extract specific fields from a document.
 * 
 * @param {string} fileId - Google Drive file ID
 * @param {string[]} fields - Fields to extract (e.g., ["company_name", "date", "total_amount"])
 * @returns {Object} Extracted field values
 */
function extractData(fileId, fields) {
  const file = DriveApp.getFileById(fileId);
  const blob = _getFileAsPdf(file);

  const fieldList = (fields || ['company_name', 'date', 'total_amount', 'description']).join(', ');

  const prompt = `Extract the following fields from this document: ${fieldList}

Return a JSON object where each key is the field name and the value is the extracted data.
If a field is not found, set its value to null.
Return ONLY valid JSON, no explanation.`;

  const result = callGeminiJSON(prompt, { temperature: 0.1 });
  
  // Call with image
  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  const url = `${GEMINI_CONFIG.API_URL}/${GEMINI_CONFIG.MODEL}:generateContent?key=${getApiKey()}`;
  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType, data: base64 } }
      ]
    }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  const text = _sendRequest(url, payload);
  try {
    return JSON.parse(text);
  } catch {
    return { raw: text };
  }
}

/**
 * Generate a summary of a document.
 * 
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} { title, summary, key_points, word_count }
 */
function summarizeDocument(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = _getFileAsPdf(file);

  const prompt = `Summarize this document. Return a JSON object with:
- "title": suggested title for this document
- "summary": 2-3 sentence summary
- "key_points": array of 3-5 key points
- "word_count": estimated word count

Return ONLY valid JSON, no explanation.`;

  return callGeminiJSON(prompt);
}

/**
 * Analyze text input (no file needed).
 * Useful for analyzing comments, descriptions, or form inputs from AppSheet.
 * 
 * @param {string} text - Text to analyze
 * @returns {Object} { sentiment, category, tags, summary }
 */
function analyzeText(text) {
  const prompt = `Analyze this text and return a JSON object with:
- "sentiment": Positive, Negative, or Neutral
- "category": best category for this text
- "tags": array of 3-5 relevant tags
- "summary": one-sentence summary

Text to analyze:
"${text}"

Return ONLY valid JSON, no explanation.`;

  return callGeminiJSON(prompt);
}


// ─────────────────────────────────────────────
// SHEET TRIGGER (Alternative to Webhook)
// ─────────────────────────────────────────────

/**
 * Process rows in a Google Sheet that have Status = "New".
 * 
 * Set this on a time-based trigger (every 1-5 minutes).
 * Works without deploying a web app — just uses the Sheet directly.
 * 
 * Expected columns: A=ID, B=File Link, C=Status, D=Type, E=Summary, F=Confidence
 */
function processNewRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Documents');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find column indexes
  const cols = {
    fileLink: headers.indexOf('File Link'),
    status: headers.indexOf('Status'),
    type: headers.indexOf('Type'),
    language: headers.indexOf('Language'),
    summary: headers.indexOf('Summary'),
    confidence: headers.indexOf('Confidence')
  };

  let processed = 0;

  for (let i = 1; i < data.length && processed < 10; i++) {
    const status = data[i][cols.status];
    if (status !== 'New') continue;

    const fileLink = data[i][cols.fileLink];
    if (!fileLink) continue;

    // Extract file ID from Drive URL
    const fileId = _extractFileId(fileLink);
    if (!fileId) {
      sheet.getRange(i + 1, cols.status + 1).setValue('Error: Invalid link');
      continue;
    }

    try {
      // Mark as processing
      sheet.getRange(i + 1, cols.status + 1).setValue('Processing...');
      SpreadsheetApp.flush();

      // Classify with Gemini
      const result = classifyDocument(fileId);

      // Write results
      const row = i + 1;
      if (result.type) sheet.getRange(row, cols.type + 1).setValue(result.type);
      if (result.language) sheet.getRange(row, cols.language + 1).setValue(result.language);
      if (result.summary) sheet.getRange(row, cols.summary + 1).setValue(result.summary);
      if (result.confidence) sheet.getRange(row, cols.confidence + 1).setValue(result.confidence);
      sheet.getRange(row, cols.status + 1).setValue('Done');

      processed++;

    } catch (error) {
      Logger.log('Error processing row ' + (i + 1) + ': ' + error.message);
      sheet.getRange(i + 1, cols.status + 1).setValue('Error: ' + error.message.substring(0, 100));
    }
  }

  if (processed > 0) {
    Logger.log(`Processed ${processed} documents`);
  }
}


// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────

/**
 * Convert any supported file to PDF blob for Gemini.
 */
function _getFileAsPdf(file) {
  const mime = file.getMimeType();

  // Already PDF or image — use directly
  if (mime === 'application/pdf' || mime.startsWith('image/')) {
    return file.getBlob();
  }

  // Google Workspace files — export as PDF
  if (mime.startsWith('application/vnd.google-apps.')) {
    const exportUrl = `https://docs.google.com/document/d/${file.getId()}/export?format=pdf`;
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      return response.getBlob();
    }
  }

  // Fallback: try to get blob directly
  return file.getBlob();
}

/**
 * Extract Google Drive file ID from various URL formats.
 */
function _extractFileId(url) {
  if (!url) return null;
  // Format: /d/FILE_ID/ or id=FILE_ID
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Write AI result back to a sheet row.
 */
function writeResultToSheet(sheetName, rowId, result) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rowId)) {
      const row = i + 1;
      const headers = data[0];

      Object.keys(result).forEach(key => {
        const col = headers.indexOf(key);
        if (col >= 0) {
          const val = Array.isArray(result[key]) ? result[key].join(', ') : result[key];
          sheet.getRange(row, col + 1).setValue(val);
        }
      });
      break;
    }
  }
}


// ─────────────────────────────────────────────
// QUICK TEST
// ─────────────────────────────────────────────

/**
 * Test the Gemini connection. Run this first to verify setup.
 */
function testGeminiConnection() {
  try {
    const result = callGemini('Reply with exactly: "Connection successful"');
    Logger.log('✅ Gemini API connected: ' + result);
    return true;
  } catch (e) {
    Logger.log('❌ Connection failed: ' + e.message);
    return false;
  }
}

/**
 * Test document classification with a specific file.
 * Replace FILE_ID with an actual Google Drive file ID.
 */
function testClassify() {
  const FILE_ID = 'YOUR_FILE_ID_HERE'; // ← Replace this
  const result = classifyDocument(FILE_ID);
  Logger.log(JSON.stringify(result, null, 2));
}
