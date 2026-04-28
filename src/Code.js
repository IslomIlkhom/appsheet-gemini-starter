/**
 * AppSheet + Gemini AI Starter Kit
 * 
 * Connects Google AppSheet to Gemini AI via Apps Script.
 * Classify documents, extract data from images, and analyze text — 
 * all triggered directly from your AppSheet app.
 * 
 * How it works:
 *   1. In AppSheet, create an automation with "Call a script" action
 *   2. The script calls Gemini AI to process files or text
 *   3. Results are written back to the sheet
 *   4. AppSheet sees the updated data instantly
 * 
 * No webhooks. No triggers. Just Apps Script + Gemini.
 * 
 * @author Islom Ilkhom
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
  MODEL: 'gemini-2.0-flash-lite',
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
 * @param {string} [options.model] - Gemini model override
 * @param {number} [options.temperature] - Creativity (0-1)
 * @returns {string} Gemini response text
 * 
 * @example
 *   const answer = callGemini("What is AppSheet?");
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
 * Call Gemini API with text + file (multimodal).
 * Accepts a Google Drive file ID or a Blob.
 * 
 * @param {string} prompt - Text prompt
 * @param {string|Blob} fileSource - Drive file ID or Blob object
 * @param {Object} [options] - Optional settings
 * @returns {string} Gemini response text
 * 
 * @example
 *   const result = callGeminiWithFile("Describe this image", "1BxiMVs...");
 */
function callGeminiWithFile(prompt, fileSource, options = {}) {
  const model = options.model || GEMINI_CONFIG.MODEL;
  const url = `${GEMINI_CONFIG.API_URL}/${model}:generateContent?key=${getApiKey()}`;

  let blob;
  if (typeof fileSource === 'string') {
    const file = DriveApp.getFileById(fileSource);
    blob = _getFileBlob(file);
  } else {
    blob = fileSource;
  }

  const base64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType, data: base64 } }
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
 * Optionally include a file for document analysis.
 * 
 * @param {string} prompt - Text prompt
 * @param {string|Blob|Object} [fileSourceOrOptions] - Drive file ID, Blob, or options
 * @param {Object} [options] - Optional settings
 * @returns {Object} Parsed JSON response
 * 
 * @example
 *   // Text only
 *   const data = callGeminiJSON('Extract {name, email} from: "John, john@test.com"');
 * 
 *   // With file
 *   const result = callGeminiJSON("Classify this document", "1BxiMVs...");
 */
function callGeminiJSON(prompt, fileSourceOrOptions, options = {}) {
  let fileSource = null;

  // Handle overloaded arguments
  if (fileSourceOrOptions && typeof fileSourceOrOptions === 'object' && 
      !fileSourceOrOptions.getBytes && !Array.isArray(fileSourceOrOptions)) {
    options = fileSourceOrOptions;
  } else {
    fileSource = fileSourceOrOptions;
  }

  const model = options.model || GEMINI_CONFIG.MODEL;
  const url = `${GEMINI_CONFIG.API_URL}/${model}:generateContent?key=${getApiKey()}`;

  const parts = [{ text: prompt }];

  if (fileSource) {
    let blob;
    if (typeof fileSource === 'string') {
      blob = _getFileBlob(DriveApp.getFileById(fileSource));
    } else {
      blob = fileSource;
    }
    parts.push({
      inlineData: {
        mimeType: blob.getContentType(),
        data: Utilities.base64Encode(blob.getBytes())
      }
    });
  }

  const payload = {
    contents: [{ parts }],
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
    const match = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (match) return JSON.parse(match[1].trim());
    throw new Error('Failed to parse Gemini JSON: ' + text.substring(0, 200));
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

      if (code === 429) {
        const delay = GEMINI_CONFIG.RETRY_DELAY_MS * Math.pow(2, attempt);
        Logger.log(`Rate limited. Retry in ${delay}ms (attempt ${attempt + 1})`);
        Utilities.sleep(delay);
        continue;
      }

      throw new Error(`Gemini API ${code}: ${JSON.stringify(body.error || body)}`);

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
// AI FUNCTIONS (Call from AppSheet automations)
// ─────────────────────────────────────────────

/**
 * Classify a document by type, language, and topic.
 * Works with PDFs, images, Google Docs, and Office files.
 * 
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} { type, language, summary, confidence }
 */
function classifyDocument(fileId) {
  const blob = _getFileBlob(DriveApp.getFileById(fileId));

  const prompt = `Analyze this document and return JSON:
{
  "type": "Invoice|Contract|Report|Letter|Specification|Certificate|Marketing|Other",
  "language": "2-letter ISO code",
  "summary": "one sentence summary",
  "confidence": 0-100
}`;

  return callGeminiJSON(prompt, blob);
}

/**
 * Extract specific fields from a document.
 * 
 * @param {string} fileId - Google Drive file ID
 * @param {string[]} [fields] - Fields to extract
 * @returns {Object} Extracted field values
 */
function extractData(fileId, fields) {
  const blob = _getFileBlob(DriveApp.getFileById(fileId));
  const fieldList = (fields || ['company_name', 'date', 'total_amount', 'description']).join(', ');

  const prompt = `Extract these fields from the document: ${fieldList}

Return JSON where each key is the field name and value is the extracted data.
If a field is not found, set value to null.`;

  return callGeminiJSON(prompt, blob);
}

/**
 * Generate a summary of a document.
 * 
 * @param {string} fileId - Google Drive file ID
 * @returns {Object} { title, summary, key_points, word_count }
 */
function summarizeDocument(fileId) {
  const blob = _getFileBlob(DriveApp.getFileById(fileId));

  const prompt = `Summarize this document. Return JSON:
{
  "title": "suggested title",
  "summary": "2-3 sentence summary",
  "key_points": ["point 1", "point 2", "point 3"],
  "word_count": estimated_number
}`;

  return callGeminiJSON(prompt, blob);
}

/**
 * Analyze text — sentiment, category, tags.
 * Great for analyzing AppSheet form inputs, comments, feedback.
 * 
 * @param {string} text - Text to analyze
 * @returns {Object} { sentiment, category, tags, summary }
 */
function analyzeText(text) {
  const prompt = `Analyze this text and return JSON:
{
  "sentiment": "Positive|Negative|Neutral",
  "category": "best category",
  "tags": ["tag1", "tag2", "tag3"],
  "summary": "one sentence summary"
}

Text: "${text}"`;

  return callGeminiJSON(prompt);
}


// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────

/**
 * Convert any supported file to a blob Gemini can read.
 * Handles PDFs, images, Google Docs, Sheets, Slides.
 */
function _getFileBlob(file) {
  const mime = file.getMimeType();

  if (mime === 'application/pdf' || mime.startsWith('image/')) {
    return file.getBlob();
  }

  if (mime === 'application/vnd.google-apps.document') {
    return _exportAsPdf('document', file.getId());
  }
  if (mime === 'application/vnd.google-apps.spreadsheet') {
    return _exportAsPdf('spreadsheets', file.getId());
  }
  if (mime === 'application/vnd.google-apps.presentation') {
    return _exportAsPdf('presentation', file.getId());
  }

  return file.getBlob();
}

/**
 * Export Google Workspace file as PDF.
 */
function _exportAsPdf(type, fileId) {
  const urls = {
    document: `https://docs.google.com/document/d/${fileId}/export?format=pdf`,
    spreadsheets: `https://docs.google.com/spreadsheets/d/${fileId}/export?format=pdf`,
    presentation: `https://docs.google.com/presentation/d/${fileId}/export?format=pdf`
  };

  const response = UrlFetchApp.fetch(urls[type], {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Export ${type} failed (HTTP ${response.getResponseCode()})`);
  }
  return response.getBlob();
}

/**
 * Extract Google Drive file ID from URL.
 */
function _extractFileId(url) {
  if (!url) return null;
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}


// ─────────────────────────────────────────────
// TESTING
// ─────────────────────────────────────────────

/** Test Gemini connection. Run this first. */
function testConnection() {
  try {
    const result = callGemini('Reply with exactly: "Connection successful"');
    Logger.log('✅ ' + result);
  } catch (e) {
    Logger.log('❌ ' + e.message);
    Logger.log('Set GEMINI_API_KEY in Script Properties');
  }
}

/** Test classification. Replace FILE_ID. */
function testClassify() {
  const FILE_ID = 'YOUR_FILE_ID_HERE'; // ← replace
  const result = classifyDocument(FILE_ID);
  Logger.log(JSON.stringify(result, null, 2));
}

/** Test text analysis. */
function testAnalyze() {
  const result = analyzeText('The product quality was excellent and delivery was fast');
  Logger.log(JSON.stringify(result, null, 2));
}
