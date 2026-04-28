/**
 * Gemini API client for Google Apps Script.
 *
 * Wraps the Gemini REST API with retry logic, multimodal support,
 * and structured JSON output. Works with Google Drive files natively.
 *
 * Usage:
 *   1. Set GEMINI_API_KEY in Script Properties
 *   2. Call any function below from your script or AppSheet automation
 *
 * @license MIT
 * @see https://github.com/IslomIlkhom/appsheet-gemini-starter
 */

// ─── Config ──────────────────────────────────────────

function getApiKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Set GEMINI_API_KEY in Script Properties');
  return key;
}

var MODEL = 'gemini-2.0-flash-lite';
var API_BASE = 'https://generativelanguage.googleapis.com/v1beta/models';
var MAX_RETRIES = 3;
var RETRY_MS = 2000;


// ─── Core API ────────────────────────────────────────

/**
 * Send a text prompt to Gemini.
 * @param {string} prompt
 * @param {Object=} opts  {model, temperature, maxTokens}
 * @return {string} response text
 */
function callGemini(prompt, opts) {
  opts = opts || {};
  var url = API_BASE + '/' + (opts.model || MODEL) + ':generateContent?key=' + getApiKey_();
  return sendRequest_(url, {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: opts.temperature || 0.2,
      maxOutputTokens: opts.maxTokens || 2048
    }
  });
}

/**
 * Send a text prompt + file to Gemini.
 * @param {string} prompt
 * @param {string|Blob} file  Drive file ID or Blob
 * @param {Object=} opts
 * @return {string} response text
 */
function callGeminiWithFile(prompt, file, opts) {
  opts = opts || {};
  var blob = typeof file === 'string' ? getBlob_(DriveApp.getFileById(file)) : file;
  var url = API_BASE + '/' + (opts.model || MODEL) + ':generateContent?key=' + getApiKey_();
  return sendRequest_(url, {
    contents: [{
      parts: [
        { text: prompt },
        { inlineData: { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } }
      ]
    }],
    generationConfig: {
      temperature: opts.temperature || 0.2,
      maxOutputTokens: opts.maxTokens || 2048
    }
  });
}

/**
 * Send a prompt (optionally with file) and get JSON back.
 * @param {string} prompt
 * @param {string|Blob=} file  Drive file ID, Blob, or omit
 * @param {Object=} opts
 * @return {Object} parsed JSON
 */
function callGeminiJSON(prompt, file, opts) {
  if (file && typeof file === 'object' && !file.getBytes) {
    opts = file; file = null;
  }
  opts = opts || {};
  var parts = [{ text: prompt }];
  if (file) {
    var blob = typeof file === 'string' ? getBlob_(DriveApp.getFileById(file)) : file;
    parts.push({ inlineData: { mimeType: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) } });
  }
  var url = API_BASE + '/' + (opts.model || MODEL) + ':generateContent?key=' + getApiKey_();
  var text = sendRequest_(url, {
    contents: [{ parts: parts }],
    generationConfig: {
      temperature: opts.temperature || 0.1,
      maxOutputTokens: opts.maxTokens || 2048,
      responseMimeType: 'application/json'
    }
  });
  try { return JSON.parse(text); }
  catch (e) {
    var m = text.match(/```(?:json)?\s*([\s\S]*?)```/);
    if (m) return JSON.parse(m[1].trim());
    throw new Error('Bad JSON from Gemini: ' + text.substring(0, 200));
  }
}


// ─── Built-in Actions ────────────────────────────────

/**
 * Classify a document. Returns {type, language, summary, confidence}.
 * @param {string} fileId  Google Drive file ID
 */
function classifyDocument(fileId) {
  return callGeminiJSON(
    'Classify this document. Return JSON: {type, language (2-letter ISO), summary (1 sentence), confidence (0-100)}. ' +
    'Types: Invoice, Contract, Report, Letter, Specification, Certificate, Marketing, Other.',
    getBlob_(DriveApp.getFileById(fileId))
  );
}

/**
 * Extract fields from a document. Returns {field: value, ...}.
 * @param {string} fileId
 * @param {string[]=} fields  e.g. ["company_name", "date", "amount"]
 */
function extractData(fileId, fields) {
  fields = fields || ['company_name', 'date', 'total_amount', 'description'];
  return callGeminiJSON(
    'Extract these fields from the document: ' + fields.join(', ') + '. ' +
    'Return JSON with each field as key. Use null if not found.',
    getBlob_(DriveApp.getFileById(fileId))
  );
}

/**
 * Summarize a document. Returns {title, summary, key_points, word_count}.
 * @param {string} fileId
 */
function summarizeDocument(fileId) {
  return callGeminiJSON(
    'Summarize this document. Return JSON: {title, summary (2-3 sentences), key_points (array of 3-5), word_count (number)}.',
    getBlob_(DriveApp.getFileById(fileId))
  );
}

/**
 * Analyze text sentiment/category. Returns {sentiment, category, tags, summary}.
 * @param {string} text
 */
function analyzeText(text) {
  return callGeminiJSON(
    'Analyze this text. Return JSON: {sentiment (Positive/Negative/Neutral), category, tags (array of 3-5), summary (1 sentence)}.\n\nText: "' + text + '"'
  );
}


// ─── Internals ───────────────────────────────────────

function sendRequest_(url, payload) {
  var lastErr;
  for (var i = 0; i < MAX_RETRIES; i++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: 'POST', contentType: 'application/json',
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      var code = resp.getResponseCode();
      var body = JSON.parse(resp.getContentText());
      if (code === 200) {
        var t = body.candidates && body.candidates[0] && body.candidates[0].content &&
                body.candidates[0].content.parts && body.candidates[0].content.parts[0] &&
                body.candidates[0].content.parts[0].text;
        if (!t) throw new Error('Empty Gemini response');
        return t.trim();
      }
      if (code === 429) {
        Utilities.sleep(RETRY_MS * Math.pow(2, i));
        continue;
      }
      throw new Error('Gemini ' + code + ': ' + JSON.stringify(body.error || body));
    } catch (e) {
      lastErr = e;
      if (i < MAX_RETRIES - 1) Utilities.sleep(RETRY_MS * Math.pow(2, i));
    }
  }
  throw lastErr;
}

function getBlob_(file) {
  var mime = file.getMimeType();
  if (mime === 'application/pdf' || mime.indexOf('image/') === 0) return file.getBlob();
  var exportUrls = {
    'application/vnd.google-apps.document': 'https://docs.google.com/document/d/' + file.getId() + '/export?format=pdf',
    'application/vnd.google-apps.spreadsheet': 'https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?format=pdf',
    'application/vnd.google-apps.presentation': 'https://docs.google.com/presentation/d/' + file.getId() + '/export?format=pdf'
  };
  if (exportUrls[mime]) {
    var r = UrlFetchApp.fetch(exportUrls[mime], {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (r.getResponseCode() === 200) return r.getBlob();
  }
  return file.getBlob();
}

function extractFileId_(url) {
  if (!url) return null;
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}


// ─── Test ────────────────────────────────────────────

function testConnection() {
  try {
    var r = callGemini('Say "ok"');
    Logger.log('Connected: ' + r);
  } catch (e) {
    Logger.log('Failed: ' + e.message);
  }
}

function testClassify() {
  var FILE_ID = 'REPLACE_WITH_YOUR_FILE_ID';
  Logger.log(JSON.stringify(classifyDocument(FILE_ID), null, 2));
}
