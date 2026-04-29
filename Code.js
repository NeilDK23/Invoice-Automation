// ============================================================
// LUMEPAY INVOICE AUTOMATION v13
// Captures all invoice types; EML + PDF saved to Drive subfolders
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '152xlE7vODMHWG-UXPcFKa25Spe8Wblz4HDMclY96mgA',
  TARGET_SHEET: 'Payment run - Claude',

  PROCESSED_LABEL_NAME: 'expense-automation-processed',

  DRIVE_ROOT_FOLDER_ID: '1Z-KCkRr1TFUTN8ivFMHHUxwkIlO3tK-l',

  DAYS_TO_SCAN: 1,
  TEST_SCAN_DAYS: 5,
  MAX_RESULTS: 10,

  CONFIDENCE_THRESHOLD: 0.80,
  PAYMENT_DAYS: [2, 4],

  NOTIFY_EMAIL: 'neil@lumepay.com',
  DIGEST_RECIPIENTS: ['neil@lumepay.com', 'chad@lumepay.com'],
// Add 'josh@lumepay.com' once confident in script

  DEBUG_MODE: true,
  APPLY_PROCESSED_LABEL: true,
  SEND_DAILY_DIGEST: false,
  LOOP_SLEEP_MS: 1400,

  ANTHROPIC_MODEL: 'claude-sonnet-4-6',
  ANTHROPIC_FILES_BETA: 'files-api-2025-04-14'
};

const COL = {
  INVOICE_NUMBER: 1,
  CUSTOMER_NAME: 2,
  BILLED_TO: 3,
  CURRENCY: 4,
  AMOUNT: 5,
  INVOICE_DATE: 6,
  INVOICE_DUE_DATE: 7,
  PAYMENT_RUN_DATE: 8,
  STATUS: 9,
  CONFIDENCE: 10,
  SOURCE_EML: 11,
  SOURCE_INVOICE: 12,
  SOURCE_EMAIL: 13,
  DATE_EMAIL_RECEIVED: 14,
  DATE_ADDED: 15,
};

// Classifications that should never be written to the sheet
const SKIP_CLASSIFICATIONS = [
  'proof_of_payment_or_remittance',
  'statement_without_specific_payment_request',
  'order_or_shipping_update'
];

// ============================================================
// ENTRY POINTS
// ============================================================

function runDailyInvoiceScan() {
  const day = new Date().getDay();
  if (day === 0 || day === 6) {
    log_('Weekend — skipping scan.');
    return;
  }
  const daysToScan = day === 1 ? 3 : CONFIG.DAYS_TO_SCAN;
  if (day === 1) log_('Monday — scanning 3 days to cover weekend.');
  runInvoiceScan_(daysToScan);
}

function testRun() {
  runInvoiceScan_(CONFIG.TEST_SCAN_DAYS);
}

function runInvoiceScan_(daysToScan) {
  const startedAt = new Date();
  const results = {
    processed: 0,
    invoices: 0,
    flagged: 0,
    duplicates: 0,
    pdf_saved: 0,
    pdf_uploaded: 0,
    skipped: 0,
    errors: [],
  };

  try {
    log_('=== Invoice scan started: ' + startedAt + ' ===');
    log_('Config: APPLY_PROCESSED_LABEL=' + CONFIG.APPLY_PROCESSED_LABEL + ', DEBUG_MODE=' + CONFIG.DEBUG_MODE);

    validateConfig_();

    let processedLabelId = '';
    if (CONFIG.APPLY_PROCESSED_LABEL) {
      log_('Resolving processed label...');
      processedLabelId = getOrCreateLabelId_(CONFIG.PROCESSED_LABEL_NAME);
      log_('Processed label id: ' + processedLabelId);
    }

    const messageIds = fetchMessageIds_(daysToScan);
    log_('Messages found: ' + messageIds.length);

    for (const msgId of messageIds) {
      let subjectForLog = msgId;

      try {
        log_('---');
        log_('Fetching message ' + msgId);

        const msg = fetchFullMessage_(msgId);
        const fields = extractMessageFields_(msg);

        subjectForLog = fields.subject || msgId;
        log_('Processing subject: "' + subjectForLog + '"');

        const outcome = processMessage_(msg, fields, results);
        results.processed++;

        if (outcome === 'invoice') results.invoices++;
        else if (outcome === 'flagged') results.flagged++;
        else if (outcome === 'duplicate') results.duplicates++;
        else results.skipped++;

        if (CONFIG.APPLY_PROCESSED_LABEL && processedLabelId) {
          try {
            log_('Applying processed label...');
            applyLabel_(msgId, processedLabelId);
            log_('Processed label applied.');
          } catch (e) {
            log_('Label skipped for "' + subjectForLog + '": ' + safeError_(e));
          }
        }

        Utilities.sleep(CONFIG.LOOP_SLEEP_MS);

      } catch (err) {
        const message = safeError_(err);
        log_('Error on "' + subjectForLog + '": ' + message);
        results.errors.push({ subject: subjectForLog, error: message });
      }
    }

    try {
      accumulateDailyStats_(results);
    } catch (e) {
      log_('Stats accumulation failed: ' + safeError_(e));
    }

    log_('=== Scan complete: ' + JSON.stringify(results) + ' ===');

  } catch (fatal) {
    const fatalMessage = safeError_(fatal);
    log_('Fatal: ' + fatalMessage);

    try {
      MailApp.sendEmail(
        CONFIG.NOTIFY_EMAIL,
        '[LUMEPAY] Invoice automation fatal error',
        fatalMessage
      );
    } catch (mailErr) {
      log_('Fatal email send failed: ' + safeError_(mailErr));
    }

    throw fatal;
  }
}

// ============================================================
// CONFIG / HEADERS
// ============================================================

function validateConfig_() {
  if (!CONFIG.DRIVE_ROOT_FOLDER_ID || CONFIG.DRIVE_ROOT_FOLDER_ID === 'PASTE_DRIVE_ROOT_FOLDER_ID_HERE') {
    throw new Error('Set CONFIG.DRIVE_ROOT_FOLDER_ID to your Invoice Automation Drive folder ID.');
  }
}

function gmailHeaders_() {
  return {
    Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
  };
}

function anthropicHeaders_(includeFilesBeta) {
  const headers = {
    'x-api-key': getApiKey_(),
    'anthropic-version': '2023-06-01'
  };
  if (includeFilesBeta) headers['anthropic-beta'] = CONFIG.ANTHROPIC_FILES_BETA;
  return headers;
}

// ============================================================
// GMAIL REST HELPERS
// ============================================================

function getOrCreateLabelId_(labelName) {
  const listResp = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/labels',
    { method: 'get', headers: gmailHeaders_(), muteHttpExceptions: true }
  );
  assertHttpOk_(listResp, 'List Gmail labels failed');

  const labels = JSON.parse(listResp.getContentText()).labels || [];
  const existing = labels.find(function(l) { return l.name === labelName; });
  if (existing) return existing.id;

  const createResp = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/labels',
    {
      method: 'post',
      headers: Object.assign({}, gmailHeaders_(), { 'Content-Type': 'application/json' }),
      payload: JSON.stringify({ name: labelName }),
      muteHttpExceptions: true,
    }
  );
  assertHttpOk_(createResp, 'Create Gmail label failed');

  const created = JSON.parse(createResp.getContentText());
  if (!created.id) throw new Error('Create label response missing id');
  return created.id;
}

function fetchMessageIds_(daysToScan) {
  const since = new Date();
  since.setDate(since.getDate() - Number(daysToScan || CONFIG.DAYS_TO_SCAN));
  const afterDate = Utilities.formatDate(since, 'UTC', 'yyyy/MM/dd');

  const rawQuery =
    'in:anywhere after:' + afterDate +
    ' -label:' + CONFIG.PROCESSED_LABEL_NAME +
    ' -from:neil@lumepay.com' +
    ' -(subject:"[Lumepay] Invoice scan")' +
    ' -(subject:"[LUMEPAY] Invoice automation fatal error")' +
    ' (to:expense@lumepay.com OR from:expense@lumepay.com OR cc:expense@lumepay.com)';

  log_('Gmail search query: ' + rawQuery);

  const url =
    'https://gmail.googleapis.com/gmail/v1/users/me/messages?q=' +
    encodeURIComponent(rawQuery) +
    '&maxResults=' + encodeURIComponent(String(CONFIG.MAX_RESULTS));

  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: gmailHeaders_(),
    muteHttpExceptions: true,
  });

  assertHttpOk_(resp, 'Search Gmail messages failed');

  const text = resp.getContentText();
  if (CONFIG.DEBUG_MODE) log_('Search response: ' + truncate_(text, 500));

  const data = JSON.parse(text);
  return (data.messages || []).map(function(m) { return m.id; });
}

function fetchFullMessage_(msgId) {
  const base = 'https://gmail.googleapis.com/gmail/v1/users/me/messages/' + encodeURIComponent(msgId);

  const fullResp = UrlFetchApp.fetch(
    base + '?format=full',
    { method: 'get', headers: gmailHeaders_(), muteHttpExceptions: true }
  );
  assertHttpOk_(fullResp, 'Fetch Gmail full message failed for ' + msgId);

  const rawResp = UrlFetchApp.fetch(
    base + '?format=raw',
    { method: 'get', headers: gmailHeaders_(), muteHttpExceptions: true }
  );
  assertHttpOk_(rawResp, 'Fetch Gmail raw failed for ' + msgId);

  const full = JSON.parse(fullResp.getContentText());
  const rawData = JSON.parse(rawResp.getContentText());

  full._snippet = rawData.snippet || full.snippet || '';
  full._raw = rawData.raw || '';
  full._messageId = msgId;
  full._internalDate = full.internalDate || rawData.internalDate || '';

  return full;
}

function applyLabel_(msgId, labelId) {
  const resp = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/messages/' + encodeURIComponent(msgId) + '/modify',
    {
      method: 'post',
      headers: Object.assign({}, gmailHeaders_(), { 'Content-Type': 'application/json' }),
      payload: JSON.stringify({ addLabelIds: [labelId] }),
      muteHttpExceptions: true,
    }
  );
  assertHttpOk_(resp, 'Apply label failed for message ' + msgId);
}

// ============================================================
// MESSAGE / ATTACHMENTS
// ============================================================

function extractMessageFields_(msg) {
  const headers = (msg.payload && msg.payload.headers) || [];

  function getHeader(name) {
    const lowered = String(name || '').toLowerCase();
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      if ((h.name || '').toLowerCase() === lowered) return h.value || '';
    }
    return '';
  }

  const subject = getHeader('Subject');
  const from = getHeader('From');
  const dateHeader = getHeader('Date');

  let body = '';
  if (msg._raw) {
    try {
      const normalized = String(msg._raw).replace(/-/g, '+').replace(/_/g, '/');
      const decoded = Utilities.newBlob(Utilities.base64Decode(normalized)).getDataAsString();

      const sep1 = decoded.indexOf('\r\n\r\n');
      const sep2 = decoded.indexOf('\n\n');
      const sep = sep1 > -1 ? sep1 : sep2;

      body = sep > -1 ? decoded.substring(sep + (sep1 > -1 ? 4 : 2)) : decoded;
    } catch (e) {
      body = msg._snippet || '';
    }
  } else {
    body = msg._snippet || '';
  }

  const attachmentNames = collectAttachmentNames_(msg.payload).join(', ');
  const pdfAttachments = collectPdfAttachments_(msg.payload);

  return {
    subject: subject || '',
    from: from || '',
    dateHeader: dateHeader || '',
    internalDate: msg._internalDate || '',
    body: String(body || '').substring(0, 5000),
    attachmentNames: attachmentNames || '',
    pdfAttachments: pdfAttachments
  };
}

function collectAttachmentNames_(payload) {
  const names = [];
  if (!payload) return names;
  if (payload.filename) names.push(payload.filename);

  const parts = payload.parts || [];
  for (let i = 0; i < parts.length; i++) {
    const child = collectAttachmentNames_(parts[i]);
    for (let j = 0; j < child.length; j++) names.push(child[j]);
  }
  return names.filter(function(n) { return !!n; });
}

function collectPdfAttachments_(payload) {
  const out = [];
  collectPdfAttachmentsRecursive_(payload, out);
  return out;
}

function collectPdfAttachmentsRecursive_(part, out) {
  if (!part) return;

  const filename = String(part.filename || '');
  const mimeType = String(part.mimeType || '');
  const isPdf = /\.pdf$/i.test(filename) || mimeType === 'application/pdf';
  const attachmentId = part.body && part.body.attachmentId ? part.body.attachmentId : '';

  if (isPdf && attachmentId) {
    out.push({
      filename: filename || 'document.pdf',
      mimeType: mimeType || 'application/pdf',
      attachmentId: attachmentId,
      size: part.body && part.body.size ? Number(part.body.size) : 0
    });
  }

  const parts = part.parts || [];
  for (let i = 0; i < parts.length; i++) {
    collectPdfAttachmentsRecursive_(parts[i], out);
  }
}

function fetchGmailAttachmentBlob_(msgId, attachment) {
  const url =
    'https://gmail.googleapis.com/gmail/v1/users/me/messages/' +
    encodeURIComponent(msgId) +
    '/attachments/' +
    encodeURIComponent(attachment.attachmentId);

  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: gmailHeaders_(),
    muteHttpExceptions: true
  });

  assertHttpOk_(resp, 'Fetch Gmail attachment failed for ' + msgId + ' / ' + attachment.filename);

  const json = JSON.parse(resp.getContentText());
  const data = String(json.data || '').replace(/-/g, '+').replace(/_/g, '/');
  const bytes = Utilities.base64Decode(data);

  return Utilities.newBlob(bytes, 'application/pdf', sanitizeFilename_(attachment.filename || 'document.pdf'));
}

// ============================================================
// DRIVE
// ============================================================

function getDriveMonthFolder_(dateObj) {
  const root = DriveApp.getFolderById(String(CONFIG.DRIVE_ROOT_FOLDER_ID).trim());
  const year = Utilities.formatDate(dateObj, 'Africa/Johannesburg', 'yyyy');
  const monthNum = Utilities.formatDate(dateObj, 'Africa/Johannesburg', 'M');
  const monthName = Utilities.formatDate(dateObj, 'Africa/Johannesburg', 'MMMM');
  const yearFolder = getOrCreateSubfolder_(root, '1. ' + year);
  return getOrCreateSubfolder_(yearFolder, monthNum + '. ' + monthName);
}

function getDrivePdfFolder_(dateObj) {
  return getOrCreateSubfolder_(getDriveMonthFolder_(dateObj), 'PDF Invoices');
}

function getDriveEmlFolder_(dateObj) {
  return getOrCreateSubfolder_(getDriveMonthFolder_(dateObj), 'EML Containing Invoices');
}

function getOrCreateSubfolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

function buildFinalNameBase_(fields, extracted) {
  const customer = sanitizeFilenameComponent_(extracted.customer_name || deriveCustomerFromEmail_(fields) || 'Unknown Customer');
  const dateReceived = formatCompactDate_(parseEmailReceivedDate_(fields));
  const invoiceNo = sanitizeFilenameComponent_(extracted.invoice_number || 'NoInvoiceNo');
  return customer + ' - ' + dateReceived + ' - ' + invoiceNo;
}

function savePdfToDrive_(blob, fields, nameBase) {
  const folder = getDrivePdfFolder_(parseEmailReceivedDate_(fields));
  return folder.createFile(blob.copyBlob().setName(nameBase + '.pdf'));
}

function saveEmlToDrive_(msg, fields, nameBase) {
  const folder = getDriveEmlFolder_(parseEmailReceivedDate_(fields));
  const raw = String(msg._raw || '').replace(/-/g, '+').replace(/_/g, '/');
  const bytes = Utilities.base64Decode(raw);
  const blob = Utilities.newBlob(bytes, 'message/rfc822', nameBase + '.eml');
  return folder.createFile(blob);
}

// ============================================================
// PROCESSING
// ============================================================

function processMessage_(msg, fields, results) {
  const subject = fields.subject;
  const from = fields.from;
  const body = fields.body;
  const attachmentNames = fields.attachmentNames;
  const hasPdf = fields.pdfAttachments && fields.pdfAttachments.length > 0;

  log_('Step: basic validation');
  if (!subject && !from && !body) return 'skipped';

  // Pre-screen only applies when there is no PDF attachment
  log_('Step: pre-screen');
  if (!hasPdf && isDefinitelyNotInvoice_(subject, from, body)) {
    log_('PRE-SCREEN SKIP: ' + subject);
    return 'skipped';
  }

  log_('Step: API key lookup');
  const apiKey = getApiKey_();
  log_('API key present: ' + (!!apiKey));

  let pdfBlob = null;
  let anthropicFileId = '';

  if (hasPdf) {
    const pdf = chooseBestPdfAttachment_(fields.pdfAttachments);
    log_('Using PDF attachment: ' + pdf.filename);
    pdfBlob = fetchGmailAttachmentBlob_(msg.id || msg._messageId || '', pdf);
    anthropicFileId = uploadPdfToAnthropicFiles_(pdfBlob);
    results.pdf_uploaded++;
    log_('Uploaded PDF to Anthropic Files API: ' + anthropicFileId);
  }

  const aiResult = hasPdf
    ? callClaudeForInvoiceFromPdf_(subject, from, body, attachmentNames, anthropicFileId)
    : callClaudeForInvoiceFromEmailOnly_(subject, from, body, attachmentNames, apiKey);

  if (!aiResult || typeof aiResult !== 'object') throw new Error('Claude result was empty or invalid');

  if (SKIP_CLASSIFICATIONS.indexOf(aiResult.classification) !== -1) {
    log_('AI SKIP (' + (aiResult.classification || 'unknown') + ', confidence=' + aiResult.confidence + '): ' + subject);
    return 'skipped';
  }

  const isLowConfidence = Number(aiResult.confidence || 0) < CONFIG.CONFIDENCE_THRESHOLD;
  let status;
  if (isLowConfidence) {
    status = '⚠️ Flagged for review';
  } else if (aiResult.classification === 'invoice_due' || aiResult.classification === 'payment_request') {
    status = 'Payable';
  } else if (aiResult.classification === 'receipt_or_invoice_already_paid') {
    status = 'Paid';
  } else {
    status = 'Other';
  }

  const nameBase = buildFinalNameBase_(fields, aiResult);

  let driveFile = null;
  if (pdfBlob) {
    try {
      driveFile = savePdfToDrive_(pdfBlob, fields, nameBase);
      results.pdf_saved++;
      log_('Saved PDF to Drive: ' + driveFile.getName());
    } catch (e) {
      log_('PDF Drive save failed: ' + safeError_(e));
    }
  }

  let emlFile = null;
  try {
    emlFile = saveEmlToDrive_(msg, fields, nameBase);
    log_('Saved EML to Drive: ' + emlFile.getName());
  } catch (e) {
    log_('EML Drive save failed: ' + safeError_(e));
  }

  const writeResult = writeInvoiceToSheet_(aiResult, msg.id || msg._messageId || '', fields, status, driveFile, emlFile);
  if (writeResult.duplicate) {
    log_('DUPLICATE SKIP: row ' + writeResult.row + ' | ' + subject);
    return 'duplicate';
  }

  log_((isLowConfidence ? 'FLAGGED' : 'INVOICE') + ': ' + subject);
  return isLowConfidence ? 'flagged' : 'invoice';
}

function chooseBestPdfAttachment_(attachments) {
  const scored = attachments.map(function(a) {
    const name = String(a.filename || '').toLowerCase();
    let score = 0;
    if (/invoice|tax invoice|bill/.test(name)) score += 30;
    if (/statement/.test(name)) score += 10;
    if (/proof of payment|remittance|receipt/.test(name)) score -= 20;
    score += Math.min(10, Math.floor((a.size || 0) / 50000));
    return { score: score, item: a };
  });

  scored.sort(function(a, b) { return b.score - a.score; });
  return scored[0].item;
}

function isDefinitelyNotInvoice_(subject, from, body) {
  // Note: this function is only called when the email has no PDF attachment
  const hardNoSubjectPatterns = [
    /newsletter/i, /unsubscribe/i, /\bpromo\b/i, /\bdiscount\b/i, /\bspecial offer\b/i,
    /\bout of office\b/i, /\bverification\b/i, /\bverify your email\b/i, /\bwelcome to\b/i,
    /\bautomatic reply\b/i, /\border placed\b/i, /\bstatus is now\b/i, /\bshipped\b/i,
    /\bpacking\b/i, /\bremittance advice\b/i, /\bdelivery\b/i, /\btracking\b/i
  ];

  const hardNoFromPatterns = [/no.?reply@/i, /donotreply@/i, /noreply@/i];

  for (let i = 0; i < hardNoSubjectPatterns.length; i++) if (hardNoSubjectPatterns[i].test(subject)) return true;
  for (let j = 0; j < hardNoFromPatterns.length; j++) if (hardNoFromPatterns[j].test(from)) return true;

  const combined = String(body || '') + ' ' + String(subject || '');
  const hasInvoiceSignal = /invoice|tax invoice|bill|statement|amount due|please pay|payment due|total due|balance due/i.test(combined);
  const hasNonInvoiceSignal = /proof of payment|order placed|shipped|packing|delivered|tracking/i.test(combined);

  if (hasNonInvoiceSignal && !hasInvoiceSignal) return true;
  return false;
}

// ============================================================
// ANTHROPIC
// ============================================================

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  return key ? String(key).trim() : '';
}

function uploadPdfToAnthropicFiles_(blob) {
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/files', {
    method: 'post',
    headers: anthropicHeaders_(true),
    payload: { file: blob },
    muteHttpExceptions: true
  });

  const status = resp.getResponseCode();
  const text = resp.getContentText();

  log_('Anthropic file upload status: ' + status);
  if (CONFIG.DEBUG_MODE) log_('Anthropic file upload response: ' + truncate_(text, 500));

  if (status < 200 || status >= 300) {
    throw new Error('Anthropic Files upload HTTP ' + status + ': ' + truncate_(text, 1000));
  }

  const json = JSON.parse(text);
  if (!json.id) throw new Error('Anthropic file upload missing file id: ' + truncate_(text, 1000));
  return json.id;
}

function callClaudeForInvoiceFromPdf_(subject, from, body, attachmentNames, anthropicFileId) {
  const systemPrompt =
    'You are an accounts payable extraction engine. ' +
    'Your primary source of truth is the PDF invoice document. ' +
    'Use the email body only as secondary context. ' +
    'Be conservative and do not invent values. ' +
    'Return JSON only.';

  const userText =
    'Analyze this invoice-related email and the attached PDF document for accounts payable processing.\n\n' +
    'EMAIL CONTEXT\n' +
    'From: ' + from + '\n' +
    'Subject: ' + subject + '\n' +
    'Attachment names: ' + (attachmentNames || 'none') + '\n' +
    'Email body:\n' + body + '\n\n' +
    'EXTRACTION RULES\n' +
    '- The PDF is the primary source of truth.\n' +
    '- Classify into one of: invoice_due, payment_request, statement_without_specific_payment_request, receipt_or_invoice_already_paid, proof_of_payment_or_remittance, order_or_shipping_update, other.\n' +
    '- invoice_number: only if explicitly an invoice or tax invoice number.\n' +
    '- customer_name: supplier / vendor issuing the invoice.\n' +
    '- billed_to: the customer/entity being invoiced.\n' +
    '- amount: total amount payable.\n' +
    '- currency: ZAR, USD, EUR, GBP, CAD or explicit code/symbol-derived currency.\n' +
    '- invoice_date: invoice issue date.\n' +
    '- due_date: due date if explicit.\n' +
    '- If multiple totals appear, choose the payable grand total.\n' +
    '- If not explicit, return null.\n\n' +
    'RESPOND ONLY WITH VALID JSON:\n' +
    '{\n' +
    '  "classification": "invoice_due|payment_request|statement_without_specific_payment_request|receipt_or_invoice_already_paid|proof_of_payment_or_remittance|order_or_shipping_update|other",\n' +
    '  "is_invoice": true,\n' +
    '  "confidence": 0.0,\n' +
    '  "reason": "string",\n' +
    '  "invoice_number": "string or null",\n' +
    '  "invoice_date": "YYYY-MM-DD or null",\n' +
    '  "due_date": "YYYY-MM-DD or null",\n' +
    '  "customer_name": "string or null",\n' +
    '  "billed_to": "string or null",\n' +
    '  "currency": "string or null",\n' +
    '  "amount": number or null\n' +
    '}';

  const payload = {
    model: CONFIG.ANTHROPIC_MODEL,
    max_tokens: 900,
    temperature: 0,
    system: systemPrompt,
    messages: [
      {
        role: 'user',
        content: [
          { type: 'text', text: userText },
          {
            type: 'document',
            source: { type: 'file', file_id: anthropicFileId },
            title: 'Invoice PDF',
            context: 'Accounts payable invoice document',
            citations: { enabled: false }
          }
        ]
      }
    ]
  };

  return callAnthropicMessages_(payload, true);
}

function callClaudeForInvoiceFromEmailOnly_(subject, from, body, attachmentNames, apiKey) {
  apiKey = String(apiKey || '').trim();
  if (!apiKey || apiKey === 'YOUR_KEY_HERE' || apiKey === 'PASTE_NEW_ANTHROPIC_KEY_HERE') {
    throw new Error('Missing valid ANTHROPIC_API_KEY in Script Properties. Run storeApiKey() with your real key.');
  }

  const systemPrompt =
    'You are an accounts payable extraction engine. ' +
    'Your job is to identify true invoices or payment requests and extract only reliable fields. ' +
    'Be conservative. If a field is not explicit, return null. ' +
    'Do not invent values. Return JSON only.';

  const userPrompt =
    'Analyze this email for accounts payable processing.\n\n' +
    'EMAIL DETAILS\n' +
    'From: ' + from + '\n' +
    'Subject: ' + subject + '\n' +
    'Attachments: ' + (attachmentNames || 'none') + '\n' +
    'Body:\n' + body + '\n\n' +
    'CLASSIFICATION RULES\n' +
    'Classify the email into one of these values:\n' +
    '- invoice_due\n' +
    '- payment_request\n' +
    '- statement_without_specific_payment_request\n' +
    '- receipt_or_invoice_already_paid\n' +
    '- proof_of_payment_or_remittance\n' +
    '- order_or_shipping_update\n' +
    '- other\n\n' +
    'RESPOND ONLY WITH VALID JSON:\n' +
    '{\n' +
    '  "classification": "invoice_due|payment_request|statement_without_specific_payment_request|receipt_or_invoice_already_paid|proof_of_payment_or_remittance|order_or_shipping_update|other",\n' +
    '  "is_invoice": true,\n' +
    '  "confidence": 0.0,\n' +
    '  "reason": "string",\n' +
    '  "invoice_number": "string or null",\n' +
    '  "invoice_date": "YYYY-MM-DD or null",\n' +
    '  "due_date": "YYYY-MM-DD or null",\n' +
    '  "customer_name": "string or null",\n' +
    '  "billed_to": "string or null",\n' +
    '  "currency": "string or null",\n' +
    '  "amount": number or null\n' +
    '}';

  const payload = {
    model: CONFIG.ANTHROPIC_MODEL,
    max_tokens: 700,
    temperature: 0,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }]
  };

  return callAnthropicMessages_(payload, false);
}

function callAnthropicMessages_(payload, useFilesBeta) {
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: anthropicHeaders_(useFilesBeta),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText();

  log_('Anthropic status (' + CONFIG.ANTHROPIC_MODEL + '): ' + status);
  if (CONFIG.DEBUG_MODE) log_('Anthropic response (' + CONFIG.ANTHROPIC_MODEL + '): ' + truncate_(text, 700));

  if (status < 200 || status >= 300) throw new Error('Anthropic HTTP ' + status + ': ' + truncate_(text, 1000));

  const json = JSON.parse(text);
  if (json.error) throw new Error('Anthropic API error: ' + JSON.stringify(json.error));

  const content = json.content || [];
  const first = content.length ? content[0] : null;
  const rawText = first && first.text ? String(first.text).trim() : '';

  if (!rawText) throw new Error('Anthropic returned no text content: ' + truncate_(text, 1000));

  const clean = rawText.replace(/```json/gi, '').replace(/```/g, '').trim();

  try {
    const parsed = JSON.parse(clean);
    parsed.confidence = Number(parsed.confidence || 0);
    parsed.amount = parsed.amount === null || parsed.amount === undefined || parsed.amount === '' ? null : Number(parsed.amount);
    return parsed;
  } catch (e) {
    log_('JSON parse failed. Raw model text: ' + truncate_(rawText, 1000));
    return {
      classification: 'other',
      is_invoice: false,
      confidence: 0.2,
      reason: 'Parse error - manual review needed',
      invoice_number: null,
      invoice_date: null,
      due_date: null,
      customer_name: null,
      billed_to: null,
      currency: null,
      amount: null
    };
  }
}

// ============================================================
// SHEETS + DUPLICATES
// ============================================================

function writeInvoiceToSheet_(data, msgId, fields, status, driveFile, emlFile) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!sheet) throw new Error('Sheet "' + CONFIG.TARGET_SHEET + '" not found');

  ensureHeaderRow_(sheet);

  const paymentRunDate = data.due_date
    ? getPaymentRunBeforeDueDate_(data.due_date)
    : getNextPaymentRunDate_(data.invoice_date || formatEmailReceived_(fields));
  const emailLink = msgId ? ('https://mail.google.com/mail/u/0/#inbox/' + msgId) : '';
  const dateEmailReceived = formatEmailReceived_(fields);

  const duplicateRow = findDuplicateRow_(sheet, data, emailLink);
  if (duplicateRow > 0) {
    return { duplicate: true, row: duplicateRow };
  }

  const driveFileUrl = driveFile ? driveFile.getUrl() : '';
  const emlFileUrl = emlFile ? emlFile.getUrl() : '';

  const row = new Array(15).fill('');
  row[COL.INVOICE_NUMBER - 1] = nullToEmpty_(data.invoice_number);
  row[COL.CUSTOMER_NAME - 1] = nullToEmpty_(data.customer_name);
  row[COL.BILLED_TO - 1] = nullToEmpty_(data.billed_to);
  row[COL.CURRENCY - 1] = nullToEmpty_(data.currency || '');
  row[COL.AMOUNT - 1] = data.amount === null || data.amount === undefined ? '' : data.amount;
  row[COL.INVOICE_DATE - 1] = nullToEmpty_(data.invoice_date);
  row[COL.INVOICE_DUE_DATE - 1] = nullToEmpty_(data.due_date);
  row[COL.PAYMENT_RUN_DATE - 1] = paymentRunDate;
  row[COL.STATUS - 1] = status;
  row[COL.CONFIDENCE - 1] = Math.round(Number(data.confidence || 0) * 100) + '%';
  row[COL.SOURCE_EMAIL - 1] = emailLink ? 'Email' : '';
  row[COL.SOURCE_INVOICE - 1] = driveFileUrl ? 'Invoice' : '';
  row[COL.DATE_EMAIL_RECEIVED - 1] = dateEmailReceived;
  row[COL.DATE_ADDED - 1] = Utilities.formatDate(new Date(), 'Africa/Johannesburg', 'yyyy-MM-dd HH:mm');
  row[COL.SOURCE_EML - 1] = emlFileUrl ? 'EML' : '';

  sheet.appendRow(row);
  const lastRow = sheet.getLastRow();

  // Columns K, L, M (SOURCE_EML=11, SOURCE_INVOICE=12, SOURCE_EMAIL=13) — contiguous, set in one call
  const richLinks = [[
    emlFileUrl
      ? SpreadsheetApp.newRichTextValue().setText('EML').setLinkUrl(emlFileUrl).build()
      : SpreadsheetApp.newRichTextValue().setText('').build(),
    driveFileUrl
      ? SpreadsheetApp.newRichTextValue().setText('Invoice').setLinkUrl(driveFileUrl).build()
      : SpreadsheetApp.newRichTextValue().setText('').build(),
    emailLink
      ? SpreadsheetApp.newRichTextValue().setText('Email').setLinkUrl(emailLink).build()
      : SpreadsheetApp.newRichTextValue().setText('').build()
  ]];
  sheet.getRange(lastRow, COL.SOURCE_EML, 1, 3).setRichTextValues(richLinks);

  if (status.indexOf('Flagged') !== -1) {
    sheet.getRange(lastRow, 1, 1, 15).setBackground('#FFF3CD');
  }

  log_('Written row: ' + (data.customer_name || '(unknown vendor)') + ' | billed to ' + (data.billed_to || '') + ' | ' + (data.amount || '') + ' ' + (data.currency || '') + ' | status: ' + status);
  return { duplicate: false, row: sheet.getLastRow() };
}

function findDuplicateRow_(sheet, data, emailLink) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  const emailRichValues = sheet.getRange(2, COL.SOURCE_EMAIL, lastRow - 1, 1).getRichTextValues();

  const newInvoiceNumber = normalizeText_(data.invoice_number);
  const newCustomerName = normalizeText_(data.customer_name);
  const newBilledTo = normalizeText_(data.billed_to);
  const newCurrency = normalizeText_(data.currency);
  const newAmount = normalizeAmount_(data.amount);
  const newInvoiceDate = normalizeDateText_(data.invoice_date);

  for (let i = 0; i < values.length; i++) {
    const rowNum = i + 2;
    const row = values[i];

    const existingInvoiceNumber = normalizeText_(row[COL.INVOICE_NUMBER - 1]);
    const existingInvoiceDate = normalizeDateText_(row[COL.INVOICE_DATE - 1]);
    const existingCustomerName = normalizeText_(row[COL.CUSTOMER_NAME - 1]);
    const existingBilledTo = normalizeText_(row[COL.BILLED_TO - 1]);
    const existingCurrency = normalizeText_(row[COL.CURRENCY - 1]);
    const existingAmount = normalizeAmount_(row[COL.AMOUNT - 1]);
    const existingEmailRich = emailRichValues[i][0];
    const existingEmailLink = existingEmailRich ? (existingEmailRich.getLinkUrl() || '') : '';

    if (emailLink && existingEmailLink && emailLink === existingEmailLink) return rowNum;

    // Strong duplicate: same invoice no + supplier + billed_to
    if (
      newInvoiceNumber && existingInvoiceNumber &&
      newInvoiceNumber === existingInvoiceNumber &&
      newCustomerName && existingCustomerName &&
      newCustomerName === existingCustomerName &&
      newBilledTo && existingBilledTo &&
      newBilledTo === existingBilledTo
    ) {
      return rowNum;
    }

    // Secondary duplicate: supplier + billed_to + amount + invoice date
    if (
      newCustomerName && existingCustomerName &&
      newCustomerName === existingCustomerName &&
      newBilledTo && existingBilledTo &&
      newBilledTo === existingBilledTo &&
      newAmount !== '' && existingAmount !== '' &&
      Number(newAmount) === Number(existingAmount) &&
      newInvoiceDate && existingInvoiceDate &&
      newInvoiceDate === existingInvoiceDate
    ) {
      return rowNum;
    }

    // Fallback duplicate only when billed_to is missing on both sides
    if (
      !newBilledTo && !existingBilledTo &&
      newCustomerName && existingCustomerName &&
      newCustomerName === existingCustomerName &&
      newInvoiceNumber && existingInvoiceNumber &&
      newInvoiceNumber === existingInvoiceNumber &&
      newAmount !== '' && existingAmount !== '' &&
      Number(newAmount) === Number(existingAmount)
    ) {
      return rowNum;
    }

    // Optional exact full match fallback
    if (
      newCustomerName && existingCustomerName &&
      newCustomerName === existingCustomerName &&
      newInvoiceNumber && existingInvoiceNumber &&
      newInvoiceNumber === existingInvoiceNumber &&
      newCurrency && existingCurrency &&
      newCurrency === existingCurrency &&
      newAmount !== '' && existingAmount !== '' &&
      Number(newAmount) === Number(existingAmount) &&
      newInvoiceDate && existingInvoiceDate &&
      newInvoiceDate === existingInvoiceDate &&
      newBilledTo === existingBilledTo
    ) {
      return rowNum;
    }
  }

  return 0;
}

function ensureHeaderRow_(sheet) {
  if (sheet.getLastRow() > 0) {
    const first = String(sheet.getRange(1, 1).getValue() || '').trim();
    if (first === 'Invoice Number') return;
  }

  const headers = [[
    'Invoice Number',
    'Customer Name',
    'Entity Invoiced',
    'Invoice Currency',
    'Total incl. VAT',
    'Invoice Date',
    'Invoice Due Date',
    'Payment Run Date',
    'Status',
    'AI Confidence',
    'Source EML',
    'Source Invoice',
    'Source Email',
    'Date email received',
    'Date Added'
  ]];

  sheet.clear();
  const range = sheet.getRange(1, 1, 1, headers[0].length);
  range.setValues(headers);
  range.setFontWeight('bold');
  range.setBackground('#1a1a2e');
  range.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

function getNextPaymentRunDate_(dateStr) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let base = new Date(today);

  if (dateStr) {
    const parsed = new Date(dateStr);
    if (!isNaN(parsed.getTime())) {
      parsed.setHours(0, 0, 0, 0);
      if (parsed > today) base = parsed;
    }
  }

  for (let i = 0; i <= 7; i++) {
    const candidate = new Date(base);
    candidate.setDate(base.getDate() + i);
    if (CONFIG.PAYMENT_DAYS.indexOf(candidate.getDay()) !== -1) {
      return Utilities.formatDate(candidate, 'Africa/Johannesburg', 'yyyy-MM-dd');
    }
  }

  return '';
}

function getPaymentRunBeforeDueDate_(dueDateStr) {
  const due = new Date(dueDateStr);
  if (isNaN(due.getTime())) return getNextPaymentRunDate_('');
  due.setHours(0, 0, 0, 0);

  // Walk backwards from due date to find the most recent Tuesday (2) or Thursday (4)
  for (let i = 0; i <= 6; i++) {
    const candidate = new Date(due);
    candidate.setDate(due.getDate() - i);
    if (CONFIG.PAYMENT_DAYS.indexOf(candidate.getDay()) !== -1) {
      return Utilities.formatDate(candidate, 'Africa/Johannesburg', 'yyyy-MM-dd');
    }
  }

  return '';
}

// ============================================================
// DIGEST / SETUP
// ============================================================

function accumulateDailyStats_(results) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'Africa/Johannesburg', 'yyyy-MM-dd');

  let stored = { date: '', processed: 0, invoices: 0, flagged: 0, duplicates: 0, pdf_saved: 0, pdf_uploaded: 0, skipped: 0, errors: [] };
  try {
    const raw = props.getProperty('DAILY_STATS');
    if (raw) stored = JSON.parse(raw);
  } catch (e) {}

  if (stored.date !== today) {
    stored = { date: today, processed: 0, invoices: 0, flagged: 0, duplicates: 0, pdf_saved: 0, pdf_uploaded: 0, skipped: 0, errors: [] };
  }

  stored.processed += results.processed;
  stored.invoices += results.invoices;
  stored.flagged += results.flagged;
  stored.duplicates += results.duplicates;
  stored.pdf_saved += results.pdf_saved;
  stored.pdf_uploaded += results.pdf_uploaded;
  stored.skipped += results.skipped;
  stored.errors = stored.errors.concat(results.errors);

  props.setProperty('DAILY_STATS', JSON.stringify(stored));
  log_('Daily stats accumulated: ' + stored.invoices + ' invoices so far today');
}

function sendEndOfDayDigest() {
  const day = new Date().getDay();
  if (day === 0 || day === 6) {
    log_('Weekend — skipping digest.');
    return;
  }
  const props = PropertiesService.getScriptProperties();
  let stored = null;
  try {
    const raw = props.getProperty('DAILY_STATS');
    if (raw) stored = JSON.parse(raw);
  } catch (e) {}

  if (!stored || !stored.date) {
    log_('No daily stats found — nothing to send.');
    return;
  }

  sendDailyDigest_(stored);
  props.deleteProperty('DAILY_STATS');
  log_('End-of-day digest sent and stats cleared.');
}

function sendDailyDigest_(results) {
  const sheetUrl = 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID;
  const subject = '[Lumepay] Invoice scan: ' + results.invoices + ' new | ' + results.flagged + ' flagged | ' + results.duplicates + ' duplicates';

  const body =
    'Daily invoice automation summary - ' +
    Utilities.formatDate(new Date(), 'Africa/Johannesburg', 'dd MMM yyyy HH:mm') +
    '\n\n' +
    'RESULTS\n' +
    '-------\n' +
    'Emails scanned:     ' + results.processed + '\n' +
    'Invoices captured:  ' + results.invoices + '\n' +
    'Flagged for review: ' + results.flagged + '\n' +
    'Duplicates skipped: ' + results.duplicates + '\n' +
    'PDFs saved:         ' + results.pdf_saved + '\n' +
    'PDFs uploaded:      ' + results.pdf_uploaded + '\n' +
    'Skipped:            ' + results.skipped + '\n' +
    'Errors:             ' + results.errors.length + '\n\n' +
    (results.flagged > 0 ? 'Flagged items need manual review.\n\n' : '') +
    (results.errors.length > 0 ? ('ERRORS:\n' + results.errors.map(function(e) { return '- ' + e.subject + ': ' + e.error; }).join('\n') + '\n\n') : '') +
    'View the payment tracker:\n' + sheetUrl + '\n\n' +
    '-\nLumepay Invoice Automation';

  MailApp.sendEmail(CONFIG.DIGEST_RECIPIENTS.join(','), subject, body);
}

function checkSetup() {
  const profileResp = UrlFetchApp.fetch(
    'https://gmail.googleapis.com/gmail/v1/users/me/profile',
    { method: 'get', headers: gmailHeaders_(), muteHttpExceptions: true }
  );

  log_('Gmail profile status: ' + profileResp.getResponseCode());
  log_('Gmail profile response: ' + truncate_(profileResp.getContentText(), 500));
  log_('API key set: ' + (getApiKey_() ? 'YES' : 'NO'));
  log_('Notify email: ' + CONFIG.NOTIFY_EMAIL);
  log_('Anthropic model: ' + CONFIG.ANTHROPIC_MODEL);
  log_('Files beta header: ' + CONFIG.ANTHROPIC_FILES_BETA);
  log_('Processed label write enabled: ' + CONFIG.APPLY_PROCESSED_LABEL);
  log_('Drive root folder configured: ' + (CONFIG.DRIVE_ROOT_FOLDER_ID && CONFIG.DRIVE_ROOT_FOLDER_ID !== 'PASTE_DRIVE_ROOT_FOLDER_ID_HERE' ? 'YES' : 'NO'));
}

function createDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'runDailyInvoiceScan' || fn === 'sendEndOfDayDigest') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  const hours = [6, 8, 10, 12, 15, 17, 19];
  for (let i = 0; i < hours.length; i++) {
    ScriptApp.newTrigger('runDailyInvoiceScan')
      .timeBased()
      .everyDays(1)
      .atHour(hours[i])
      .inTimezone('Africa/Johannesburg')
      .create();
  }

  ScriptApp.newTrigger('sendEndOfDayDigest')
    .timeBased()
    .everyDays(1)
    .atHour(16)
    .nearMinute(45)
    .inTimezone('Africa/Johannesburg')
    .create();

  ScriptApp.newTrigger('sendEndOfDayDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .inTimezone('Africa/Johannesburg')
    .create();

  log_('Scan triggers created for 06:00, 08:00, 10:00, 12:00, 15:00, 17:00, 19:00 Africa/Johannesburg');
  log_('Digest trigger created for ~16:45 Africa/Johannesburg (daily weekdays)');
  log_('Digest trigger created for 07:00 Monday Africa/Johannesburg (weekend backlog)');
}

// ============================================================
// UTILITIES
// ============================================================

function formatEmailReceived_(fields) {
  try {
    if (fields.internalDate) {
      const d = new Date(Number(fields.internalDate));
      if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Africa/Johannesburg', 'yyyy-MM-dd HH:mm');
    }
    if (fields.dateHeader) {
      const d2 = new Date(fields.dateHeader);
      if (!isNaN(d2.getTime())) return Utilities.formatDate(d2, 'Africa/Johannesburg', 'yyyy-MM-dd HH:mm');
    }
  } catch (e) {}
  return '';
}

function parseEmailReceivedDate_(fields) {
  try {
    if (fields.internalDate) {
      const d = new Date(Number(fields.internalDate));
      if (!isNaN(d.getTime())) return d;
    }
    if (fields.dateHeader) {
      const d2 = new Date(fields.dateHeader);
      if (!isNaN(d2.getTime())) return d2;
    }
  } catch (e) {}
  return new Date();
}

function formatCompactDate_(dateObj) {
  return Utilities.formatDate(dateObj, 'Africa/Johannesburg', 'ddMMMyy');
}

function deriveCustomerFromEmail_(fields) {
  const subject = String(fields.subject || '');
  const from = String(fields.from || '');

  const m1 = subject.match(/invoice\s+(?:for|from)\s+(.+)$/i);
  if (m1 && m1[1]) return m1[1];

  const m2 = from.match(/^"?([^"<]+)"?\s*</);
  if (m2 && m2[1]) return m2[1];

  return '';
}

function sanitizeFilename_(name) {
  return String(name || 'document.pdf')
    .replace(/[<>:"|?*\\\/]/g, '-')
    .replace(/[ -]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 200);
}

function sanitizeFilenameComponent_(name) {
  return sanitizeFilename_(String(name || '')).replace(/\.pdf$/i, '').trim() || 'Unknown';
}

function normalizeText_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeAmount_(value) {
  if (value === null || value === undefined || value === '') return '';
  const n = Number(String(value).replace(/[^0-9.-]/g, ''));
  return isNaN(n) ? '' : Number(n.toFixed(2));
}

function normalizeDateText_(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, 'Africa/Johannesburg', 'yyyy-MM-dd');
}

function assertHttpOk_(response, context) {
  const code = response.getResponseCode();
  if (code >= 200 && code < 300) return;
  throw new Error(context + ' | HTTP ' + code + ' | ' + truncate_(response.getContentText(), 1000));
}

function log_(message) {
  Logger.log(message);
}

function safeError_(err) {
  if (!err) return 'Unknown error';
  if (err.message) return String(err.message);
  return String(err);
}

function truncate_(text, maxLen) {
  text = String(text || '');
  maxLen = Number(maxLen || 500);
  return text.length > maxLen ? text.substring(0, maxLen) + '...' : text;
}

function nullToEmpty_(value) {
  return value === null || value === undefined ? '' : value;
}
