/************************************************************
 * GoConstruct Small Jobs Portal (Apps Script Web App)
 * Base: GoConstruct_Portal_Rebuild_2025-12-21.gs
 * Amendments applied: 21 Dec 2025 change-capture "latest amendments only"
 *
 * Key changes baked in:
 * - Single-tab only (no target="_blank" / no window.open)
 * - Cache-busting + no-cache meta (BUILD_VERSION)
 * - Customer public entry shows customer portal only (no worker/manager links)
 * - Customer "Current jobs" hides invoiced jobs (invoice_id present)
 * - Thumbnails up to 5 + same-tab photo viewer page
 * - Same-tab invoice viewer page
 * - Worker "ready for invoice" as large toggle ("Job is complete and ready for billing"), stateful
 * - Client-side image downscale/compress before upload
 * - Manager invoice preview page includes "Back to job (edit)" (same tab)
 * - Invoice email step creates a Gmail Draft with PDF attached (no auto-send)
 * - More robust doPost action inference (prevents "Unknown action")
 * - More robust invoice doc/pdf creation (light retries)
 ************************************************************/

/************************************************************
 * CONFIG
 ************************************************************/
const BUILD_VERSION = '2025-12-22.02';

// Prefer Script Properties for easy handover; fall back to default constant.
const DEFAULT_SPREADSHEET_ID = '1NleL8ZsQi4hrdgUj4hkt50Vbjvxf06YsIJokl8KNlIQ';
const PROP_SPREADSHEET_ID = 'GOCONSTRUCT_SPREADSHEET_ID';

const SHEET_JOBS         = 'Jobs';
const SHEET_INVOICELINES = 'InvoiceLines';

// Jobs sheet column indexes (1-based)
const JOB_COLS = {
  job_id:               1,
  created:              2,
  customer_name:        3,
  customer_email:       4,
  customer_phone:       5,
  site_address:         6,
  priority:             7,
  callout_type:         8,
  status:               9,
  description:          10,
  before_photo_link:    11,
  before_photo_file_id: 12,
  after_photo_link:     13,
  after_photo_file_id:  14,
  hours_on_site:        15,
  workers_count:        16,
  extras_description:   17,
  extras_amount:        18,
  ready_for_invoice:    19,
  invoice_id:           20,
  invoice_status:       21,
  invoice_pdf_url:      22
};

// Invoice lines sheet column indexes (1-based)
const INV_COLS = {
  invoice_id:  1,
  job_id:      2,
  line_no:     3,
  item_code:   4,
  description: 5,
  qty:         6,
  unit_price:  7,
  line_total:  8
};

// Pricing
const HOURLY_RATE = 60;   // £/hour (per labour hour)
const VAT_RATE    = 0.20; // 20% VAT
const CURRENCY    = '£';

// Company details / invoice heading / payment details
const COMPANY = {
  name: 'Go Construct Ltd',
  addressLines: [
    '5 West Gorgie Parks',
    'Edinburgh',
    'EH14 1UT'
  ],
  phone: '+44 (0)0000 000000',
  email: 'info@goconstruct.example',
  website: 'https://goconstruct.example',
  bankName: 'Example Bank',
  bankSortCode: '00-00-00',
  bankAccount: '00000000',
  paymentTermsDays: 7
};

/************************************************************
 * CONFIG HELPERS
 ************************************************************/
function getSpreadsheetId_() {
  // Hardcoded as requested (single source of truth).
  return DEFAULT_SPREADSHEET_ID;
}

function getSS() {
  return SpreadsheetApp.openById(getSpreadsheetId_());
}

/************************************************************
 * URL / CACHE HELPERS
 ************************************************************/
function getAppUrl_() {
  return ScriptApp.getService().getUrl();
}

function withBuildVersion_(url) {
  if (!url) return '';
  const joiner = url.indexOf('?') >= 0 ? '&' : '?';
  // Keep any existing v=; prefer latest value.
  if (/[?&]v=/.test(url)) {
    return url.replace(/([?&]v=)[^&]*/g, '$1' + encodeURIComponent(BUILD_VERSION));
  }
  return url + joiner + 'v=' + encodeURIComponent(BUILD_VERSION);
}

function buildUrl_(mode, params) {
  const base = getAppUrl_();
  const q = [];
  if (mode) q.push('mode=' + encodeURIComponent(mode));
  params = params || {};
  Object.keys(params).forEach(k => {
    const v = params[k];
    if (v === undefined || v === null || v === '') return;
    q.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(v)));
  });
  const url = base + (q.length ? ('?' + q.join('&')) : '');
  return withBuildVersion_(url);
}

function safeReturnUrl_(candidate) {
  const appUrl = getAppUrl_();
  candidate = (candidate || '').toString();
  if (candidate && candidate.indexOf(appUrl) === 0) {
    return withBuildVersion_(candidate);
  }
  return buildUrl_('home', {});
}

function wrapHtmlNoCache_(html) {
  // Insert meta no-cache + a small JS guard against bfcache / back-forward cache.
  // (Apps Script cannot reliably set HTTP headers here.)
  const meta = [
    '<meta http-equiv="Cache-Control" content="no-store, no-cache, must-revalidate, max-age=0">',
    '<meta http-equiv="Pragma" content="no-cache">',
    '<meta http-equiv="Expires" content="0">',
    '<meta name="robots" content="noindex,nofollow">',
    '<script>',
    '  (function(){',
    '    window.addEventListener("pageshow", function(e){',
    '      if (e && e.persisted) {',
    '        try { window.location.replace(window.location.href.replace(/([?&])v=[^&]*/,"$1v=' + BUILD_VERSION + '")); } catch(err) {}',
    '      }',
    '    });',
    '  })();',
    '</script>'
  ].join('\n');
  if (!html) return html;
  if (/<head[^>]*>/i.test(html)) {
    return html.replace(/<head[^>]*>/i, function(m){ return m + '\n' + meta + '\n'; });
  }
  return meta + '\n' + html;
}

/************************************************************
 * BASIC SHEET HELPERS
 ************************************************************/
function getOrCreateSheet(name, headers) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers);
  }
  return sh;
}

function getJobsSheet() {
  return getOrCreateSheet(SHEET_JOBS, [
    'job_id',
    'created',
    'customer_name',
    'customer_email',
    'customer_phone',
    'site_address',
    'priority',
    'callout_type',
    'status',
    'description',
    'before_photo_link',
    'before_photo_file_id',
    'after_photo_link',
    'after_photo_file_id',
    'hours_on_site',
    'workers_count',
    'extras_description',
    'extras_amount',
    'ready_for_invoice',
    'invoice_id',
    'invoice_status',
    'invoice_pdf_url'
  ]);
}

function getInvoiceLinesSheet() {
  return getOrCreateSheet(SHEET_INVOICELINES, [
    'invoice_id',
    'job_id',
    'line_no',
    'item_code',
    'description',
    'qty',
    'unit_price',
    'line_total'
  ]);
}

function generateJobId() {
  const sh = getJobsSheet();
  const last = sh.getLastRow();
  const seq = Math.max(1, last);
  const datePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return 'J' + datePart + '-' + seq;
}

function generateInvoiceId() {
  const sh = getInvoiceLinesSheet();
  const last = sh.getLastRow();
  const seq = Math.max(1, last);
  const datePart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  return 'I' + datePart + '-' + seq.toString().padStart(3, '0');
}

function findJobRow(jobId) {
  const sh = getJobsSheet();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rng = sh.getRange(2, 1, last - 1, sh.getLastColumn());
  const data = rng.getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][JOB_COLS.job_id - 1] === jobId) {
      return { rowIndex: i + 2, row: data[i] };
    }
  }
  return null;
}

function normEmail(v) {
  return String(v || '').trim().toLowerCase();
}

/************************************************************
 * CALLOUT PRICING – FROM "CalloutPricing" SHEET
 ************************************************************/
function getCalloutFee(priority, calloutType) {
  priority = String(priority || '').trim();
  calloutType = String(calloutType || '').trim();

  if (!calloutType || calloutType === 'No call-out') return 0;

  let timing;
  if (calloutType === 'Out of hours call-out') timing = 'Out of hours';
  else timing = 'Within hours';

  const sh = getOrCreateSheet('CalloutPricing', ['Priority', 'Timing', 'Fee']);
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const data = sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const p = String(row[0] || '').trim();
    const t = String(row[1] || '').trim();
    if (p === priority && t === timing) {
      const fee = Number(row[2]) || 0;
      return fee > 0 ? fee : 0;
    }
  }
  return 0;
}

/************************************************************
 * LINK / IMAGE HELPERS
 ************************************************************/
function extractDriveIdFromText(text) {
  if (!text) return '';
  const parts = String(text).split(/\s+/);
  for (let i = 0; i < parts.length; i++) {
    const url = parts[i];
    let m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (m && m[1]) return m[1];
    m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m && m[1]) return m[1];
    m = url.match(/\/d\/([a-zA-Z0-9_-]+)\//);
    if (m && m[1]) return m[1];
  }
  return '';
}

function driveThumbnailUrl(fileId, sizePx) {
  if (!fileId) return '';
  const sz = sizePx ? ('w' + sizePx) : 'w200';
  return 'https://drive.google.com/thumbnail?sz=' + encodeURIComponent(sz) + '&id=' + encodeURIComponent(fileId);
}

function driveImageViewUrl(fileId) {
  if (!fileId) return '';
  // Works for "Anyone with link" files.
  return 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(fileId);
}

/************************************************************
 * DRIVE / BLOB HELPERS (for camera / uploads via base64)
 ************************************************************/
function getJobPhotosFolder(jobId) {
  const ssFile = DriveApp.getFileById(getSpreadsheetId_());
  const parents = ssFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  let jobPhotosFolder;
  const it = parentFolder.getFoldersByName('Job Photos');
  jobPhotosFolder = it.hasNext() ? it.next() : parentFolder.createFolder('Job Photos');

  let jobFolder;
  const it2 = jobPhotosFolder.getFoldersByName(jobId);
  jobFolder = it2.hasNext() ? it2.next() : jobPhotosFolder.createFolder(jobId);

  return jobFolder;
}

function saveBlobToDrive(blob, jobId, prefix) {
  const folder = getJobPhotosFolder(jobId);
  const ts = new Date().getTime();
  const name = prefix + '_' + ts + '.jpg';
  const file = folder.createFile(blob.setName(name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), id: file.getId() };
}

function saveImageFromDataUrl(dataUrl, jobId, prefix) {
  if (!dataUrl) return { url: '', id: '' };
  const m = dataUrl.match(/^data:(.+);base64,(.+)$/);
  if (!m) throw new Error('Invalid image data URL');
  const contentType = m[1];
  const bytes = Utilities.base64Decode(m[2]);
  const blob = Utilities.newBlob(bytes, contentType, prefix + '.jpg');
  return saveBlobToDrive(blob, jobId, prefix);
}

/************************************************************
 * CUSTOMER HELPERS
 ************************************************************/
function getCustomerSummaryByEmail(email) {
  if (!email) return null;
  const sh = getJobsSheet();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const needle = normEmail(email);
  let latest = null;
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (normEmail(r[JOB_COLS.customer_email - 1]) === needle) {
      latest = r;
    }
  }
  if (!latest) return null;
  return {
    customer_name: latest[JOB_COLS.customer_name - 1] || '',
    customer_email: latest[JOB_COLS.customer_email - 1] || '',
    customer_phone: latest[JOB_COLS.customer_phone - 1] || '',
    site_address: latest[JOB_COLS.site_address - 1] || ''
  };
}

function getJobsForCustomerEmail(email) {
  const list = [];
  if (!email) return list;
  const sh = getJobsSheet();
  const last = sh.getLastRow();
  if (last < 2) return list;
  const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const needle = normEmail(email);
  data.forEach(r => {
    if (normEmail(r[JOB_COLS.customer_email - 1]) !== needle) return;

    // Latest amendment: Hide invoiced jobs from "Current jobs" (invoice_id present).
    const invId = r[JOB_COLS.invoice_id - 1];
    if (invId) return;

    list.push({
      job_id: r[JOB_COLS.job_id - 1],
      created: r[JOB_COLS.created - 1],
      description: r[JOB_COLS.description - 1],
      status: r[JOB_COLS.status - 1],
      site_address: r[JOB_COLS.site_address - 1],
      before_ids: (r[JOB_COLS.before_photo_file_id - 1] || '').toString().split('\n').filter(String),
      after_ids:  (r[JOB_COLS.after_photo_file_id  - 1] || '').toString().split('\n').filter(String)
    });
  });
  return list;
}

function getOutstandingInvoicesForCustomerEmail(email) {
  const result = [];
  if (!email) return result;

  const jobsSh = getJobsSheet();
  const lastJobs = jobsSh.getLastRow();
  if (lastJobs < 2) return result;
  const jobsData = jobsSh.getRange(2, 1, lastJobs - 1, jobsSh.getLastColumn()).getValues();
  const needle = normEmail(email);

  const invoiceJobs = {};
  jobsData.forEach(r => {
    const mail = normEmail(r[JOB_COLS.customer_email - 1]);
    if (mail !== needle) return;
    const invId = r[JOB_COLS.invoice_id - 1];
    if (!invId) return;
    const status = r[JOB_COLS.invoice_status - 1] || '';
    if (status === 'Paid') return;
    invoiceJobs[invId] = r;
  });

  if (Object.keys(invoiceJobs).length === 0) return result;

  const invSh = getInvoiceLinesSheet();
  const lastInv = invSh.getLastRow();
  if (lastInv < 2) return result;
  const invData = invSh.getRange(2, 1, lastInv - 1, invSh.getLastColumn()).getValues();

  const totals = {};
  invData.forEach(r => {
    const invId = r[INV_COLS.invoice_id - 1];
    if (!invId || !invoiceJobs[invId]) return;
    const lineTotal = Number(r[INV_COLS.line_total - 1]) || 0;
    totals[invId] = (totals[invId] || 0) + lineTotal;
  });

  Object.keys(totals).forEach(invId => {
    const jobRow = invoiceJobs[invId];
    result.push({
      invoice_id: invId,
      amount: totals[invId],
      job_id: jobRow[JOB_COLS.job_id - 1],
      job_description: jobRow[JOB_COLS.description - 1] || '',
      pdf_url: jobRow[JOB_COLS.invoice_pdf_url - 1] || ''
    });
  });

  return result;
}

/************************************************************
 * ACTION INFERENCE (prevents "Unknown action" on some mobile submits)
 ************************************************************/
function inferAction_(params) {
  params = params || {};
  const action = (params.action || '').toString().trim();
  if (action) return action;

  // Customer create job
  if (params.jobDescription && params.customerEmail) return 'customerNewJob';

  // Worker update job
  if (params.jobId && (params.status || params.hoursOnSite || params.afterPhotoDataJson || params.afterPhotoLink)) {
    return 'workerUpdateJob';
  }

  // Manager draft invoice
  if (params.jobId && params.invoiceId && params.emailTo && params.pdfFileId) {
    return 'managerCreateDraftInvoice';
  }

  return '';
}

/************************************************************
 * HTTP HANDLERS
 ************************************************************/
function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || 'home';

  try {
    let html = '';
    if (mode === 'home') {
      html = buildCustomerStartPage_();
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Small Job Portal')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'customer') {
      const email = (e.parameter && e.parameter.email) || '';
      html = buildCustomerPage_(email);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Customer Portal')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    
    if (mode === 'customerEditJob') {
      const email = (e.parameter && e.parameter.email) || '';
      const jobId = (e.parameter && e.parameter.jobId) || '';
      html = buildCustomerEditJobPage_(email, jobId);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Edit Job')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'customerAddPhotos') {
      const email = (e.parameter && e.parameter.email) || '';
      const jobId = (e.parameter && e.parameter.jobId) || '';
      html = buildCustomerAddPhotosPage_(email, jobId);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Add Photos')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

if (mode === 'worker') {
      const jobId = (e.parameter && e.parameter.jobId) || '';
      html = jobId ? buildJobEditPage_(jobId, 'worker') : buildWorkerListPage_();
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Worker')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'manager') {
      html = buildManagerPage_();
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Manager')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'managerJob') {
      const jobId = (e.parameter && e.parameter.jobId) || '';
      html = buildJobEditPage_(jobId, 'manager');
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Manager Job Edit')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'managerCreateInvoice') {
      const jobId = (e.parameter && e.parameter.jobId) || '';
      if (!jobId) {
        html = buildMessagePage_('Error: No job ID supplied for invoice generation.');
        return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
          .setTitle(COMPANY.name + ' — Error')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      const found = findJobRow(jobId);
      if (!found) {
        html = buildMessagePage_('Error: Job not found for invoice generation (' + jobId + ').');
        return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
          .setTitle(COMPANY.name + ' — Error')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      const result = createInvoiceForJob(jobId);
      html = buildManagerInvoiceConfirmationPage_(found.row, result);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Invoice Created')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'invoiceView') {
      const jobId = (e.parameter && e.parameter.jobId) || '';
      const email = (e.parameter && e.parameter.email) || '';
      html = buildInvoiceViewPage_(jobId, email);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Invoice')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (mode === 'photoView') {
      const fileId = (e.parameter && e.parameter.fileId) || '';
      const returnTo = safeReturnUrl_((e.parameter && e.parameter.returnTo) || '');
      html = buildPhotoViewerPage_(fileId, returnTo);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Photo')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    html = buildMessagePage_('Unknown mode.');
    return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
      .setTitle(COMPANY.name + ' — Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    Logger.log('doGet error: %s', err.stack || err);
    const html = buildMessagePage_('Error: ' + err.message);
    return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
      .setTitle(COMPANY.name + ' — Error')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // ABSOLUTE SAFETY NET — never return nothing
return HtmlService
  .createHtmlOutput(buildMessagePage_('Something went wrong. Please go back and try again.', { backUrl: buildUrl_('home', {}) }))
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}

function doPost(e) {
  const params = (e && e.parameter) || {};
  const action = inferAction_(params);
  let html = '';

  try {
    if (action === 'customerNewJob') {
      const jobId = handleCustomerNewJob(e);
      const email = (params.customerEmail || '').toString().trim();
      const redirectUrl = buildUrl_('customer', { email: email });
      html = buildToastRedirectPage_('Job sent to the team (' + jobId + ').', redirectUrl);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Redirecting')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    
    if (action === 'customerUpdateJob') {
      const jobId = handleCustomerUpdateJob(e);
      const email = (params.customerEmail || params.email || '').toString().trim();
      const redirectUrl = buildUrl_('customer', { email: email });
      html = buildToastRedirectPage_('Job updated (' + jobId + ').', redirectUrl);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Redirecting')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (action === 'customerAddPhotos') {
      const jobId = handleCustomerAddPhotos(e);
      const email = (params.customerEmail || params.email || '').toString().trim();
      const redirectUrl = buildUrl_('customer', { email: email });
      html = buildToastRedirectPage_('Photos added to job ' + jobId + '.', redirectUrl);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Redirecting')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

if (action === 'workerUpdateJob') {
      handleWorkerUpdateJob(e);
      const jobId = (params.jobId || '').toString().trim();
      const role = (params.role || 'worker');
      const redirectUrl = role === 'manager'
        ? buildUrl_('managerJob', { jobId: jobId })
        : buildUrl_('worker', { jobId: jobId });
      html = buildToastRedirectPage_('Job update saved (' + jobId + ').', redirectUrl);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Redirecting')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    if (action === 'managerCreateDraftInvoice') {
      const jobId     = (params.jobId || '').toString().trim();
      const invoiceId = (params.invoiceId || '').toString().trim();
      const pdfFileId = (params.pdfFileId || '').toString().trim();
      const emailTo   = (params.emailTo || '').toString().trim();

      if (!jobId || !invoiceId || !pdfFileId || !emailTo) {
        throw new Error('Missing data to create invoice draft.');
      }

      const subject = 'Invoice ' + invoiceId + ' for your recent ' + COMPANY.name + ' job';

      const pdfUrl = 'https://drive.google.com/file/d/' + pdfFileId + '/preview';
      const portalUrl = buildUrl_('invoiceView', { jobId: jobId, email: emailTo });

      const bodyText =
        'Hi,\n\n' +
        'Thanks again for asking the ' + COMPANY.name + ' team to help with your recent job.\n\n' +
        'Invoice: ' + invoiceId + '\n' +
        'Portal view (keeps the app style): ' + portalUrl + '\n' +
        'PDF link: ' + pdfUrl + '\n\n' +
        'Please make payment within ' + COMPANY.paymentTermsDays + ' days using the invoice number as the payment reference.\n\n' +
        'If anything doesn’t look right, just reply to this email.\n\n' +
        'All the best,\n' +
        COMPANY.name;

      const bodyHtml =
        '<p>Hi,</p>' +
        '<p>Thanks again for asking the <strong>' + COMPANY.name + '</strong> team to help with your recent job.</p>' +
        '<p><strong>Invoice:</strong> ' + invoiceId + '<br>' +
        '<strong>Portal view:</strong> <a href="' + portalUrl + '">Open invoice (same tab style)</a><br>' +
        '<strong>PDF:</strong> <a href="' + pdfUrl + '">View invoice PDF</a></p>' +
        '<p>Use the portal link above from your iPhone, Android, or desktop browser to keep everything in one tab with the portal styling.</p>' +
        '<p>Please make payment within <strong>' + COMPANY.paymentTermsDays + ' days</strong> using the invoice number as the payment reference.</p>' +
        '<p>If anything doesn’t look right, just reply to this email.</p>' +
        '<p>All the best,<br>' + COMPANY.name + '</p>';

      const pdfBlob = DriveApp.getFileById(pdfFileId).getBlob();
      const draft = GmailApp.createDraft(emailTo, subject, bodyText, {
        htmlBody: bodyHtml,
        attachments: [pdfBlob]
      });

      // Update job invoice status as "Drafted" (not "Sent" because it isn't sent yet).
      const job = findJobRow(jobId);
      if (job) {
        const sh = getJobsSheet();
        const row = job.row;
        row[JOB_COLS.invoice_status - 1] = 'Drafted';
        // Keep job status untouched; manager may still need to correct.
        sh.getRange(job.rowIndex, 1, 1, sh.getLastColumn()).setValues([row]);
      }

      const msg =
        'Draft created in Gmail for <strong>' + emailTo + '</strong> with invoice PDF attached.<br><br>' +
        'Open Gmail → <strong>Drafts</strong> and look for subject: <strong>' + subject + '</strong>.';

      const redirectUrl = buildUrl_('manager', {});
      html = buildToastRedirectPage_('Draft created for ' + emailTo + '.', redirectUrl);
      return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
        .setTitle(COMPANY.name + ' — Redirecting')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    html = buildMessagePage_('Unknown action.');
    return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
      .setTitle(COMPANY.name + ' — Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    Logger.log('doPost error: %s', err.stack || err);
    html = buildMessagePage_('Error: ' + err.message);
    return HtmlService.createHtmlOutput(wrapHtmlNoCache_(html))
      .setTitle(COMPANY.name + ' — Error')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  // ABSOLUTE SAFETY NET — never return empty response
return HtmlService
  .createHtmlOutput(
    '<!doctype html><html><head>' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '</head><body>' +
    '<script>' +
    'requestAnimationFrame(function(){' +
    '  window.location.href = "?role=customer";' +
    '});' +
    '</script>' +
    '</body></html>'
  )
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

}

/************************************************************
 * UI HELPERS (shared HTML/CSS blocks)
 ************************************************************/
function sharedStyles_(role) {
  const accent = ({
    home:    '#1976d2',
    customer:'#1b5e20',
    worker:  '#ef6c00',
    manager: '#6a1b9a'
  }[role] || '#1976d2');

  return `
<style>
:root{
  --accent:${accent};
  --bg:#f5f7fb;
  --card:#ffffff;
  --text:#111827;
  --muted:#6b7280;
  --border:rgba(17,24,39,0.12);
  --shadow:0 10px 22px rgba(17,24,39,0.08);
  --radius:16px;
  --navH:64px;
}
*{box-sizing:border-box;}
html,body{height:100%;}
body{
  font-family:Arial,sans-serif;
  padding:14px;
  padding-bottom:calc(var(--navH) + 18px + env(safe-area-inset-bottom));
  max-width:860px;
  margin:0 auto;
  background:var(--bg);
  color:var(--text);
  -webkit-text-size-adjust:100%;
}
h1{margin:10px 0 6px;font-size:22px;line-height:1.2;}
h2{margin:10px 0 8px;font-size:16px;color:var(--text);}
.small{font-size:12px;color:var(--muted);}
.header{
  display:flex;align-items:center;justify-content:space-between;gap:10px;
  padding:10px 12px;border:1px solid var(--border);background:var(--card);
  border-radius:var(--radius);box-shadow:var(--shadow);margin-bottom:12px;
}
.brand{font-size:20px;font-weight:800;letter-spacing:0.2px;}
.badge{
  background:rgba(0,0,0,0.06);
  border:1px solid var(--border);
  padding:6px 10px;border-radius:999px;font-weight:700;font-size:12px;
}
.badge b{color:var(--accent);}
.card{
  border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 14px;margin-bottom:12px;background:var(--card);box-shadow:var(--shadow);
}
.section-title{
  display:flex;align-items:center;gap:8px;margin:0 0 10px;
  font-weight:800;color:var(--accent);
}
.section-title .dot{width:10px;height:10px;border-radius:50%;background:var(--accent);display:inline-block;}
.btn{
  display:inline-block;padding:12px 16px;margin:6px 0;background:var(--accent);
  color:#fff;text-decoration:none;border-radius:999px;border:none;cursor:pointer;
  font-size:16px;font-weight:800;text-align:center;
  min-height:44px;
}
.btn-secondary{background:#455a64;}
.btn-ghost{background:transparent;color:var(--accent);border:1px solid var(--accent);}
.btn[disabled]{opacity:0.65;cursor:not-allowed;}
input,textarea,select{
  width:100%;
  padding:10px;
  margin:4px 0 12px;
  border-radius:10px;
  border:1px solid var(--border);
  font-size:16px; /* iOS: prevent zoom */
  background:#fff;
}
label{font-weight:800;font-size:14px;display:block;margin-top:2px;color:#111;}
.table{width:100%;border-collapse:collapse;font-size:13px;}
.table th,.table td{border:1px solid rgba(0,0,0,0.10);padding:6px;vertical-align:top;}
.table th{background:rgba(0,0,0,0.04);font-weight:800;}
.thumbstrip{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;}
.thumb{width:26px;height:26px;border-radius:6px;border:1px solid rgba(0,0,0,0.15);object-fit:cover;}
.thumb-lg{width:44px;height:44px;border-radius:10px;border:1px solid rgba(0,0,0,0.15);object-fit:cover;}
.toggle{
  width:100%;
  padding:14px 14px;border-radius:14px;
  border:2px solid rgba(0,0,0,0.10);
  background:rgba(0,0,0,0.04);
  font-weight:900;font-size:16px;text-align:center;
  min-height:56px;
}
.toggle.on{border-color:var(--accent);background:rgba(25,118,210,0.10);}
hr{border:none;border-top:1px solid rgba(0,0,0,0.10);margin:12px 0;}

/* Whole-card tap targets */
.cardlink{display:block;color:inherit;text-decoration:none;}
.cardlink:active{transform:scale(0.997);}

/* Sticky primary action bar */
.stickybar{
  position:sticky;
  bottom:calc(var(--navH) + env(safe-area-inset-bottom) + 10px);
  z-index:50;
  margin-top:12px;
}
.stickybar .btn{width:100%; margin:0;}

/* Toast */
.toast{
  position:fixed;
  left:50%;
  transform:translateX(-50%);
  bottom:calc(var(--navH) + 14px + env(safe-area-inset-bottom));
  z-index:9999;
  min-width:220px;
  max-width:min(560px, calc(100vw - 24px));
  background:rgba(17,24,39,0.92);
  color:#fff;
  padding:10px 12px;
  border-radius:999px;
  font-size:14px;
  font-weight:800;
  text-align:center;
  opacity:0;
  pointer-events:none;
  transition:opacity 160ms ease;
}
.toast.show{opacity:1;}

/* Bottom nav (mobile only) */
.bottomnav{
  position:fixed;
  left:0; right:0;
  bottom:0;
  z-index:9998;
  background:rgba(255,255,255,0.92);
  backdrop-filter:saturate(1.2) blur(8px);
  border-top:1px solid rgba(0,0,0,0.10);
  padding:8px 10px calc(8px + env(safe-area-inset-bottom));
}
.bottomnav .wrap{
  max-width:860px;
  margin:0 auto;
  display:flex;
  gap:8px;
  justify-content:space-between;
}
.bottomnav a{
  flex:1;
  display:block;
  text-decoration:none;
  color:var(--text);
  background:rgba(0,0,0,0.04);
  border:1px solid rgba(0,0,0,0.10);
  border-radius:14px;
  padding:10px 8px;
  text-align:center;
  font-weight:900;
  font-size:13px;
  min-height:44px;
}
.bottomnav a.primary{
  background:rgba(25,118,210,0.12);
  border-color:rgba(25,118,210,0.25);
  color:var(--accent);
}
@media (min-width: 820px){
  .bottomnav{display:none;}
  body{padding-bottom:18px;}
  .stickybar{bottom:10px;}
  .toast{bottom:14px;}
}


*{box-sizing:border-box;}
body{font-family:Arial,sans-serif;padding:14px;max-width:860px;margin:0 auto;background:var(--bg);color:var(--text);}
h1{margin:10px 0 6px;font-size:22px;line-height:1.2;}
h2{margin:10px 0 8px;font-size:16px;color:var(--text);}
.small{font-size:12px;color:var(--muted);}
.header{
  display:flex;align-items:center;justify-content:space-between;gap:10px;
  padding:10px 12px;border:1px solid var(--border);background:var(--card);
  border-radius:var(--radius);box-shadow:var(--shadow);margin-bottom:12px;
}
.brand{font-size:20px;font-weight:800;letter-spacing:0.2px;}
.badge{
  background:rgba(0,0,0,0.06);
  border:1px solid var(--border);
  padding:6px 10px;border-radius:999px;font-weight:700;font-size:12px;
}
.badge b{color:var(--accent);}
.card{
  border:1px solid var(--border);border-radius:var(--radius);
  padding:14px 14px;margin-bottom:12px;background:var(--card);box-shadow:var(--shadow);
}
.section-title{
  display:flex;align-items:center;gap:8px;margin:0 0 10px;
  font-weight:800;color:var(--accent);
}
.section-title .dot{width:10px;height:10px;border-radius:50%;background:var(--accent);display:inline-block;}
.btn{
  display:inline-block;padding:12px 16px;margin:6px 0;background:var(--accent);
  color:#fff;text-decoration:none;border-radius:999px;border:none;cursor:pointer;
  font-size:16px;font-weight:800;text-align:center;
}
.btn-secondary{background:#455a64;}
.btn-ghost{
  background:transparent;color:var(--accent);border:1px solid var(--accent);
}
input,textarea,select{
  width:100%;padding:10px;margin:4px 0 12px;border-radius:10px;border:1px solid var(--border);
  font-size:16px;background:#fff;
}
label{font-weight:800;font-size:14px;display:block;margin-top:2px;color:#111;}
.table{width:100%;border-collapse:collapse;font-size:13px;}
.table th,.table td{border:1px solid rgba(0,0,0,0.10);padding:6px;vertical-align:top;}
.table th{background:rgba(0,0,0,0.04);font-weight:800;}
.thumbstrip{display:flex;flex-wrap:wrap;gap:4px;margin-top:6px;}
.thumb{
  width:26px;height:26px;border-radius:6px;border:1px solid rgba(0,0,0,0.15);object-fit:cover;
}
.thumb-lg{
  width:44px;height:44px;border-radius:10px;border:1px solid rgba(0,0,0,0.15);object-fit:cover;
}
.toggle{
  width:100%;
  padding:14px 14px;border-radius:14px;
  border:2px solid rgba(0,0,0,0.10);
  background:rgba(0,0,0,0.04);
  font-weight:900;font-size:16px;text-align:center;
}
.toggle.on{
  border-color:var(--accent);
  background:rgba(25,118,210,0.10);
}
hr{border:none;border-top:1px solid rgba(0,0,0,0.10);margin:12px 0;}
</style>`;
}

function headerBar_(role, titleRight) {
  titleRight = titleRight || '';
  const label = ({
    home: 'Home',
    customer: 'Customer',
    worker: 'Worker',
    manager: 'Manager'
  }[role] || 'Portal');

  return `
<div class="header">
  <div class="brand">${escapeHtml_(COMPANY.name)}</div>
  <div class="badge"><b>${escapeHtml_(label)}</b>${titleRight ? ' — ' + escapeHtml_(titleRight) : ''}</div>
</div>`;
}

function buildToastRedirectPage_(message, redirectUrl) {
  const msgJson = JSON.stringify(message || '');
  const urlJson = JSON.stringify(redirectUrl || buildUrl_('home', {}));
  const safeUrl = redirectUrl || buildUrl_('home', {});
  return `
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="refresh" content="0;url=${escapeHtml_(safeUrl)}">
${sharedStyles_('home')}
${appFeelScript_()}
</head>
<body>
<script>
  (function(){
    try{ sessionStorage.setItem('GCFlash', ${msgJson}); }catch(e){}
    var url = ${urlJson} || '${escapeHtml_(buildUrl_('home', {}))}';
    try{ window.location.replace(url); }catch(err){ window.location.href = url; }
    setTimeout(function(){ try{ window.location.href = url; }catch(e){} }, 600);
  })();
</script>
<div style="padding:18px;font-weight:800;">Redirecting…</div>
<div class="small" style="padding:0 18px 12px;">If nothing happens, <a href="${escapeHtml_(safeUrl)}" target="_self">tap here</a>.</div>
</body>
</html>`;
}


function appFeelScript_() {
  // Lightweight mobile "app feel": toast + disable double-taps + scroll-to-error.
  return `
<script>
(function(){
  function ensureToast(){
    var t = document.getElementById('toast');
    if (!t){
      t = document.createElement('div');
      t.id = 'toast';
      t.className = 'toast';
      document.body.appendChild(t);
    }
    return t;
  }
  window.GCToast = function(msg, ms){
    try{
      var t = ensureToast();
      t.textContent = msg || '';
      t.classList.add('show');
      setTimeout(function(){ t.classList.remove('show'); }, ms || 1600);
    }catch(e){}
  };

  function playFlash(){
    try{
      var msg = sessionStorage.getItem('GCFlash');
      if (msg){
        sessionStorage.removeItem('GCFlash');
        window.GCToast(msg, 2000);
      }
    }catch(e){}
  }
  document.addEventListener('DOMContentLoaded', playFlash);

  window.GCFlashAndGo = function(msg, url){
    try{ sessionStorage.setItem('GCFlash', msg || ''); }catch(e){}
    try{ window.location.replace(url); }catch(e){ window.location.href = url; }
  };

  // Enhance forms: prevent double-tap + show toast.
  document.addEventListener('submit', function(ev){
    var form = ev.target;
    if (!form || form.getAttribute('data-gc-enhance') !== '1') return;
    try{
      // basic invalid handling
      var firstInvalid = form.querySelector(':invalid');
      if (firstInvalid){
        ev.preventDefault();
        firstInvalid.scrollIntoView({behavior:'smooth', block:'center'});
        firstInvalid.focus({preventScroll:true});
        window.GCToast('Check the highlighted fields', 1800);
        return;
      }
    }catch(e){}

    try{
      if (form.getAttribute('data-gc-submitting') === '1'){
        ev.preventDefault();
        return;
      }
      form.setAttribute('data-gc-submitting','1');
      var btn = form.querySelector('button[type="submit"],input[type="submit"]');
      if (btn){
        btn.disabled = true;
        btn.setAttribute('data-gc-old', btn.textContent || '');
        btn.textContent = (btn.textContent || 'Submit') + '…';
      }
      window.GCToast('Working…', 1200);
    }catch(e){}
  }, true);
})();
</script>`;
}

function bottomNav_(role, ctx) {
  ctx = ctx || {};
  let items = [];

  if (role === 'customer') {
    items = [
      { label: 'Jobs', href: '#jobs', primary: false },
      { label: 'New Job', href: '#newjob', primary: true },
      { label: 'Invoices', href: '#invoices', primary: false }
    ];
  } else if (role === 'worker') {
    items = [
      { label: 'Jobs', href: buildUrl_('worker', {}), primary: true },
      { label: 'Home', href: buildUrl_('home', {}), primary: false }
    ];
  } else if (role === 'manager') {
    items = [
      { label: 'Jobs', href: buildUrl_('manager', {}), primary: true },
      { label: 'Home', href: buildUrl_('home', {}), primary: false }
    ];
  } else {
    items = [
      { label: 'Home', href: buildUrl_('home', {}), primary: true }
    ];
  }

  return `
<div class="bottomnav" role="navigation" aria-label="Navigation">
  <div class="wrap">
    ${items.map(it => `<a target="_self" class="${it.primary ? 'primary' : ''}" href="${it.href}">${escapeHtml_(it.label)}</a>`).join('')}
  </div>
</div>`;
}


function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function renderThumbStrip_(fileIds, returnToUrl, sizeClass) {
  fileIds = (fileIds || []).filter(String).slice(0, 5);
  if (!fileIds.length) return '';
  sizeClass = sizeClass || 'thumb';
  const appUrl = getAppUrl_();
  const returnTo = safeReturnUrl_(returnToUrl || '');
  return `
<div class="thumbstrip">
  ${fileIds.map(id => {
    const viewUrl = buildUrl_('photoView', { fileId: id, returnTo: returnTo });
    return `<a href="${viewUrl}" target="_self"><img class="${sizeClass}" src="${driveThumbnailUrl(id, 200)}" alt=""></a>`;
  }).join('')}
</div>`;
}

/************************************************************
 * PAGES
 ************************************************************/
function buildCustomerStartPage_() {
  const appUrl = getAppUrl_();
  const actionUrl = buildUrl_('customer', {});
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Small Job Portal</title>
${sharedStyles_('home')}
${appFeelScript_()}
</head>
<body>
${headerBar_('home', 'Small Job Portal')}

<div class="card">
  <div class="section-title"><span class="dot"></span>Customer portal</div>
  <form method="get" action="${actionUrl}" target="_self">
    <input type="hidden" name="mode" value="customer">
    <label>Your invoice email</label>
    <input type="email" name="email" placeholder="you@example.com" required>
    <button class="btn" type="submit">Continue</button>
    <div class="small">This portal shows your jobs and invoices and lets you create a new job.</div>
  </form>
</div>

</body>
</html>`;
}

function buildCustomerPage_(email) {
  const trimmedEmail = (email || '').toString().trim();
  const summary = trimmedEmail ? getCustomerSummaryByEmail(trimmedEmail) : null;
  const jobs = trimmedEmail ? getJobsForCustomerEmail(trimmedEmail) : [];
  const invoices = trimmedEmail ? getOutstandingInvoicesForCustomerEmail(trimmedEmail) : [];

  const backUrl = buildUrl_('home', {});
  const returnToCustomer = buildUrl_('customer', { email: trimmedEmail });

  let jobsHtml = '';
  if (!jobs.length) {
    jobsHtml = '<div class="small">No current jobs.</div>';
  } else {
    jobsHtml = jobs.map(j => {
      const beforeStrip = renderThumbStrip_(j.before_ids, returnToCustomer, 'thumb');
      const afterStrip  = renderThumbStrip_(j.after_ids, returnToCustomer, 'thumb');
      return `
<div style="padding:8px 0;border-top:1px solid rgba(0,0,0,0.08);">
  <div style="display:flex;align-items:flex-start;gap:10px;justify-content:space-between;">
    <div style="flex:1;">
      <div style="font-weight:900;">${escapeHtml_(j.description || 'Job')}</div>
      <div class="small">${escapeHtml_(j.status || '')}</div>
      <div class="small">${escapeHtml_(j.site_address || '')}</div>
      <div class="small">Created: ${escapeHtml_(j.created || '')}</div>
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap;">
        <a class="btn btn-ghost" href="${buildUrl_('customerEditJob',{ email: trimmedEmail, jobId: j.job_id })}" target="_self">Edit job</a>
        <a class="btn btn-ghost" href="${buildUrl_('customerAddPhotos',{ email: trimmedEmail, jobId: j.job_id })}" target="_self">Add photos</a>
      </div>
    </div>
  </div>
  ${beforeStrip ? `<div class="small" style="margin-top:6px;font-weight:800;">Before</div>${beforeStrip}` : ''}
  ${afterStrip ? `<div class="small" style="margin-top:6px;font-weight:800;">After</div>${afterStrip}` : ''}
</div>`;
    }).join('');
  }

  let invHtml = '';
  if (!invoices.length) {
    invHtml = '<div class="small">No outstanding invoices.</div>';
  } else {
    invHtml = invoices.map(inv => {
      const jobId = inv.job_id;
      const viewUrl = buildUrl_('invoiceView', { jobId: jobId, email: trimmedEmail });
      const amountStr = CURRENCY + Number(inv.amount || 0).toFixed(2);
      return `
<div style="padding:10px 0;border-top:1px solid rgba(0,0,0,0.08);display:flex;gap:10px;align-items:center;">
  <div style="flex:1;">
    <div style="font-weight:900;">${escapeHtml_(inv.invoice_id)}</div>
    <div class="small">${escapeHtml_(inv.job_description || ('Job ' + jobId))}</div>
    <div class="small">Total (incl. VAT): <b>${escapeHtml_(amountStr)}</b></div>
  </div>
  <a class="btn btn-ghost" href="${viewUrl}" target="_self">View</a>
</div>`;
    }).join('');
  }

  const jobFormHtml = summary ? buildCustomerJobFormKnown_(summary, trimmedEmail) : buildCustomerJobFormFirst_(trimmedEmail);

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Customer</title>
${sharedStyles_('customer')}
${appFeelScript_()}
</head>
<body>
${headerBar_('customer', trimmedEmail ? trimmedEmail : '')}

<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

${summary ? `
<div class="card">
  <div class="section-title"><span class="dot"></span>Your details</div>
  <div style="font-weight:900;">${escapeHtml_(summary.customer_name || 'Customer')}</div>
  <div class="small">${escapeHtml_(summary.site_address || '')}</div>
  <div class="small">${escapeHtml_(summary.customer_email || '')}${summary.customer_phone ? ' | ' + escapeHtml_(summary.customer_phone) : ''}</div>
</div>

<div class="card" id="jobs">
  <div class="section-title"><span class="dot"></span>Current jobs</div>
  <div class="small">Jobs that have already been invoiced are shown under “Outstanding invoices”.</div>
  ${jobsHtml}
</div>

<div class="card" id="invoices">
  <div class="section-title"><span class="dot"></span>Outstanding invoices</div>
  ${invHtml}
</div>

${jobFormHtml}
` : `
<div class="card">
  <div class="section-title"><span class="dot"></span>Create your first job</div>
  <div class="small">We don’t recognise that email yet. Fill this in once and we’ll save your details.</div>
</div>
${jobFormHtml}
`}


${bottomNav_('customer', {email: trimmedEmail})}
</body>
</html>`;
}

function buildCustomerJobFormKnown_(summary, email) {
  const appUrl = getAppUrl_();
  const postUrl = appUrl;
  return `
<div class="card" id="newjob">
  <div class="section-title"><span class="dot"></span>New job</div>
  <div class="small">Attach photos if you want. Images are compressed before upload for reliability.</div>

  <form id="customerJobForm" method="post" action="${postUrl}" target="_self" data-gc-enhance="1">
    <input type="hidden" name="action" value="customerNewJob">
    <input type="hidden" name="customerName" value="${escapeHtml_(summary.customer_name)}">
    <input type="hidden" name="customerEmail" value="${escapeHtml_(email)}">
    <input type="hidden" name="customerPhone" value="${escapeHtml_(summary.customer_phone)}">
    <input type="hidden" name="siteAddress" value="${escapeHtml_(summary.site_address)}">
    <input type="hidden" name="beforePhotoDataJson" id="beforePhotoDataJson">
    

    <label>Describe the job</label>
    <textarea name="jobDescription" rows="3" required placeholder="What do you need us to do?"></textarea>

    <label>Priority</label>
    <select name="priority">
      <option value="Normal" selected>Normal</option>
      <option value="Urgent">Urgent</option>
      <option value="Emergency">Emergency</option>
    </select>

    <label>Photos (optional)</label>
    <div class="card" style="padding:12px;border-radius:14px;background:rgba(0,0,0,0.02);">
      <button type="button" class="btn" id="btnTakePhoto" style="width:100%;">Take photo</button>
      <button type="button" class="btn btn-ghost" id="btnChoosePhotos" style="width:100%;">Choose photos</button>
      <div class="small" id="photoPickHint">No photos selected.</div>
      <input type="file" id="beforePhotoCamera" accept="image/*" capture="environment" style="display:none;">
      <input type="file" id="beforePhotoLibrary" accept="image/*" multiple style="display:none;">
    </div>

    <label>Or paste photo links (optional)</label>
    <textarea name="beforePhotoLink" rows="2" placeholder="Google Photos / Drive / WhatsApp Web etc."></textarea>

    <div class="stickybar">
      <button class="btn" type="submit" id="customerSubmitBtn">Send job to the team</button>
    </div>
  </form>
</div>
${clientImageCompressScriptMulti_('customerJobForm',['beforePhotoCamera','beforePhotoLibrary'],'beforePhotoDataJson','customerSubmitBtn','photoPickHint')}
`;
}


function buildCustomerJobFormFirst_(email) {
  const appUrl = getAppUrl_();
  const postUrl = appUrl;
  return `
<div class="card" id="newjob">
  <div class="section-title"><span class="dot"></span>New job</div>
  <div class="small">Attach photos if you want. Images are compressed before upload for reliability.</div>

  <form id="customerJobForm" method="post" action="${postUrl}" target="_self" data-gc-enhance="1">
    <input type="hidden" name="action" value="customerNewJob">
    <input type="hidden" name="beforePhotoDataJson" id="beforePhotoDataJson">
    
    <label>Your name</label>
    <input type="text" name="customerName" required>

    <label>Your email</label>
    <input type="email" name="customerEmail" value="${escapeHtml_(email)}" required>

    <label>Your phone</label>
    <input type="text" name="customerPhone">

    <label>Address where work is needed</label>
    <textarea name="siteAddress" rows="2" required></textarea>
    

    <label>Describe the job</label>
    <textarea name="jobDescription" rows="3" required placeholder="What do you need us to do?"></textarea>

    <label>Priority</label>
    <select name="priority">
      <option value="Normal" selected>Normal</option>
      <option value="Urgent">Urgent</option>
      <option value="Emergency">Emergency</option>
    </select>

    <label>Photos (optional)</label>
    <div class="card" style="padding:12px;border-radius:14px;background:rgba(0,0,0,0.02);">
      <button type="button" class="btn" id="btnTakePhoto" style="width:100%;">Take photo</button>
      <button type="button" class="btn btn-ghost" id="btnChoosePhotos" style="width:100%;">Choose photos</button>
      <div class="small" id="photoPickHint">No photos selected.</div>
      <input type="file" id="beforePhotoCamera" accept="image/*" capture="environment" style="display:none;">
      <input type="file" id="beforePhotoLibrary" accept="image/*" multiple style="display:none;">
    </div>

    <label>Or paste photo links (optional)</label>
    <textarea name="beforePhotoLink" rows="2" placeholder="Google Photos / Drive / WhatsApp Web etc."></textarea>

    <div class="stickybar">
      <button class="btn" type="submit" id="customerSubmitBtn">Create job</button>
    </div>
  </form>
</div>
${clientImageCompressScriptMulti_('customerJobForm',['beforePhotoCamera','beforePhotoLibrary'],'beforePhotoDataJson','customerSubmitBtn','photoPickHint')}
`;
}



function assertCustomerOwnsJob_(email, jobId) {
  const found = findJobRow(jobId);
  if (!found) throw new Error('Job not found: ' + jobId);
  const row = found.row;
  const jobEmail = normEmail(row[JOB_COLS.customer_email - 1]);
  if (!email || normEmail(email) !== jobEmail) throw new Error('Not authorised for this job.');
  const invId = row[JOB_COLS.invoice_id - 1];
  if (invId) throw new Error('This job has been invoiced and can no longer be edited.');
  return found;
}

function buildCustomerEditJobPage_(email, jobId) {
  const trimmedEmail = (email || '').toString().trim();
  const found = findJobRow(jobId);
  if (!jobId || !found) {
    return buildMessagePage_('Error: Job not found.', { backUrl: buildUrl_('customer', { email: trimmedEmail }) });
  }
  const row = found.row;
  const jobEmail = row[JOB_COLS.customer_email - 1] || '';
  const invId = row[JOB_COLS.invoice_id - 1] || '';
  if (normEmail(jobEmail) !== normEmail(trimmedEmail)) {
    return buildMessagePage_('Error: Not authorised for this job.', { backUrl: buildUrl_('customer', { email: trimmedEmail }) });
  }
  if (invId) {
    return buildMessagePage_('This job has already been invoiced and can no longer be edited.', { backUrl: buildUrl_('customer', { email: trimmedEmail }) });
  }

  const desc = row[JOB_COLS.description - 1] || '';
  const priority = row[JOB_COLS.priority - 1] || 'Normal';
  const backUrl = buildUrl_('customer', { email: trimmedEmail });

  const postUrl = getAppUrl_();

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Edit Job</title>
${sharedStyles_('customer')}
${appFeelScript_()}
</head>
<body>
${headerBar_('customer', 'Edit job')}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Edit job</div>
  <div class="small">Job: <b>${escapeHtml_(jobId)}</b></div>

  <form method="post" action="${postUrl}" target="_self" data-gc-enhance="1">
    <input type="hidden" name="action" value="customerUpdateJob">
    <input type="hidden" name="jobId" value="${escapeHtml_(jobId)}">
    <input type="hidden" name="customerEmail" value="${escapeHtml_(trimmedEmail)}">

    <label>Describe the job</label>
    <textarea name="jobDescription" rows="4" required>${escapeHtml_(desc)}</textarea>

    <label>Priority</label>
    <select name="priority">
      <option value="Normal" ${priority==='Normal'?'selected':''}>Normal</option>
      <option value="Urgent" ${priority==='Urgent'?'selected':''}>Urgent</option>
      <option value="Emergency" ${priority==='Emergency'?'selected':''}>Emergency</option>
    </select>

    <div class="stickybar">
      <button class="btn" type="submit" id="custEditSubmit">Save changes</button>
    </div>
  </form>
</div>

<div class="card">
  <div class="section-title"><span class="dot"></span>Add photos</div>
  <div class="small">Add more photos to this job (before invoicing).</div>
  <a class="btn btn-ghost" href="${buildUrl_('customerAddPhotos',{email: trimmedEmail, jobId: jobId})}" target="_self">Add photos</a>
</div>

${bottomNav_('customer', {email: trimmedEmail})}
</body>
</html>`;
}

function buildCustomerAddPhotosPage_(email, jobId) {
  const trimmedEmail = (email || '').toString().trim();
  const found = findJobRow(jobId);
  if (!jobId || !found) {
    return buildMessagePage_('Error: Job not found.', { backUrl: buildUrl_('customer', { email: trimmedEmail }) });
  }
  const row = found.row;
  const jobEmail = row[JOB_COLS.customer_email - 1] || '';
  const invId = row[JOB_COLS.invoice_id - 1] || '';
  if (normEmail(jobEmail) !== normEmail(trimmedEmail)) {
    return buildMessagePage_('Error: Not authorised for this job.', { backUrl: buildUrl_('customer', { email: trimmedEmail }) });
  }
  if (invId) {
    return buildMessagePage_('This job has already been invoiced and can no longer accept new photos.', { backUrl: buildUrl_('customer', { email: trimmedEmail }) });
  }

  const backUrl = buildUrl_('customer', { email: trimmedEmail });
  const postUrl = getAppUrl_();

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Add Photos</title>
${sharedStyles_('customer')}
${appFeelScript_()}
</head>
<body>
${headerBar_('customer', 'Add photos')}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Add photos</div>
  <div class="small">Job: <b>${escapeHtml_(jobId)}</b></div>

  <form id="custAddPhotosForm" method="post" action="${postUrl}" target="_self" data-gc-enhance="1">
    <input type="hidden" name="action" value="customerAddPhotos">
    <input type="hidden" name="jobId" value="${escapeHtml_(jobId)}">
    <input type="hidden" name="customerEmail" value="${escapeHtml_(trimmedEmail)}">
    <input type="hidden" name="beforePhotoDataJson" id="custMorePhotosJson">

    <label>Photos</label>
    <div class="card" style="padding:12px;border-radius:14px;background:rgba(0,0,0,0.02);">
      <button type="button" class="btn" id="btnTakePhoto" style="width:100%;">Take photo</button>
      <button type="button" class="btn btn-ghost" id="btnChoosePhotos" style="width:100%;">Choose photos</button>
      <div class="small" id="photoPickHint">No photos selected.</div>
      <input type="file" id="beforePhotoCamera" accept="image/*" capture="environment" style="display:none;">
      <input type="file" id="beforePhotoLibrary" accept="image/*" multiple style="display:none;">
    </div>

    <div class="stickybar">
      <button class="btn" type="submit" id="custAddPhotosBtn">Upload photos</button>
    </div>
  </form>
</div>

${clientImageCompressScriptMulti_('custAddPhotosForm',['beforePhotoCamera','beforePhotoLibrary'],'custMorePhotosJson','custAddPhotosBtn','photoPickHint')}

${bottomNav_('customer', {email: trimmedEmail})}
</body>
</html>`;
}

function buildWorkerListPage_() {
  const sh = getJobsSheet();
  const last = sh.getLastRow();
  let rowsHtml = '';

  const appUrl = getAppUrl_();
  const backUrl = buildUrl_('home', {});
  const role = 'worker';

  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    rowsHtml = data.map(r => {
      const jobId = r[JOB_COLS.job_id - 1];
      if (!jobId) return '';

      const customer = r[JOB_COLS.customer_name - 1] || '';
      const addr = r[JOB_COLS.site_address - 1] || '';
      const status = r[JOB_COLS.status - 1] || '';
      const priority = r[JOB_COLS.priority - 1] || '';
      const invStatus = r[JOB_COLS.invoice_status - 1] || '';

      // Keep the list lean: hide complete jobs that are sent/paid
      if (status === 'Complete' && (invStatus === 'Sent' || invStatus === 'Paid')) return '';

      const jobUrl = buildUrl_('worker', { jobId: jobId });

      const beforeIds = (r[JOB_COLS.before_photo_file_id - 1] || '').toString().split('\n').filter(String);
      const afterIds  = (r[JOB_COLS.after_photo_file_id  - 1] || '').toString().split('\n').filter(String);

      const beforeStrip = renderThumbStrip_(beforeIds, buildUrl_('worker', {}), 'thumb');
      const afterStrip  = renderThumbStrip_(afterIds, buildUrl_('worker', {}), 'thumb');

      return `
<tr>
  <td>${escapeHtml_(jobId)}</td>
  <td>${escapeHtml_(customer)}</td>
  <td>${escapeHtml_(addr)}</td>
  <td>${escapeHtml_(status)}</td>
  <td>${escapeHtml_(priority)}</td>
  <td>${beforeStrip || ''}</td>
  <td>${afterStrip || ''}</td>
  <td><a class="btn" href="${jobUrl}" target="_self">Open</a></td>
</tr>`;
    }).join('');
  }

  if (!rowsHtml) rowsHtml = '<tr><td colspan="8">No jobs found.</td></tr>';

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Worker</title>
${sharedStyles_(role)}
</head>
<body>
${headerBar_(role, 'Jobs')}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Jobs</div>
  <table class="table">
    <tr>
      <th>Job ID</th>
      <th>Customer</th>
      <th>Site</th>
      <th>Status</th>
      <th>Priority</th>
      <th>Before</th>
      <th>After</th>
      <th></th>
    </tr>
    ${rowsHtml}
  </table>
</div>


${bottomNav_('worker', {})}
</body>
</html>`;
}

function buildManagerPage_() {
  const sh = getJobsSheet();
  const last = sh.getLastRow();
  let rowsHtml = '';

  const backUrl = buildUrl_('home', {});

  if (last >= 2) {
    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    rowsHtml = data.map(r => {
      const jobId = r[JOB_COLS.job_id - 1];
      if (!jobId) return '';

      const customer= r[JOB_COLS.customer_name - 1] || '';
      const addr    = r[JOB_COLS.site_address - 1] || '';
      const status  = r[JOB_COLS.status - 1] || '';
      const ready   = r[JOB_COLS.ready_for_invoice - 1] || '';
      const invId   = r[JOB_COLS.invoice_id - 1] || '';
      const invStatus = r[JOB_COLS.invoice_status - 1] || '';

      const canInv = (ready === 'Yes');
      const invUrl = buildUrl_('managerCreateInvoice', { jobId: jobId });
      const editUrl = buildUrl_('managerJob', { jobId: jobId });

      const beforeIds = (r[JOB_COLS.before_photo_file_id - 1] || '').toString().split('\n').filter(String);
      const afterIds  = (r[JOB_COLS.after_photo_file_id  - 1] || '').toString().split('\n').filter(String);
      const beforeStrip = renderThumbStrip_(beforeIds, buildUrl_('manager', {}), 'thumb');
      const afterStrip  = renderThumbStrip_(afterIds, buildUrl_('manager', {}), 'thumb');

      return `
<tr>
  <td>${escapeHtml_(jobId)}</td>
  <td>${escapeHtml_(customer)}</td>
  <td>${escapeHtml_(addr)}</td>
  <td>${escapeHtml_(status)}</td>
  <td>${escapeHtml_(ready)}</td>
  <td>${escapeHtml_(invId)}</td>
  <td>${escapeHtml_(invStatus)}</td>
  <td>${beforeStrip || ''}</td>
  <td>${afterStrip || ''}</td>
  <td>
    <a class="btn btn-ghost" href="${editUrl}" target="_self">Edit job</a>
    ${canInv ? `<a class="btn" href="${invUrl}" target="_self">Create / amend invoice</a>` : ''}
  </td>
</tr>`;
    }).join('');
  }

  if (!rowsHtml) rowsHtml = '<tr><td colspan="10">No jobs found.</td></tr>';

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Manager</title>
${sharedStyles_('manager')}
${appFeelScript_()}
</head>
<body>
${headerBar_('manager', 'Jobs')}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Jobs overview</div>
  <table class="table">
    <tr>
      <th>Job ID</th>
      <th>Customer</th>
      <th>Site</th>
      <th>Status</th>
      <th>Ready</th>
      <th>Invoice</th>
      <th>Invoice status</th>
      <th>Before</th>
      <th>After</th>
      <th>Actions</th>
    </tr>
    ${rowsHtml}
  </table>
</div>


${bottomNav_('manager', {})}
</body>
</html>`;
}

function buildJobEditPage_(jobId, role) {
  role = role || 'worker';
  const found = findJobRow(jobId);
  if (!found) return buildMessagePage_('Error: Job not found: ' + jobId);

  const job = found.row;
  const desc = job[JOB_COLS.description - 1] || '';
  const addr = job[JOB_COLS.site_address - 1] || '';
  const cust = job[JOB_COLS.customer_name - 1] || '';
  const email = job[JOB_COLS.customer_email - 1] || '';
  const phone = job[JOB_COLS.customer_phone - 1] || '';
  const status = job[JOB_COLS.status - 1] || 'New';
  const priority = job[JOB_COLS.priority - 1] || 'Normal';

  let calloutVal = job[JOB_COLS.callout_type - 1] || '';
  if (!calloutVal) {
    calloutVal = (priority === 'Urgent' || priority === 'Emergency') ? 'Within hours call-out' : 'No call-out';
  }

  const readyCurrent = (job[JOB_COLS.ready_for_invoice - 1] || '') === 'Yes';
  const beforeIds = (job[JOB_COLS.before_photo_file_id - 1] || '').toString().split('\n').filter(String);
  const afterIds  = (job[JOB_COLS.after_photo_file_id  - 1] || '').toString().split('\n').filter(String);
  const beforeLinksRaw = (job[JOB_COLS.before_photo_link - 1] || '').toString();
  const afterLinksRaw  = (job[JOB_COLS.after_photo_link  - 1] || '').toString();

  const backUrl = role === 'manager' ? buildUrl_('manager', {}) : buildUrl_('worker', {});
  const returnToThis = role === 'manager' ? buildUrl_('managerJob', { jobId: jobId }) : buildUrl_('worker', { jobId: jobId });

  const beforeStrip = renderThumbStrip_(beforeIds, returnToThis, 'thumb-lg');
  const afterStrip  = renderThumbStrip_(afterIds, returnToThis, 'thumb-lg');

  // Pre-fill numeric fields if present
  const hoursVal = job[JOB_COLS.hours_on_site - 1];
  const workersVal = job[JOB_COLS.workers_count - 1] || 1;
  const extrasDesc = job[JOB_COLS.extras_description - 1] || '';
  const extrasAmt = job[JOB_COLS.extras_amount - 1];

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — ${role === 'manager' ? 'Manager' : 'Worker'} Job</title>
${sharedStyles_(role)}
</head>
<body>
${headerBar_(role, 'Job ' + jobId)}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Job details</div>
  <div class="small"><b>Job ID:</b> ${escapeHtml_(jobId)}</div>
  <div class="small"><b>Priority:</b> ${escapeHtml_(priority)}</div>
  <div class="small"><b>Customer:</b> ${escapeHtml_(cust)} (${escapeHtml_(email)}${phone ? ' | ' + escapeHtml_(phone) : ''})</div>
  <div class="small"><b>Address:</b> ${escapeHtml_(addr)}</div>
  <div style="margin-top:8px;"><b>Description:</b><br>${escapeHtml_(desc)}</div>

  ${(beforeIds.length || beforeLinksRaw) ? `
    <hr>
    <div class="small" style="font-weight:900;">Before photos</div>
    ${beforeStrip || ''}
    ${beforeLinksRaw ? `<div class="small" style="margin-top:6px;">${escapeHtml_(beforeLinksRaw).replace(/\\n/g,'<br>')}</div>` : ''}
  ` : ''}

  ${(afterIds.length || afterLinksRaw) ? `
    <hr>
    <div class="small" style="font-weight:900;">After photos</div>
    ${afterStrip || ''}
    ${afterLinksRaw ? `<div class="small" style="margin-top:6px;">${escapeHtml_(afterLinksRaw).replace(/\\n/g,'<br>')}</div>` : ''}
  ` : ''}
</div>

<div class="card">
  <div class="section-title"><span class="dot"></span>Update</div>
  <form id="jobUpdateForm" method="post" action="${getAppUrl_()}" target="_self">
    <input type="hidden" name="action" value="workerUpdateJob">
    <input type="hidden" name="role" value="${escapeHtml_(role)}">
    <input type="hidden" name="jobId" value="${escapeHtml_(jobId)}">
    <input type="hidden" name="afterPhotoDataJson" id="afterPhotoDataJson">
    <input type="hidden" name="readyForInvoice" id="readyForInvoice" value="${readyCurrent ? 'Yes' : ''}">

    <label>Status</label>
    <select name="status">
      <option value="In progress" ${status === 'In progress' ? 'selected' : ''}>In progress</option>
      <option value="Complete" ${status === 'Complete' ? 'selected' : ''}>Complete</option>
      <option value="On hold" ${status === 'On hold' ? 'selected' : ''}>On hold</option>
    </select>

    <label>Call-out</label>
    <select name="calloutType">
      <option value="No call-out" ${calloutVal === 'No call-out' ? 'selected' : ''}>No call-out</option>
      <option value="Within hours call-out" ${calloutVal === 'Within hours call-out' ? 'selected' : ''}>Within hours call-out</option>
      <option value="Out of hours call-out" ${calloutVal === 'Out of hours call-out' ? 'selected' : ''}>Out of hours call-out</option>
    </select>

    <label>Hours on site (per worker)</label>
    <input type="number" step="0.25" min="0" name="hoursOnSite" placeholder="e.g. 4" value="${hoursVal !== '' && hoursVal !== null && hoursVal !== undefined ? escapeHtml_(hoursVal) : ''}">

    <label>Number of workers on site</label>
    <input type="number" step="1" min="1" name="workersCount" placeholder="e.g. 2" value="${escapeHtml_(workersVal)}">
    <div class="small">Example: if two of you are there for 4 hours each, enter 4 hours and 2 workers (8 hours total).</div>

    <label>Materials / extras description</label>
    <textarea name="extrasDescription" rows="2" placeholder="e.g. fittings, silicone, screws, consumables">${escapeHtml_(extrasDesc)}</textarea>

    <label>Materials / extras total (£)</label>
    <input type="number" step="0.01" min="0" name="extrasAmount" value="${extrasAmt !== '' && extrasAmt !== null && extrasAmt !== undefined ? escapeHtml_(extrasAmt) : ''}">

    <label>Take or upload completion photos (optional)</label>
    <input type="file" id="afterPhotoFile" accept="image/*" capture="environment" multiple>

    <label>Or paste completion photo links (optional)</label>
    <textarea name="afterPhotoLink" rows="2" placeholder="Google Photos / Drive etc."></textarea>

    <button type="button" id="readyToggle" class="toggle ${readyCurrent ? 'on' : ''}">
      ${readyCurrent ? '✓ Job is complete and ready for billing' : 'Job is complete and ready for billing'}
    </button>
    <div class="small" style="margin-top:6px;">Tap the button above to turn billing-ready on/off.</div>

    <button class="btn" type="submit" id="jobSubmitBtn">Save update</button>
  </form>
</div>

<script>
(function(){
  var btn = document.getElementById('readyToggle');
  var hidden = document.getElementById('readyForInvoice');
  function sync(){
    var on = (hidden.value === 'Yes');
    btn.classList.toggle('on', on);
    btn.textContent = on ? '✓ Job is complete and ready for billing' : 'Job is complete and ready for billing';
  }
  if (btn && hidden){
    btn.addEventListener('click', function(){
      hidden.value = (hidden.value === 'Yes') ? '' : 'Yes';
      sync();
    });
    sync();
  }
})();
</script>

${clientImageCompressScript_('jobUpdateForm','afterPhotoFile','afterPhotoDataJson','jobSubmitBtn')}

</body>
</html>`;
}

function buildInvoiceViewPage_(jobId, email) {
  if (!jobId) return buildMessagePage_('Error: Missing jobId for invoice view.');
  const found = findJobRow(jobId);
  if (!found) return buildMessagePage_('Error: Job not found: ' + jobId);

  const row = found.row;
  const invId = row[JOB_COLS.invoice_id - 1] || '';
  const pdfUrl = row[JOB_COLS.invoice_pdf_url - 1] || '';
  const custEmail = email || row[JOB_COLS.customer_email - 1] || '';

  const backUrl = buildUrl_('customer', { email: custEmail });

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Invoice</title>
${sharedStyles_('customer')}
${appFeelScript_()}
</head>
<body>
${headerBar_('customer', invId ? ('Invoice ' + invId) : 'Invoice')}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Invoice preview</div>
  ${pdfUrl ? `
    <iframe src="${escapeHtml_(pdfUrl)}" style="width:100%;height:70vh;border:1px solid rgba(0,0,0,0.15);border-radius:12px;"></iframe>
  ` : `<div class="small">No PDF preview available for this invoice yet.</div>`}
</div>

</body>
</html>`;
}

function buildPhotoViewerPage_(fileId, returnToUrl) {
  if (!fileId) return buildMessagePage_('Error: Missing photo fileId.');
  const backUrl = safeReturnUrl_(returnToUrl);
  const imgSrc = driveImageViewUrl(fileId);
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Photo</title>
${sharedStyles_('home')}
${appFeelScript_()}
</head>
<body>
${headerBar_('home', 'Photo')}
<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>

<div class="card">
  <div class="section-title"><span class="dot"></span>Photo</div>
  <img src="${escapeHtml_(imgSrc)}" alt="" style="width:100%;height:auto;border-radius:14px;border:1px solid rgba(0,0,0,0.15);">
</div>

</body>
</html>`;
}

function buildMessagePage_(messageHtml, opts) {
  opts = opts || {};
  const isError = /Error:/i.test(messageHtml);
  const heading = isError ? 'Something went wrong' : 'Done';
  const subheading = isError ? 'Please check and try again.' : 'Saved.';
  const iconChar = isError ? '!' : '✓';
  const backUrl = opts.backUrl || buildUrl_('home', {});
  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Portal</title>
${sharedStyles_('home')}
${appFeelScript_()}
</head>
<body>
${headerBar_('home', '')}

<div class="card" style="text-align:center;">
  <div style="width:64px;height:64px;border-radius:999px;margin:0 auto 10px;display:flex;align-items:center;justify-content:center;
              color:#fff;font-size:34px;font-weight:900;background:${isError ? '#e53935' : '#43a047'};">
    ${iconChar}
  </div>
  <h1 style="margin:6px 0 2px;">${heading}</h1>
  <div class="small" style="margin-bottom:10px;">${subheading}</div>
  <div style="font-size:15px;">${messageHtml}</div>
</div>

<a class="btn btn-secondary" href="${backUrl}" target="_self">Back</a>
</body>
</html>`;
}

function buildManagerInvoiceConfirmationPage_(jobRow, result) {
  const jobId = jobRow[JOB_COLS.job_id - 1];
  const custName = jobRow[JOB_COLS.customer_name - 1] || '';
  const custEmail = jobRow[JOB_COLS.customer_email - 1] || '';
  const siteAddr = jobRow[JOB_COLS.site_address - 1] || '';
  const desc = jobRow[JOB_COLS.description - 1] || '';

  const invoiceId = result.invoiceId;
  const pdfUrl = result.pdfUrl || '';
  const pdfFileId = result.pdfFileId || '';
  const net = result.net || 0;
  const vat = result.vat || 0;
  const total = result.total || 0;

  const backManager = buildUrl_('manager', {});
  const backJobEdit = buildUrl_('managerJob', { jobId: jobId });

  return `
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml_(COMPANY.name)} — Invoice Created</title>
${sharedStyles_('manager')}
${appFeelScript_()}
</head>
<body>
${headerBar_('manager', 'Invoice ' + invoiceId)}

<div class="card">
  <div class="section-title"><span class="dot"></span>Summary</div>
  <div class="small"><b>Job:</b> ${escapeHtml_(jobId)} — ${escapeHtml_(siteAddr)}</div>
  <div class="small"><b>Customer:</b> ${escapeHtml_(custName)} (${escapeHtml_(custEmail)})</div>
  <div class="small"><b>Description:</b> ${escapeHtml_(desc)}</div>
  <hr>
  <div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;">
    <div><span class="small">Net</span><div style="font-weight:900;">${CURRENCY}${net.toFixed(2)}</div></div>
    <div><span class="small">VAT</span><div style="font-weight:900;">${CURRENCY}${vat.toFixed(2)}</div></div>
    <div><span class="small">Total</span><div style="font-weight:900;color:var(--accent);">${CURRENCY}${total.toFixed(2)}</div></div>
  </div>
</div>

<div class="card">
  <div class="section-title"><span class="dot"></span>Invoice preview</div>
  ${pdfUrl ? `
    <iframe src="${escapeHtml_(pdfUrl)}" style="width:100%;height:70vh;border:1px solid rgba(0,0,0,0.15);border-radius:12px;"></iframe>
  ` : `<div class="small">No PDF preview available.</div>`}
</div>

<div class="card">
  <div class="section-title"><span class="dot"></span>Next steps</div>
  <div class="small">If anything is wrong, go back to the job, correct hours/materials/call-out, then regenerate the invoice.</div>
  <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:10px;">
    <a class="btn btn-ghost" href="${backJobEdit}" target="_self">Back to job (edit)</a>
    <a class="btn btn-secondary" href="${backManager}" target="_self">Back to manager</a>
  </div>
  <hr>
  <form method="post" action="${getAppUrl_()}" target="_self">
    <input type="hidden" name="action" value="managerCreateDraftInvoice">
    <input type="hidden" name="jobId" value="${escapeHtml_(jobId)}">
    <input type="hidden" name="invoiceId" value="${escapeHtml_(invoiceId)}">
    <input type="hidden" name="pdfFileId" value="${escapeHtml_(pdfFileId)}">

    <label>Customer email</label>
    <input type="email" name="emailTo" value="${escapeHtml_(custEmail)}" required>

    <button class="btn" type="submit">Create Gmail draft (PDF attached)</button>
    <div class="small">A draft is created in Gmail Drafts; you review and send manually.</div>
  </form>
</div>

</body>
</html>`;
}

/************************************************************
 * CLIENT-SIDE IMAGE COMPRESSION (mobile reliability)
 ************************************************************/
function clientImageCompressScript_(formId, fileInputId, hiddenJsonId, submitBtnId) {
  return `
<script>
(function(){
  var form = document.getElementById(${JSON.stringify(formId)});
  var fileInput = document.getElementById(${JSON.stringify(fileInputId)});
  var hidden = document.getElementById(${JSON.stringify(hiddenJsonId)});
  var submitBtn = document.getElementById(${JSON.stringify(submitBtnId)});

  if (!form || !fileInput || !hidden) return;

  function lock(){
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.textContent = 'Uploading...';
    }
  }

  function unlock(){
    if (submitBtn) {
      submitBtn.disabled = false;
    }
  }

  function compressToDataUrl(file, maxDim, quality){
    return new Promise(function(resolve){
      var reader = new FileReader();
      reader.onload = function(e){
        var img = new Image();
        img.onload = function(){
          var w = img.width, h = img.height;
          var scale = Math.min(1, maxDim / Math.max(w, h));
          var nw = Math.round(w * scale);
          var nh = Math.round(h * scale);

          var canvas = document.createElement('canvas');
          canvas.width = nw; canvas.height = nh;
          var ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, nw, nh);

          try {
            var out = canvas.toDataURL('image/jpeg', quality);
            resolve(out);
          } catch (err) {
            resolve(e.target.result); // fallback original
          }
        };
        img.onerror = function(){ resolve(e.target.result); };
        img.src = e.target.result;
      };
      reader.onerror = function(){ resolve(''); };
      reader.readAsDataURL(file);
    });
  }

  form.addEventListener('submit', function(ev){
    try {
      if (!fileInput.files || fileInput.files.length === 0) return;
      ev.preventDefault();
      lock();
      var files = Array.prototype.slice.call(fileInput.files);
      var dataUrls = [];
      var maxDim = 1600;
      var quality = 0.78;

      (async function(){
        for (var i=0; i<files.length; i++){
          var d = await compressToDataUrl(files[i], maxDim, quality);
          if (d) dataUrls.push(d);
        }
        hidden.value = JSON.stringify(dataUrls);
        form.submit();
      })();

    } catch (err) {
      unlock();
      // last resort: submit without photos
      form.submit();
    }
  });

})();
</script>
`;
}



function clientImageCompressScriptMulti_(formId, fileInputIds, hiddenJsonId, submitBtnId, hintId) {
  const hintIdStr = hintId ? JSON.stringify(String(hintId)) : '""';
  return `
<script>
(function(){
  var form = document.getElementById(${JSON.stringify(formId)});
  var hidden = document.getElementById(${JSON.stringify(hiddenJsonId)});
  var submitBtn = document.getElementById(${JSON.stringify(submitBtnId)});
  var hint = ${hintIdStr} ? document.getElementById(${hintIdStr}) : null;

  if (!form || !hidden) return;

  var ids = ${JSON.stringify(fileInputIds)};
  var inputs = ids.map(function(id){ return document.getElementById(id); }).filter(Boolean);

  // Wire optional buttons if present
  var btnTake = document.getElementById('btnTakePhoto');
  var btnChoose = document.getElementById('btnChoosePhotos');
  var cam = document.getElementById('beforePhotoCamera');
  var lib = document.getElementById('beforePhotoLibrary');
  if (btnTake && cam) btnTake.addEventListener('click', function(){ cam.click(); });
  if (btnChoose && lib) btnChoose.addEventListener('click', function(){ lib.click(); });

  function toast(msg, ms){
    try { if (window.GCToast) window.GCToast(msg, ms); } catch(e){}
  }

  function lock(label){
    if (submitBtn) {
      submitBtn.disabled = true;
      submitBtn.textContent = label || 'Uploading…';
    }
  }
  function unlock(){
    if (submitBtn) submitBtn.disabled = false;
  }

  function getAllFiles(){
    var files = [];
    inputs.forEach(function(inp){
      if (inp && inp.files && inp.files.length){
        for (var i=0;i<inp.files.length;i++) files.push(inp.files[i]);
      }
    });
    // Deduplicate (name+size+lastModified)
    var seen = {};
    var out = [];
    files.forEach(function(f){
      var k = [f.name,f.size,f.lastModified].join('|');
      if (!seen[k]) { seen[k]=1; out.push(f); }
    });
    return out;
  }

  function updateHint(){
    if (!hint) return;
    var n = getAllFiles().length;
    hint.textContent = n ? (n + ' photo' + (n===1?'':'s') + ' selected') : 'No photos selected.';
  }
  inputs.forEach(function(inp){
    if (inp) inp.addEventListener('change', updateHint);
  });
  updateHint();

  function compressToDataUrl(file, maxDim, quality){
    return new Promise(function(resolve){
      var reader = new FileReader();
      reader.onload = function(e){
        var img = new Image();
        img.onload = function(){
          var w = img.width, h = img.height;
          var scale = Math.min(1, maxDim / Math.max(w, h));
          var nw = Math.max(1, Math.round(w * scale));
          var nh = Math.max(1, Math.round(h * scale));
          var canvas = document.createElement('canvas');
          canvas.width = nw; canvas.height = nh;
          var ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, nw, nh);
          try {
            resolve(canvas.toDataURL('image/jpeg', quality));
          } catch (err) {
            resolve(e.target.result || '');
          }
        };
        img.onerror = function(){ resolve(e.target.result || ''); };
        img.src = e.target.result;
      };
      reader.onerror = function(){ resolve(''); };
      reader.readAsDataURL(file);
    });
  }

  form.addEventListener('submit', function(ev){
    try {
      var files = getAllFiles();
      if (!files.length) return; // no photos: let normal submit happen
      ev.preventDefault();
      lock('Preparing photos…');
      toast('Preparing photos…', 1200);

      var maxDim = 1600;
      var quality = 0.78;

      (async function(){
        var dataUrls = [];
        for (var i=0; i<files.length; i++){
          if (submitBtn) submitBtn.textContent = 'Preparing ' + (i+1) + '/' + files.length + '…';
          var d = await compressToDataUrl(files[i], maxDim, quality);
          if (d) dataUrls.push(d);
        }
        hidden.value = JSON.stringify(dataUrls);
        if (submitBtn) submitBtn.textContent = 'Uploading…';
        toast('Uploading…', 1200);
        form.submit();
      })();

    } catch (err) {
      unlock();
      form.submit();
    }
  }, true);

})();
</script>
`;
}

/************************************************************
 * ACTION HANDLERS
 ************************************************************/
function handleCustomerNewJob(e) {
  const p = e.parameter || {};

  const name    = (p.customerName   || '').toString().trim();
  const email   = (p.customerEmail  || '').toString().trim();
  const phone   = (p.customerPhone  || '').toString().trim();
  const address = (p.siteAddress    || '').toString().trim();
  const desc    = (p.jobDescription || '').toString().trim();
  const priority= (p.priority       || 'Normal').toString().trim();
  const beforeLinkText = (p.beforePhotoLink || '').toString().trim();
  const beforeDataJson = (p.beforePhotoDataJson || '').toString();

  if (!name || !email || !address || !desc) {
    throw new Error('Name, email, address and job description are required.');
  }

  const sh = getJobsSheet();
  const jobId = generateJobId();

  let beforeUrl = beforeLinkText || '';
  let beforeIds = [];

  if (beforeDataJson) {
    let arr;
    try { arr = JSON.parse(beforeDataJson); } catch (err) { arr = []; }
    if (Array.isArray(arr)) {
      arr.forEach((dataUrl, idx) => {
        if (!dataUrl) return;
        const saved = saveImageFromDataUrl(dataUrl, jobId, 'before_' + (idx+1));
        if (saved.url) beforeUrl = beforeUrl ? (beforeUrl + '\n' + saved.url) : saved.url;
        if (saved.id) beforeIds.push(saved.id);
      });
    }
  }

  const row = [];
  row[JOB_COLS.job_id - 1]               = jobId;
  row[JOB_COLS.created - 1]              = new Date();
  row[JOB_COLS.customer_name - 1]        = name;
  row[JOB_COLS.customer_email - 1]       = email;
  row[JOB_COLS.customer_phone - 1]       = phone;
  row[JOB_COLS.site_address - 1]         = address;
  row[JOB_COLS.priority - 1]             = priority;
  row[JOB_COLS.callout_type - 1]         = 'No call-out';
  row[JOB_COLS.status - 1]               = 'New';
  row[JOB_COLS.description - 1]          = desc;
  row[JOB_COLS.before_photo_link - 1]    = beforeUrl;
  row[JOB_COLS.before_photo_file_id - 1] = beforeIds.join('\n');
  row[JOB_COLS.hours_on_site - 1]        = '';
  row[JOB_COLS.workers_count - 1]        = 1;
  row[JOB_COLS.extras_description - 1]   = '';
  row[JOB_COLS.extras_amount - 1]        = '';
  row[JOB_COLS.ready_for_invoice - 1]    = 'No';
  row[JOB_COLS.invoice_id - 1]           = '';
  row[JOB_COLS.invoice_status - 1]       = '';
  row[JOB_COLS.invoice_pdf_url - 1]      = '';

  const maxCol = Object.keys(JOB_COLS).length;
  for (let i = 0; i < maxCol; i++) {
    if (row[i] === undefined) row[i] = '';
  }

  sh.appendRow(row);
  Logger.log('New job created %s, beforeIds=%s', jobId, beforeIds.join(','));
  return jobId;
}


function handleCustomerUpdateJob(e) {
  const p = (e && e.parameter) || {};
  const jobId = (p.jobId || '').toString().trim();
  const email = (p.customerEmail || p.email || '').toString().trim();
  const desc = (p.jobDescription || '').toString().trim();
  const priority = (p.priority || 'Normal').toString().trim();

  if (!jobId) throw new Error('Missing jobId.');
  if (!email) throw new Error('Missing customer email.');
  if (!desc) throw new Error('Missing job description.');

  const found = assertCustomerOwnsJob_(email, jobId);
  const sh = getJobsSheet();
  const row = found.row;

  row[JOB_COLS.description - 1] = desc;
  row[JOB_COLS.priority - 1] = priority;

  // Optional: bump status if needed (do not mark complete etc.)
  // Keep status as-is.

  sh.getRange(found.rowIndex, 1, 1, sh.getLastColumn()).setValues([row]);
  return jobId;
}

function handleCustomerAddPhotos(e) {
  const p = (e && e.parameter) || {};
  const jobId = (p.jobId || '').toString().trim();
  const email = (p.customerEmail || p.email || '').toString().trim();
  const dataJson = (p.beforePhotoDataJson || '').toString();

  if (!jobId) throw new Error('Missing jobId.');
  if (!email) throw new Error('Missing customer email.');

  const found = assertCustomerOwnsJob_(email, jobId);
  const sh = getJobsSheet();
  const row = found.row;

  // Append new photos to BEFORE slot (customer-supplied evidence). Store IDs; keep strip limited to 5 in UI.
  let existingIds = (row[JOB_COLS.before_photo_file_id - 1] || '').toString().split('\n').filter(String);
  let existingLinks = (row[JOB_COLS.before_photo_link - 1] || '').toString().split('\n').filter(String);

  let dataUrls = [];
  try { dataUrls = dataJson ? JSON.parse(dataJson) : []; } catch (err) { dataUrls = []; }

  if (dataUrls && dataUrls.length) {
    dataUrls.forEach((d, idx) => {
      const saved = saveImageFromDataUrl(d, jobId, 'before_more');
      if (saved && saved.id) existingIds.push(saved.id);
      if (saved && saved.url) existingLinks.push(saved.url);
    });
  }

  // Persist (store all; UI shows first 5)
  row[JOB_COLS.before_photo_file_id - 1] = existingIds.join('\n');
  row[JOB_COLS.before_photo_link - 1] = existingLinks.join('\n');

  sh.getRange(found.rowIndex, 1, 1, sh.getLastColumn()).setValues([row]);
  return jobId;
}

function handleWorkerUpdateJob(e) {
  const p = e.parameter || {};

  const jobId  = (p.jobId || '').toString().trim();
  if (!jobId) throw new Error('Job ID is required.');

  const status = (p.status || 'In progress').toString().trim();
  const hours  = p.hoursOnSite ? parseFloat(p.hoursOnSite) : NaN;
  const workersCount = p.workersCount ? parseInt(p.workersCount, 10) : NaN;
  const extrasDesc = (p.extrasDescription || '').toString().trim();
  const extrasAmt  = p.extrasAmount ? parseFloat(p.extrasAmount) : NaN;
  const afterLinkText = (p.afterPhotoLink || '').toString().trim();
  const afterDataJson = (p.afterPhotoDataJson || '').toString();
  const readyForInvoice = (p.readyForInvoice === 'Yes') ? 'Yes' : 'No';
  let calloutType = (p.calloutType || '').toString().trim();

  if (!isNaN(hours) && hours < 0) throw new Error('Hours on site cannot be negative.');
  if (!isNaN(extrasAmt) && extrasAmt < 0) throw new Error('Extras amount cannot be negative.');

  const found = findJobRow(jobId);
  if (!found) throw new Error('Job not found: ' + jobId);

  const sh = getJobsSheet();
  const rowIndex = found.rowIndex;
  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

  const currentPriority = row[JOB_COLS.priority - 1] || 'Normal';

  if (!calloutType) {
    calloutType = row[JOB_COLS.callout_type - 1] || '';
    if (!calloutType) {
      calloutType = (currentPriority === 'Urgent' || currentPriority === 'Emergency')
        ? 'Within hours call-out'
        : 'No call-out';
    }
  }

  row[JOB_COLS.status - 1]       = status;
  row[JOB_COLS.callout_type - 1] = calloutType;

  if (!isNaN(hours)) row[JOB_COLS.hours_on_site - 1] = hours;
  if (!isNaN(workersCount) && workersCount > 0) row[JOB_COLS.workers_count - 1] = workersCount;
  row[JOB_COLS.extras_description - 1] = extrasDesc;
  if (!isNaN(extrasAmt)) row[JOB_COLS.extras_amount - 1] = extrasAmt;

  let afterUrlCombined = row[JOB_COLS.after_photo_link - 1] || '';
  let afterFileIds = (row[JOB_COLS.after_photo_file_id - 1] || '').toString().split('\n').filter(String);

  if (afterDataJson) {
    let arr;
    try { arr = JSON.parse(afterDataJson); } catch (err) { arr = []; }
    if (Array.isArray(arr)) {
      arr.forEach((dataUrl, idx) => {
        if (!dataUrl) return;
        const saved = saveImageFromDataUrl(dataUrl, jobId, 'after_' + (afterFileIds.length + idx + 1));
        if (saved.url) afterUrlCombined = afterUrlCombined ? (afterUrlCombined + '\n' + saved.url) : saved.url;
        if (saved.id) afterFileIds.push(saved.id);
      });
    }
  }

  if (afterLinkText) {
    afterUrlCombined = afterUrlCombined ? (afterUrlCombined + '\n' + afterLinkText) : afterLinkText;
  }

  row[JOB_COLS.after_photo_link - 1]    = afterUrlCombined;
  row[JOB_COLS.after_photo_file_id - 1] = afterFileIds.join('\n');
  row[JOB_COLS.ready_for_invoice - 1]   = readyForInvoice;

  sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).setValues([row]);
  Logger.log('Job updated %s, afterIds=%s', jobId, afterFileIds.join(','));
}

/************************************************************
 * INVOICE GENERATION
 ************************************************************/
function createInvoiceForJob(jobId) {
  if (!jobId) throw new Error('Job ID is required.');

  const found = findJobRow(jobId);
  if (!found) throw new Error('Job not found: ' + jobId);

  const sh = getJobsSheet();
  const rowIndex = found.rowIndex;
  const row = found.row;

  let invoiceId = row[JOB_COLS.invoice_id - 1];
  if (!invoiceId) {
    invoiceId = generateInvoiceId();
    row[JOB_COLS.invoice_id - 1] = invoiceId;
    row[JOB_COLS.invoice_status - 1] = 'Draft';
  }

  // Remove existing lines for this invoice ID (simple rebuild)
  const invSh = getInvoiceLinesSheet();
  const lastInv = invSh.getLastRow();
  if (lastInv >= 2) {
    const rng = invSh.getRange(2, 1, lastInv - 1, invSh.getLastColumn());
    const data = rng.getValues();
    const keep = data.filter(r => r[INV_COLS.invoice_id - 1] !== invoiceId);
    rng.clearContent();
    if (keep.length) {
      invSh.getRange(2, 1, keep.length, invSh.getLastColumn()).setValues(keep);
    }
  }

  const hours = parseFloat(row[JOB_COLS.hours_on_site - 1] || '0') || 0;
  const workersCount = parseInt(row[JOB_COLS.workers_count - 1] || '1', 10) || 1;
  const extrasAmt = parseFloat(row[JOB_COLS.extras_amount - 1] || '0') || 0;
  const extrasDesc = row[JOB_COLS.extras_description - 1] || '';
  const description = row[JOB_COLS.description - 1] || '';
  const priority = row[JOB_COLS.priority - 1] || 'Normal';
  const calloutType = row[JOB_COLS.callout_type - 1] || 'No call-out';

  const lines = [];
  let lineNo = 1;
  let net = 0;

  const calloutFee = getCalloutFee(priority, calloutType);
  if (calloutFee > 0) {
    const calloutDesc =
      'Call-out – ' +
      (calloutType === 'Out of hours call-out' ? 'Out of hours' : 'Within hours') +
      ' (' + priority + ')';
    lines.push({
      invoice_id: invoiceId, job_id: jobId, line_no: lineNo++,
      item_code: 'CALLOUT', description: calloutDesc, qty: 1,
      unit_price: calloutFee, line_total: calloutFee
    });
    net += calloutFee;
  }

  const billedHours = hours * workersCount;
  const labourTotal = billedHours * HOURLY_RATE;
  if (labourTotal > 0) {
    const labourDesc = 'Labour – ' + workersCount + ' worker' + (workersCount !== 1 ? 's' : '') +
      ' × ' + hours.toFixed(2) + ' hours (' + billedHours.toFixed(2) + ' total) – ' + description;
    lines.push({
      invoice_id: invoiceId, job_id: jobId, line_no: lineNo++,
      item_code: 'LABOUR', description: labourDesc, qty: 1,
      unit_price: labourTotal, line_total: labourTotal
    });
    net += labourTotal;
  }

  if (extrasAmt > 0) {
    lines.push({
      invoice_id: invoiceId, job_id: jobId, line_no: lineNo++,
      item_code: 'EXTRA', description: extrasDesc || 'Materials & extras', qty: 1,
      unit_price: extrasAmt, line_total: extrasAmt
    });
    net += extrasAmt;
  }

  if (!lines.length) {
    throw new Error('No billable items – set hours or extras first before generating an invoice.');
  }

  const vat = net * VAT_RATE;
  if (vat > 0) {
    lines.push({
      invoice_id: invoiceId, job_id: jobId, line_no: lineNo++,
      item_code: 'VAT', description: 'VAT @ ' + (VAT_RATE * 100).toFixed(0) + '%', qty: 1,
      unit_price: vat, line_total: vat
    });
  }
  const total = net + vat;

  const values = lines.map(l => [
    l.invoice_id, l.job_id, l.line_no, l.item_code,
    l.description, l.qty, l.unit_price, l.line_total
  ]);
  invSh.getRange(invSh.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);

  const pdfInfo = createInvoiceDocAndPdfForJob_(row, lines, net, vat, total);

  row[JOB_COLS.invoice_pdf_url - 1] = pdfInfo.previewUrl;
  sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).setValues([row]);

  Logger.log('Invoice created %s for job %s, total=%s', invoiceId, jobId, total);
  return {
    invoiceId: invoiceId,
    pdfUrl: pdfInfo.previewUrl,
    pdfFileId: pdfInfo.pdfFileId,
    lines: lines,
    net: net,
    vat: vat,
    total: total
  };
}

/************************************************************
 * INVOICE DOC/PDF CREATION (more robust)
 ************************************************************/
function withRetries_(fn, attempts, sleepMs) {
  attempts = attempts || 3;
  sleepMs = sleepMs || 300;
  let lastErr;
  for (let i = 0; i < attempts; i++) {
    try { return fn(); } catch (err) { lastErr = err; Utilities.sleep(sleepMs * (i + 1)); }
  }
  throw lastErr;
}

function createInvoiceDocAndPdfForJob_(jobRow, lines, net, vat, total) {
  const jobId    = jobRow[JOB_COLS.job_id - 1];
  const custName = jobRow[JOB_COLS.customer_name - 1] || '';
  const address  = jobRow[JOB_COLS.site_address - 1] || '';
  const descFull = jobRow[JOB_COLS.description - 1] || '';
  const invId    = jobRow[JOB_COLS.invoice_id - 1] || '';

  const beforeIds = (jobRow[JOB_COLS.before_photo_file_id - 1] || '').toString().split('\n').filter(String);
  const afterIds  = (jobRow[JOB_COLS.after_photo_file_id  - 1] || '').toString().split('\n').filter(String);

  // Keep within one page: trim description a bit.
  let desc = descFull;
  const maxDescLen = 180;
  if (desc.length > maxDescLen) desc = desc.substring(0, maxDescLen - 3) + '...';

  const ssFile = DriveApp.getFileById(getSpreadsheetId_());
  const parents = ssFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  let invoicesFolder;
  const it = parentFolder.getFoldersByName('Invoices');
  invoicesFolder = it.hasNext() ? it.next() : parentFolder.createFolder('Invoices');
  try { invoicesFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}

  // Create doc then move into folder
  const doc = DocumentApp.create('Invoice ' + invId + ' – ' + custName);
  const docId = doc.getId();
  const docFile = DriveApp.getFileById(docId);
  invoicesFolder.addFile(docFile);

  try { docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}

  const body = doc.getBody();
  body.clear();

  // Safe defaults
  try {
    body.setAttributes({
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial',
      [DocumentApp.Attribute.FONT_SIZE]: 9
    });
  } catch (err) {}

  // Title (company name prominent)
  const title1 = body.appendParagraph(COMPANY.name);
  title1.setBold(true)
        .setFontSize(18)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingBefore(0)
        .setSpacingAfter(2);

  const title2 = body.appendParagraph('INVOICE');
  title2.setBold(true)
        .setForegroundColor('#1976d2')
        .setFontSize(14)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setSpacingBefore(0)
        .setSpacingAfter(6);

  // Header table
  const headerTable = body.appendTable();
  headerTable.setBorderWidth(0);
  const headerRow = headerTable.appendTableRow();
  const leftCell  = headerRow.appendTableCell();
  const rightCell = headerRow.appendTableCell();

  COMPANY.addressLines.forEach(line => leftCell.appendParagraph(line).setSpacingAfter(0));
  if (COMPANY.phone) leftCell.appendParagraph('Tel: ' + COMPANY.phone).setSpacingAfter(0);
  if (COMPANY.email) leftCell.appendParagraph('Email: ' + COMPANY.email).setSpacingAfter(0);
  if (COMPANY.website) leftCell.appendParagraph('Web: ' + COMPANY.website).setSpacingAfter(0);

  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  rightCell.appendParagraph('Invoice #: ' + invId).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0);
  rightCell.appendParagraph('Date: ' + todayStr).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0);
  rightCell.appendParagraph('Job ID: ' + jobId).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0);

  body.appendParagraph('').setSpacingAfter(6).setSpacingBefore(0);

  // Bill-to + Job section with before thumbnails
  const detailsTable = body.appendTable();
  detailsTable.setBorderWidth(0);
  const dRow = detailsTable.appendTableRow();
  const billCell = dRow.appendTableCell();
  const jobCell  = dRow.appendTableCell();

  billCell.appendParagraph('Bill to').setBold(true).setForegroundColor('#1976d2').setSpacingAfter(1);
  billCell.appendParagraph(custName || '').setSpacingAfter(0);
  billCell.appendParagraph(address || '').setSpacingAfter(0);

  jobCell.appendParagraph('Job').setBold(true).setForegroundColor('#1976d2').setSpacingAfter(1);
  jobCell.appendParagraph(desc || '').setSpacingAfter(2);

  if (beforeIds.length) {
    const thumbSize = 44;
    const imgPara = jobCell.appendParagraph('');
    imgPara.setSpacingBefore(0).setSpacingAfter(0);
    beforeIds.slice(0, 5).forEach(id => {
      try {
        const file = DriveApp.getFileById(id);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
        const img = imgPara.appendInlineImage(file.getBlob());
        img.setWidth(thumbSize);
        img.setHeight(thumbSize);
        img.setLinkUrl(file.getUrl());
        imgPara.appendText(' ');
      } catch (err) {
        Logger.log('Error adding before photo %s: %s', id, err);
      }
    });
  }

  body.appendParagraph('').setSpacingAfter(6).setSpacingBefore(0);

  // Line items table
  const linesTable = body.appendTable();
  linesTable.setBorderWidth(0.5);
  const headerRow2 = linesTable.appendTableRow();
  ['Description','Qty','Unit','Total'].forEach(text => {
    const cell = headerRow2.appendTableCell(text);
    cell.setBold(true).setBackgroundColor('#f0f4ff');
  });

  lines.forEach(l => {
    const r = linesTable.appendTableRow();
    r.appendTableCell(l.description);
    r.appendTableCell(String(l.qty));
    r.appendTableCell(CURRENCY + Number(l.unit_price).toFixed(2));
    r.appendTableCell(CURRENCY + Number(l.line_total).toFixed(2));
  });

  body.appendParagraph('').setSpacingAfter(2).setSpacingBefore(0);

  // After thumbs under table
  if (afterIds.length) {
    const thumbSize = 44;
    const table = body.appendTable();
    table.setBorderWidth(0);
    const row = table.appendTableRow();
    afterIds.slice(0, 5).forEach(id => {
      const cell = row.appendTableCell();
      try {
        const file = DriveApp.getFileById(id);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
        const para2 = cell.appendParagraph('');
        para2.setSpacingBefore(0).setSpacingAfter(0);
        const img = para2.appendInlineImage(file.getBlob());
        img.setWidth(thumbSize);
        img.setHeight(thumbSize);
        img.setLinkUrl(file.getUrl());
      } catch (err) {
        Logger.log('Error adding after photo %s: %s', id, err);
      }
    });
    body.appendParagraph('').setSpacingAfter(2).setSpacingBefore(0);
  }

  // Totals
  const totalsTable = body.appendTable();
  totalsTable.setBorderWidth(0);
  const tRow = totalsTable.appendTableRow();
  tRow.appendTableCell('');
  const tRight = tRow.appendTableCell();

  tRight.appendParagraph('Net: ' + CURRENCY + net.toFixed(2)).setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0);
  tRight.appendParagraph('VAT @ ' + (VAT_RATE * 100).toFixed(0) + '%: ' + CURRENCY + vat.toFixed(2))
        .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0);
  tRight.appendParagraph('Total: ' + CURRENCY + total.toFixed(2))
        .setBold(true).setForegroundColor('#1976d2')
        .setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setSpacingAfter(0);

  // Payment
  const payTitle = body.appendParagraph('Payment');
  payTitle.setBold(true).setForegroundColor('#1976d2').setSpacingBefore(4).setSpacingAfter(0);

  body.appendParagraph('Bank: ' + COMPANY.bankName).setSpacingAfter(0);
  body.appendParagraph('Sort: ' + COMPANY.bankSortCode + '   Acc: ' + COMPANY.bankAccount).setSpacingAfter(0);
  body.appendParagraph('Please pay within ' + COMPANY.paymentTermsDays + ' days using invoice ' + invId + ' as the reference.').setSpacingAfter(0);

  withRetries_(function(){ doc.saveAndClose(); }, 3, 400);

  const pdfFile = withRetries_(function(){
    const pdfBlob = docFile.getAs(MimeType.PDF).setName('Invoice ' + invId + ' – ' + custName + '.pdf');
    const f = invoicesFolder.createFile(pdfBlob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
    return f;
  }, 3, 500);

  const previewUrl = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/preview';
  return { pdfFileId: pdfFile.getId(), previewUrl: previewUrl, docId: docId };
}
