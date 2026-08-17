/**
 * IDD BIM CPG Namecard — Google Apps Script Backend v1.0
 * 
 * Target Google Sheet: 13XsLUK792UkdDuXYHSdjIGJPX7AazQt6hROtTa0LRCI
 * Target Google Drive Folder: 1Sy6Lb509_seEfP-EQd_1Wh_1lipWIxG3 (IDD BIM CPG NAMECARD)
 */

const CONFIG = {
  SPREADSHEET_ID: '13XsLUK792UkdDuXYHSdjIGJPX7AazQt6hROtTa0LRCI',
  DRIVE_FOLDER_ID: '1Sy6Lb509_seEfP-EQd_1Wh_1lipWIxG3',
  TAB_CONTACTS: 'Namecard_CRM_Contacts',
  TAB_INBOX: 'Cards_Inbox',
  TAB_REVIEW: 'Review_Queue',
  DEFAULT_SECRET_KEY: 'PATH_SECRET'
};

// 27-Column Master Schema Definition
const CONTACT_HEADERS = [
  'Contact_ID',
  'Họ tên',
  'Tên chuẩn hóa',
  'Xưng hô',
  'Chức danh (VI)',
  'Job Title (EN)',
  'Cấp bậc',
  'Chức năng / Phòng ban',
  'Công ty / Tổ chức',
  'Tên ngắn',
  'Ngành cấp 1',
  'Ngành cấp 2',
  'Quốc gia',
  'Thành phố',
  'Địa chỉ',
  'Mobile',
  'Điện thoại bàn',
  'Fax',
  'Email chính',
  'Email phụ',
  'Website',
  'Vai trò CRM',
  'Ngôn ngữ',
  'Số card trùng',
  'Nguồn card',
  'Độ tin cậy',
  'Ghi chú kiểm tra'
];

const INBOX_HEADERS = [
  'Card_ID',
  'Received_At',
  'Contact_ID',
  'Status',
  'Source_File_Name',
  'Uploader',
  'OCR_Confidence',
  'Drive_File_ID',
  'Drive_View_URL',
  'Drive_Direct_URL',
  'Raw_JSON'
];

const REVIEW_HEADERS = [
  'Review_ID',
  'Card_ID',
  'Candidate_Contact_ID',
  'Candidate_Name',
  'Candidate_Company',
  'Incoming_Name',
  'Incoming_Company',
  'Incoming_Mobile',
  'Incoming_Email',
  'Duplicate_Score',
  'Reason',
  'Created_At',
  'Status'
];

/**
 * Run setupSystem() once to initialize sheet headers, tabs, and security secret.
 */
function setupSystem() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // 1. Check or create Namecard_CRM_Contacts
  let contactSheet = ss.getSheetByName(CONFIG.TAB_CONTACTS);
  if (!contactSheet) {
    contactSheet = ss.insertSheet(CONFIG.TAB_CONTACTS, 0);
    contactSheet.appendRow(CONTACT_HEADERS);
    contactSheet.getRange(1, 1, 1, CONTACT_HEADERS.length).setFontWeight('bold').setBackground('#f4f6f8');
    contactSheet.setFrozenRows(1);
  }
  
  // 2. Check or create Cards_Inbox
  let inboxSheet = ss.getSheetByName(CONFIG.TAB_INBOX);
  if (!inboxSheet) {
    inboxSheet = ss.insertSheet(CONFIG.TAB_INBOX);
    inboxSheet.appendRow(INBOX_HEADERS);
    inboxSheet.getRange(1, 1, 1, INBOX_HEADERS.length).setFontWeight('bold').setBackground('#f4f6f8');
    inboxSheet.setFrozenRows(1);
  }
  
  // 3. Check or create Review_Queue
  let reviewSheet = ss.getSheetByName(CONFIG.TAB_REVIEW);
  if (!reviewSheet) {
    reviewSheet = ss.insertSheet(CONFIG.TAB_REVIEW);
    reviewSheet.appendRow(REVIEW_HEADERS);
    reviewSheet.getRange(1, 1, 1, REVIEW_HEADERS.length).setFontWeight('bold').setBackground('#f4f6f8');
    reviewSheet.setFrozenRows(1);
  }
  
  // 4. Generate & save path secret if missing
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(CONFIG.DEFAULT_SECRET_KEY);
  if (!secret) {
    secret = 'nmc_' + Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    props.setProperty(CONFIG.DEFAULT_SECRET_KEY, secret);
  }
  
  Logger.log('=== IDD BIM CPG NAMECARD SYSTEM READY ===');
  Logger.log('Spreadsheet ID: ' + CONFIG.SPREADSHEET_ID);
  Logger.log('Drive Folder ID: ' + CONFIG.DRIVE_FOLDER_ID);
  Logger.log('Path Secret: ' + secret);
  Logger.log('Action URL Suffix: /exec/' + secret);
}

/**
 * Handle HTTP POST from Custom GPT Action
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30s lock to prevent race conditions on Contact_ID
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Server busy. Could not obtain lock.' });
  }

  try {
    // 1. Verify Path Secret / Authorization
    const props = PropertiesService.getScriptProperties();
    const expectedSecret = props.getProperty(CONFIG.DEFAULT_SECRET_KEY) || '';
    
    if (expectedSecret && e && e.pathInfo) {
      const path = (e.pathInfo || '').replace(/^\/+|\/+$/g, '');
      if (path !== expectedSecret) {
        return jsonResponse({ ok: false, error: 'Unauthorized: Invalid path secret' }, 401);
      }
    }

    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'Empty payload' }, 400);
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'processNamecard';

    if (action !== 'processNamecard' && action !== 'saveNamecard') {
      return jsonResponse({ ok: false, error: 'Unknown action: ' + action }, 400);
    }

    const result = handleProcessNamecard(payload);
    return jsonResponse(result);
  } catch (err) {
    Logger.log('Error processing POST: ' + err.toString());
    return jsonResponse({ ok: false, error: err.message || err.toString() }, 500);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handle HTTP GET for health-check / data reading
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const contactSheet = ss.getSheetByName(CONFIG.TAB_CONTACTS);
    const contacts = contactSheet ? loadExistingContacts(contactSheet) : [];
    const count = contacts.length;
    
    return jsonResponse({
      ok: true,
      service: 'IDD BIM CPG Namecard Intake API',
      version: '1.0.0',
      total_contacts: count,
      data_version: 'v' + count + '_' + new Date().toISOString().slice(0, 10),
      contacts: contacts
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

/**
 * Core Namecard Processing Logic
 */
function handleProcessNamecard(payload) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const contactSheet = ss.getSheetByName(CONFIG.TAB_CONTACTS);
  const inboxSheet = ss.getSheetByName(CONFIG.TAB_INBOX);
  const reviewSheet = ss.getSheetByName(CONFIG.TAB_REVIEW);

  if (!contactSheet || !inboxSheet) {
    setupSystem();
  }

  const card = payload.card || {};
  const sourceFileName = payload.source_file_name || '';
  const uploader = payload.uploader || 'ChatGPT Business';
  const ocrConfidence = payload.ocr_confidence || 95;

  // 1. Read existing contacts for duplicate matching
  const allContacts = loadExistingContacts(contactSheet);
  
  // 2. Perform duplicate check
  const dupCheck = evaluateDuplicate(card, allContacts);
  
  let finalStatus = dupCheck.status; // NEW | MERGED | REVIEW
  let targetContactId = '';
  let reviewId = '';
  const cardId = generateNextCardId(inboxSheet);

  if (finalStatus === 'NEW') {
    targetContactId = generateNextContactId(contactSheet);
    const newRow = buildContactRow(targetContactId, card, 1, sourceFileName || cardId);
    contactSheet.appendRow(newRow);
  } else if (finalStatus === 'MERGED') {
    targetContactId = dupCheck.existingContact.Contact_ID;
    const rowIndex = dupCheck.existingRowIndex; // 1-based row in sheet
    updateExistingContact(contactSheet, rowIndex, dupCheck.existingContact, card, sourceFileName || cardId);
  } else if (finalStatus === 'REVIEW') {
    targetContactId = dupCheck.existingContact ? dupCheck.existingContact.Contact_ID : '';
    reviewId = generateNextReviewId(reviewSheet);
    
    // Add to review queue
    reviewSheet.appendRow([
      reviewId,
      cardId,
      targetContactId || 'N/A',
      dupCheck.existingContact ? (dupCheck.existingContact['Họ tên'] || '') : '',
      dupCheck.existingContact ? (dupCheck.existingContact['Công ty / Tổ chức'] || '') : '',
      card.full_name || '',
      card.company || '',
      card.mobile || '',
      card.primary_email || '',
      dupCheck.score || 50,
      dupCheck.reason || 'Needs human verification',
      new Date().toISOString(),
      'PENDING'
    ]);
  }

  // 3. Handle image upload to Google Drive if provided
  let imageSaved = false;
  let imageStatus = 'IMAGE_NOT_RECEIVED_BY_ACTION';
  let driveViewUrl = '';
  let driveDirectUrl = '';
  let driveFileId = '';

  const imageBase64 = payload.image_base64;
  const imageMimeType = payload.image_mime_type || 'image/jpeg';
  const imageUrl = payload.image_url;

  if (imageBase64 || imageUrl) {
    try {
      const driveResult = saveCardImageToDrive(
        cardId,
        targetContactId || 'REVIEW',
        card.full_name || 'Unknown',
        card.short_company || card.company || 'Unknown',
        imageBase64,
        imageMimeType,
        imageUrl
      );
      if (driveResult && driveResult.fileId) {
        imageSaved = true;
        imageStatus = 'SAVED';
        driveFileId = driveResult.fileId;
        driveViewUrl = driveResult.viewUrl;
        driveDirectUrl = driveResult.directUrl;
      }
    } catch (imgErr) {
      Logger.log('Error saving image: ' + imgErr.toString());
      imageStatus = 'IMAGE_SAVE_FAILED: ' + imgErr.message;
    }
  }

  // 4. Record entry in Cards_Inbox
  inboxSheet.appendRow([
    cardId,
    new Date().toISOString(),
    targetContactId || (finalStatus === 'REVIEW' ? reviewId : ''),
    finalStatus,
    sourceFileName,
    uploader,
    ocrConfidence,
    driveFileId,
    driveViewUrl,
    driveDirectUrl,
    JSON.stringify(payload)
  ]);

  const totalContacts = Math.max(0, contactSheet.getLastRow() - 1);

  return {
    ok: true,
    status: finalStatus,
    contact_id: targetContactId,
    card_id: cardId,
    review_id: reviewId || undefined,
    duplicate_score: dupCheck.score,
    duplicate_reason: dupCheck.reason,
    image_saved: imageSaved,
    image_status: imageStatus,
    image_url: driveViewUrl || undefined,
    data_version: 'v' + totalContacts + '_' + new Date().toISOString().slice(0, 10),
    message: finalStatus === 'NEW'
      ? 'Created new contact ' + targetContactId
      : finalStatus === 'MERGED'
      ? 'Merged into existing contact ' + targetContactId
      : 'Card flagged for review: ' + (dupCheck.reason || 'Potential duplicate')
  };
}

/**
 * Duplicate Rule Engine:
 * - Email exact match -> MERGED (100)
 * - Mobile exact match -> MERGED (100)
 * - Name exact match + Company partial match -> MERGED / REVIEW based on score
 * - Unclear / ambiguity -> REVIEW
 */
function evaluateDuplicate(card, existingContacts) {
  const normEmail = normalizeString(card.primary_email || card.secondary_email);
  const normPhone = normalizePhoneDigits(card.mobile || card.phone);
  const normName = normalizeVietnamese(card.full_name || card.normalized_name);
  const normComp = normalizeVietnamese(card.company || card.short_company);

  if (!existingContacts || existingContacts.length === 0) {
    return { status: 'NEW', score: 0, reason: 'First contact in database' };
  }

  // Check 1: Primary Email Exact Match
  if (normEmail) {
    for (let i = 0; i < existingContacts.length; i++) {
      const c = existingContacts[i];
      const cEmail1 = normalizeString(c['Email chính']);
      const cEmail2 = normalizeString(c['Email phụ']);
      if ((cEmail1 && cEmail1 === normEmail) || (cEmail2 && cEmail2 === normEmail)) {
        return {
          status: 'MERGED',
          score: 100,
          reason: 'Exact Email match (' + normEmail + ')',
          existingContact: c,
          existingRowIndex: i + 2 // +2 because 1-based index and header is row 1
        };
      }
    }
  }

  // Check 2: Mobile Exact Match
  if (normPhone && normPhone.length >= 8) {
    for (let i = 0; i < existingContacts.length; i++) {
      const c = existingContacts[i];
      const cPhone1 = normalizePhoneDigits(c['Mobile']);
      const cPhone2 = normalizePhoneDigits(c['Điện thoại bàn']);
      if ((cPhone1 && (cPhone1 === normPhone || cPhone1.endsWith(normPhone) || normPhone.endsWith(cPhone1))) ||
          (cPhone2 && (cPhone2 === normPhone || cPhone2.endsWith(normPhone) || normPhone.endsWith(cPhone2)))) {
        return {
          status: 'MERGED',
          score: 100,
          reason: 'Exact Mobile/Phone match (' + normPhone + ')',
          existingContact: c,
          existingRowIndex: i + 2
        };
      }
    }
  }

  // Check 3: Full Name + Company Match
  if (normName && normName.length >= 3) {
    for (let i = 0; i < existingContacts.length; i++) {
      const c = existingContacts[i];
      const cName = normalizeVietnamese(c['Họ tên'] || c['Tên chuẩn hóa']);
      const cComp = normalizeVietnamese(c['Công ty / Tổ chức'] || c['Tên ngắn']);

      if (cName === normName) {
        if (normComp && cComp && (normComp.includes(cComp) || cComp.includes(normComp))) {
          // Strong name + company match
          return {
            status: 'MERGED',
            score: 90,
            reason: 'Strong Name + Company match: ' + c['Họ tên'] + ' @ ' + c['Công ty / Tổ chức'],
            existingContact: c,
            existingRowIndex: i + 2
          };
        } else if (normComp && cComp) {
          // Name identical but company differs -> REVIEW
          return {
            status: 'REVIEW',
            score: 65,
            reason: 'Same Name (' + c['Họ tên'] + ') but different Company (' + c['Công ty / Tổ chức'] + ' vs ' + (card.company || '') + ')',
            existingContact: c,
            existingRowIndex: i + 2
          };
        }
      }
    }
  }

  // If low OCR confidence or missing crucial contact info
  if ((card.confidence === 'Low') || (!normEmail && !normPhone && (!normName || normName.length < 2))) {
    return {
      status: 'REVIEW',
      score: 40,
      reason: 'Low data confidence or missing critical contact identifiers'
    };
  }

  return { status: 'NEW', score: 0, reason: 'No matching contact found' };
}

/**
 * Update existing contact with new/better fields from the scanned card
 */
function updateExistingContact(sheet, rowIndex, existing, card, sourceInfo) {
  const currentDuplicates = parseInt(existing['Số card trùng'] || '1', 10) || 1;
  const newDuplicates = currentDuplicates + 1;
  
  // Merge source card reference
  let currentSources = (existing['Nguồn card'] || '').split(/,\s*/).filter(Boolean);
  if (sourceInfo && !currentSources.includes(sourceInfo)) {
    currentSources.push(sourceInfo);
  }

  // Fill in missing fields if provided in new card
  const updated = {
    'Contact_ID': existing['Contact_ID'],
    'Họ tên': existing['Họ tên'] || card.full_name || '',
    'Tên chuẩn hóa': existing['Tên chuẩn hóa'] || card.normalized_name || '',
    'Xưng hô': existing['Xưng hô'] || card.salutation || '',
    'Chức danh (VI)': card.title_vi || existing['Chức danh (VI)'] || '',
    'Job Title (EN)': card.title_en || existing['Job Title (EN)'] || '',
    'Cấp bậc': card.seniority || existing['Cấp bậc'] || '',
    'Chức năng / Phòng ban': card.department || existing['Chức năng / Phòng ban'] || '',
    'Công ty / Tổ chức': card.company || existing['Công ty / Tổ chức'] || '',
    'Tên ngắn': card.short_company || existing['Tên ngắn'] || '',
    'Ngành cấp 1': card.industry_1 || existing['Ngành cấp 1'] || '',
    'Ngành cấp 2': card.industry_2 || existing['Ngành cấp 2'] || '',
    'Quốc gia': card.country || existing['Quốc gia'] || '',
    'Thành phố': card.city || existing['Thành phố'] || '',
    'Địa chỉ': card.address || existing['Địa chỉ'] || '',
    'Mobile': existing['Mobile'] || card.mobile || '',
    'Điện thoại bàn': existing['Điện thoại bàn'] || card.phone || '',
    'Fax': existing['Fax'] || card.fax || '',
    'Email chính': existing['Email chính'] || card.primary_email || '',
    'Email phụ': existing['Email phụ'] || card.secondary_email || '',
    'Website': existing['Website'] || card.website || '',
    'Vai trò CRM': card.crm_role || existing['Vai trò CRM'] || '',
    'Ngôn ngữ': card.language || existing['Ngôn ngữ'] || '',
    'Số card trùng': newDuplicates,
    'Nguồn card': currentSources.join(', '),
    'Độ tin cậy': card.confidence || existing['Độ tin cậy'] || 'High',
    'Ghi chú kiểm tra': [existing['Ghi chú kiểm tra'], card.review_note].filter(Boolean).join(' · ')
  };

  const rowValues = CONTACT_HEADERS.map(h => updated[h] !== undefined ? updated[h] : '');
  sheet.getRange(rowIndex, 1, 1, CONTACT_HEADERS.length).setValues([rowValues]);
}

/**
 * Build 27-element array corresponding to CONTACT_HEADERS for a new row
 */
function buildContactRow(contactId, card, duplicateCount, source) {
  const record = {
    'Contact_ID': contactId,
    'Họ tên': card.full_name || '',
    'Tên chuẩn hóa': card.normalized_name || normalizeVietnamese(card.full_name || ''),
    'Xưng hô': card.salutation || '',
    'Chức danh (VI)': card.title_vi || '',
    'Job Title (EN)': card.title_en || '',
    'Cấp bậc': card.seniority || '',
    'Chức năng / Phòng ban': card.department || '',
    'Công ty / Tổ chức': card.company || '',
    'Tên ngắn': card.short_company || '',
    'Ngành cấp 1': card.industry_1 || '',
    'Ngành cấp 2': card.industry_2 || '',
    'Quốc gia': card.country || 'Vietnam',
    'Thành phố': card.city || '',
    'Địa chỉ': card.address || '',
    'Mobile': card.mobile || '',
    'Điện thoại bàn': card.phone || '',
    'Fax': card.fax || '',
    'Email chính': card.primary_email || '',
    'Email phụ': card.secondary_email || '',
    'Website': card.website || '',
    'Vai trò CRM': card.crm_role || '',
    'Ngôn ngữ': card.language || 'Vietnamese / English',
    'Số card trùng': duplicateCount || 1,
    'Nguồn card': source || '',
    'Độ tin cậy': card.confidence || 'High',
    'Ghi chú kiểm tra': card.review_note || ''
  };

  return CONTACT_HEADERS.map(h => record[h] !== undefined ? record[h] : '');
}

/**
 * Save image to Google Drive folder with the standardized naming format:
 * NCxxxxx_[ContactID]_[Name]_[Company]_[YYYYMMDD]
 */
function saveCardImageToDrive(cardId, contactId, name, company, base64Data, mimeType, imageUrl) {
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const yyyymmdd = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd');
  
  const cleanName = sanitizeFileNamePart(name);
  const cleanCompany = sanitizeFileNamePart(company);
  const fileName = cardId + '_' + contactId + '_' + cleanName + '_' + cleanCompany + '_' + yyyymmdd + '.jpg';

  let blob;
  if (base64Data) {
    const cleanBase64 = base64Data.replace(/^data:image\/[a-z]+;base64,/, '');
    const decodedBytes = Utilities.base64Decode(cleanBase64);
    blob = Utilities.newBlob(decodedBytes, mimeType || 'image/jpeg', fileName);
  } else if (imageUrl) {
    const response = UrlFetchApp.fetch(imageUrl);
    blob = response.getBlob().setName(fileName);
  }

  if (!blob) return null;

  const file = folder.createFile(blob);
  file.setDescription('IDD BIM CPG Namecard auto-intake archive for ' + name + ' (' + contactId + ')');

  return {
    fileId: file.getId(),
    viewUrl: file.getUrl(),
    directUrl: 'https://drive.google.com/uc?export=view&id=' + file.getId()
  };
}

/**
 * Helper: Generate next Contact_ID (e.g. C001, C002, ..., C018)
 */
function generateNextContactId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'C001';

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  for (let i = 0; i < data.length; i++) {
    const val = String(data[i][0] || '');
    const match = val.match(/^C(\d+)$/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'C' + String(maxNum + 1).padStart(3, '0');
}

/**
 * Helper: Generate next Card_ID (e.g. NC00001, NC00002)
 */
function generateNextCardId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'NC00001';

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  for (let i = 0; i < data.length; i++) {
    const val = String(data[i][0] || '');
    const match = val.match(/^NC(\d+)$/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'NC' + String(maxNum + 1).padStart(5, '0');
}

/**
 * Helper: Generate next Review_ID (e.g. RV00001)
 */
function generateNextReviewId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 'RV00001';

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;
  for (let i = 0; i < data.length; i++) {
    const match = String(data[i][0] || '').match(/^RV(\d+)$/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'RV' + String(maxNum + 1).padStart(5, '0');
}

/**
 * Load existing contacts into structured objects
 */
function loadExistingContacts(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const headers = sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, CONTACT_HEADERS.length).getValues();

  return rows.map(r => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = r[idx] !== undefined ? String(r[idx]).trim() : '';
    });
    return obj;
  });
}

/**
 * String normalization helpers
 */
function normalizeString(str) {
  return String(str || '').toLowerCase().trim();
}

function normalizePhoneDigits(phone) {
  return String(phone || '').replace(/\D/g, '').replace(/^00/, '').replace(/^84/, '0');
}

function normalizeVietnamese(str) {
  if (!str) return '';
  return String(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function sanitizeFileNamePart(str) {
  return normalizeVietnamese(str)
    .replace(/[^a-zA-Z0-9]/g, '_')
    .replace(/_+/g, '_')
    .slice(0, 30);
}

function jsonResponse(data, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
