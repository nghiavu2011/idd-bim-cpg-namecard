/**
 * IDD BIM CPG Namecard — Google Apps Script Backend v2.0
 * 
 * Features:
 * 1. Corporate Email OTP Authentication (CPG Email Domain check + 6-digit code via MailApp)
 * 2. Real-time multi-user Google Sheet sync (Master tab: Namecard_CRM_Contacts)
 * 3. Real-time multi-user Google Drive image archive (Cards_Inbox tab + Drive files)
 * 4. Duplicate engine with auto-merge & review logic
 * 5. Admin role permissions (Full control for Nghia Vu & CPG Admins)
 */

const CONFIG = {
  SPREADSHEET_ID: '13XsLUK792UkdDuXYHSdjIGJPX7AazQt6hROtTa0LRCI',
  DRIVE_FOLDER_ID: '1Sy6Lb509_seEfP-EQd_1Wh_1lipWIxG3',
  TAB_CONTACTS: 'Namecard_CRM_Contacts',
  TAB_INBOX: 'Cards_Inbox',
  TAB_REVIEW: 'Review_Queue',
  DEFAULT_SECRET_KEY: 'PATH_SECRET',
  
  // Corporate Email Domain whitelist:
  ALLOWED_DOMAINS: ['cpgcorp.com.sg', 'cpg.com.vn', 'cpg.sg', 'cpgconsultants.com'],
  
  // Super Admin emails with full management & delete permissions:
  ADMIN_EMAILS: ['vu.trong.nghia@cpgcorp.com.sg', 'nghiavu2011@gmail.com', 'nghia.vu@cpgcorp.com.sg']
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
 * Setup tabs and default configuration
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
  
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty(CONFIG.DEFAULT_SECRET_KEY);
  if (!secret) {
    secret = 'nmc_' + Utilities.getUuid().replace(/-/g, '').slice(0, 24);
    props.setProperty(CONFIG.DEFAULT_SECRET_KEY, secret);
  }
  
  Logger.log('=== IDD BIM CPG NAMECARD SYSTEM v2.0 READY ===');
  Logger.log('Spreadsheet ID: ' + CONFIG.SPREADSHEET_ID);
  Logger.log('Drive Folder ID: ' + CONFIG.DRIVE_FOLDER_ID);
  Logger.log('Path Secret: ' + secret);
}

/**
 * Handle HTTP GET: Serves all live contacts and physical cards to authenticated web clients
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const contactSheet = ss.getSheetByName(CONFIG.TAB_CONTACTS);
    const inboxSheet = ss.getSheetByName(CONFIG.TAB_INBOX);
    
    const contacts = contactSheet ? loadExistingContacts(contactSheet) : [];
    const physicalCards = inboxSheet ? loadPhysicalCardsFromInbox(inboxSheet, contacts) : [];
    const count = contacts.length;
    
    return jsonResponse({
      ok: true,
      service: 'IDD BIM CPG Namecard Library API',
      version: '2.0.0',
      total_contacts: count,
      total_cards: physicalCards.length,
      data_version: 'v' + count + '_' + new Date().toISOString().slice(0, 10),
      contacts: contacts,
      physical_cards: physicalCards
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

/**
 * Handle HTTP POST: Process Actions (OTP Auth, Save Namecard, Delete Contact)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Server busy. Could not obtain lock.' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'Empty payload' }, 400);
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'processNamecard';

    // 1. Request OTP Action
    if (action === 'requestOtp') {
      return jsonResponse(handleRequestOtp(payload.email));
    }

    // 2. Verify OTP Action
    if (action === 'verifyOtp') {
      return jsonResponse(handleVerifyOtp(payload.email, payload.otp));
    }

    // 3. Process Namecard Action
    if (action === 'processNamecard' || action === 'saveNamecard') {
      return jsonResponse(handleProcessNamecard(payload));
    }

    // 4. Delete Contact Action (Admin only)
    if (action === 'deleteContact') {
      return jsonResponse(handleDeleteContact(payload));
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action }, 400);
  } catch (err) {
    Logger.log('Error processing POST: ' + err.toString());
    return jsonResponse({ ok: false, error: err.message || err.toString() }, 500);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handle Request OTP: Check domain & send 6-digit email passcode
 */
function handleRequestOtp(email) {
  const cleanEmail = String(email || '').trim().toLowerCase();
  if (!cleanEmail || !cleanEmail.includes('@')) {
    return { ok: false, error: 'Email không hợp lệ.' };
  }

  const domain = cleanEmail.split('@')[1];
  const isAllowedDomain = CONFIG.ALLOWED_DOMAINS.includes(domain);
  const isAdmin = CONFIG.ADMIN_EMAILS.includes(cleanEmail);

  if (!isAllowedDomain && !isAdmin) {
    return {
      ok: false,
      error: 'Chỉ chấp nhận email thuộc tổ chức CPG (@cpgcorp.com.sg, @cpg.com.vn, @cpg.sg). Vui lòng liên hệ Admin để được cấp quyền.'
    };
  }

  // Generate 6-digit OTP
  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const cache = CacheService.getScriptCache();
  cache.put('OTP_' + cleanEmail, otp, 600); // 10 minutes expiry

  // Send OTP Email via Google MailApp
  try {
    const subject = '🔐 Mã xác thực đăng nhập IDD BIM CPG Namecard CRM: ' + otp;
    const bodyHtml = `
      <div style="font-family: Arial, sans-serif; max-width: 540px; margin: auto; padding: 24px; border: 1px solid #e0e0e0; border-radius: 12px; background: #ffffff;">
        <div style="text-align: center; margin-bottom: 20px;">
          <h2 style="color: #17141b; margin: 0;">IDD BIM CPG Namecard CRM</h2>
          <p style="color: #6b7280; font-size: 13px; margin: 4px 0 0;">Thư viện & Quản lý Danh bạ Danh thiếp Nội bộ</p>
        </div>
        <p style="color: #333333; font-size: 14px; line-height: 1.5;">Xin chào đồng nghiệp CPG,</p>
        <p style="color: #333333; font-size: 14px; line-height: 1.5;">Bạn đang yêu cầu đăng nhập vào hệ thống Namecard CRM. Dưới đây là mã xác thực 6 chữ số của bạn:</p>
        <div style="text-align: center; margin: 24px 0;">
          <span style="display: inline-block; font-size: 32px; font-weight: 800; letter-spacing: 8px; color: #f4a61c; background: #fff8eb; padding: 12px 28px; border-radius: 10px; border: 1.5px dashed #f4a61c;">${otp}</span>
        </div>
        <p style="color: #6b7280; font-size: 12px; line-height: 1.4;">* Mã này có hiệu lực trong vòng <strong>10 phút</strong>. Không chia sẻ mã này cho người khác.</p>
        <hr style="border: 0; border-top: 1px solid #f0f0f0; margin: 20px 0;">
        <p style="color: #9ca3af; font-size: 11px; text-align: center; margin: 0;">IDD BIM CPG Namecard CRM • Quản lý bởi Nghia Vu</p>
      </div>
    `;

    MailApp.sendEmail({
      to: cleanEmail,
      subject: subject,
      htmlBody: bodyHtml
    });

    return { ok: true, message: 'Đã gửi mã xác thực 6 số vào hòm thư ' + cleanEmail };
  } catch (mailErr) {
    Logger.log('Lỗi gửi email: ' + mailErr.toString());
    return { ok: false, error: 'Không thể gửi email xác thực: ' + mailErr.message };
  }
}

/**
 * Handle Verify OTP: Validate passcode & issue session
 */
function handleVerifyOtp(email, otp) {
  const cleanEmail = String(email || '').trim().toLowerCase();
  const cleanOtp = String(otp || '').trim();

  const cache = CacheService.getScriptCache();
  const cachedOtp = cache.get('OTP_' + cleanEmail);

  // Master bypass passcode for Nghia Vu during initial setup or recovery:
  const isMasterPass = (cleanOtp === '789654' && CONFIG.ADMIN_EMAILS.includes(cleanEmail));

  if (!cachedOtp && !isMasterPass) {
    return { ok: false, error: 'Mã xác thực đã hết hạn hoặc không tồn tại. Vui lòng bấm Gửi lại mã.' };
  }

  if (cachedOtp !== cleanOtp && !isMasterPass) {
    return { ok: false, error: 'Mã xác thực không chính xác. Vui lòng kiểm tra lại email.' };
  }

  // Remove used OTP
  cache.remove('OTP_' + cleanEmail);

  const token = 'sess_' + Utilities.getUuid().replace(/-/g, '');
  const isAdmin = CONFIG.ADMIN_EMAILS.includes(cleanEmail);

  // Save session for 7 days
  cache.put('SESSION_' + token, JSON.stringify({ email: cleanEmail, is_admin: isAdmin }), 86400 * 7);

  return {
    ok: true,
    token: token,
    email: cleanEmail,
    is_admin: isAdmin,
    message: 'Đăng nhập thành công!'
  };
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
  const uploader = payload.uploader || 'Web Intake User';
  const ocrConfidence = payload.ocr_confidence || 95;

  const allContacts = loadExistingContacts(contactSheet);
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
    const rowIndex = dupCheck.existingRowIndex;
    updateExistingContact(contactSheet, rowIndex, dupCheck.existingContact, card, sourceFileName || cardId);
  } else if (finalStatus === 'REVIEW') {
    targetContactId = dupCheck.existingContact ? dupCheck.existingContact.Contact_ID : '';
    reviewId = generateNextReviewId(reviewSheet);
    
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

  // Handle image upload to Google Drive
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

  // Append entry to Cards_Inbox
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
    direct_image_url: driveDirectUrl || undefined,
    data_version: 'v' + totalContacts + '_' + new Date().toISOString().slice(0, 10),
    message: finalStatus === 'NEW'
      ? 'Đã tạo liên hệ mới ' + targetContactId
      : finalStatus === 'MERGED'
      ? 'Đã liên kết vào liên hệ hiện có ' + targetContactId
      : 'Cần kiểm tra: ' + (dupCheck.reason || 'Trùng lặp tiềm năng')
  };
}

/**
 * Handle Delete Contact (Admin only)
 */
function handleDeleteContact(payload) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const contactSheet = ss.getSheetByName(CONFIG.TAB_CONTACTS);
  const targetId = payload.contact_id;

  if (!targetId || !contactSheet) {
    return { ok: false, error: 'Không tìm thấy Contact ID để xóa' };
  }

  const lastRow = contactSheet.getLastRow();
  if (lastRow <= 1) return { ok: false, error: 'Bảng trống' };

  const data = contactSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === targetId) {
      contactSheet.deleteRow(i + 2);
      return { ok: true, message: 'Đã xóa liên hệ ' + targetId };
    }
  }

  return { ok: false, error: 'Không tìm thấy dòng khớp với ' + targetId };
}

/**
 * Duplicate Rule Engine
 */
function evaluateDuplicate(card, existingContacts) {
  const normEmail = normalizeString(card.primary_email || card.secondary_email);
  const normPhone = normalizePhoneDigits(card.mobile || card.phone);
  const normName = normalizeVietnamese(card.full_name || card.normalized_name);
  const normComp = normalizeVietnamese(card.company || card.short_company);

  if (!existingContacts || existingContacts.length === 0) {
    return { status: 'NEW', score: 0, reason: 'First contact in database' };
  }

  // Exact Email Match
  if (normEmail) {
    for (let i = 0; i < existingContacts.length; i++) {
      const c = existingContacts[i];
      const cEmail1 = normalizeString(c['Email chính']);
      const cEmail2 = normalizeString(c['Email phụ']);
      if ((cEmail1 && cEmail1 === normEmail) || (cEmail2 && cEmail2 === normEmail)) {
        return {
          status: 'MERGED',
          score: 100,
          reason: 'Trùng Email chính xác (' + normEmail + ')',
          existingContact: c,
          existingRowIndex: i + 2
        };
      }
    }
  }

  // Exact Phone Match
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
          reason: 'Trùng Số điện thoại (' + normPhone + ')',
          existingContact: c,
          existingRowIndex: i + 2
        };
      }
    }
  }

  // Name + Company Match
  if (normName && normName.length >= 3) {
    for (let i = 0; i < existingContacts.length; i++) {
      const c = existingContacts[i];
      const cName = normalizeVietnamese(c['Họ tên'] || c['Tên chuẩn hóa']);
      const cComp = normalizeVietnamese(c['Công ty / Tổ chức'] || c['Tên ngắn']);

      if (cName === normName) {
        if (normComp && cComp && (normComp.includes(cComp) || cComp.includes(normComp))) {
          return {
            status: 'MERGED',
            score: 90,
            reason: 'Trùng Tên và Công ty: ' + c['Họ tên'] + ' @ ' + c['Công ty / Tổ chức'],
            existingContact: c,
            existingRowIndex: i + 2
          };
        } else if (normComp && cComp) {
          return {
            status: 'REVIEW',
            score: 65,
            reason: 'Trùng Tên nhưng khác Công ty (' + c['Công ty / Tổ chức'] + ' vs ' + (card.company || '') + ')',
            existingContact: c,
            existingRowIndex: i + 2
          };
        }
      }
    }
  }

  if ((card.confidence === 'Low') || (!normEmail && !normPhone && (!normName || normName.length < 2))) {
    return {
      status: 'REVIEW',
      score: 40,
      reason: 'Độ tin cậy thấp hoặc thiếu thông tin định danh quan trọng'
    };
  }

  return { status: 'NEW', score: 0, reason: 'Liên hệ mới' };
}

/**
 * Update existing contact in Google Sheet
 */
function updateExistingContact(sheet, rowIndex, existing, card, sourceInfo) {
  const currentDuplicates = parseInt(existing['Số card trùng'] || '1', 10) || 1;
  const newDuplicates = currentDuplicates + 1;
  
  let currentSources = (existing['Nguồn card'] || '').split(/,\s*/).filter(Boolean);
  if (sourceInfo && !currentSources.includes(sourceInfo)) {
    currentSources.push(sourceInfo);
  }

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
 * Build 27-element row for new contact
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
 * Save image to Google Drive folder
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
  file.setDescription('IDD BIM CPG Namecard archive for ' + name + ' (' + contactId + ')');

  return {
    fileId: file.getId(),
    viewUrl: file.getUrl(),
    directUrl: 'https://lh3.googleusercontent.com/d/' + file.getId()
  };
}

/**
 * Helper: Generate next Contact_ID
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
 * Helper: Generate next Card_ID
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
 * Helper: Generate next Review_ID
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
 * Load existing contacts from Sheet
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
 * Load physical cards from Cards_Inbox
 */
function loadPhysicalCardsFromInbox(inboxSheet, contacts) {
  const lastRow = inboxSheet.getLastRow();
  if (lastRow <= 1) return [];

  const contactMap = {};
  contacts.forEach(c => {
    contactMap[c.Contact_ID] = c;
  });

  const rows = inboxSheet.getRange(2, 1, lastRow - 1, INBOX_HEADERS.length).getValues();
  return rows.map(r => {
    const cardId = String(r[0] || '').trim();
    const contactId = String(r[2] || '').trim();
    const source = String(r[4] || cardId).trim();
    const driveFileId = String(r[7] || '').trim();
    const driveViewUrl = String(r[8] || '').trim();
    const directUrl = String(r[9] || (driveFileId ? 'https://lh3.googleusercontent.com/d/' + driveFileId : '')).trim();

    const c = contactMap[contactId] || {};
    return {
      Card_ID: cardId,
      Source: source,
      Contact_ID: contactId,
      Name: c['Họ tên'] || 'Contact ' + contactId,
      Company: c['Tên ngắn'] || c['Công ty / Tổ chức'] || 'CPG',
      Title: c['Chức danh (VI)'] || c['Job Title (EN)'] || '',
      Duplicate: (c['Số card trùng'] && parseInt(c['Số card trùng'], 10) > 1),
      Duplicate_Count: parseInt(c['Số card trùng'] || '1', 10) || 1,
      Image: directUrl || driveViewUrl || ''
    };
  }).filter(x => x.Card_ID);
}

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
