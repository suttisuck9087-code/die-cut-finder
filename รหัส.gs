// ══════════════════════════════════════════════════════════════════
//  Code.gs  —  Die Cut Finder Backend
//  เก็บรูปภาพใน Google Drive (folder) เพื่อให้คงอยู่ถาวร
// ══════════════════════════════════════════════════════════════════

var FOLDER_NAME = "DiecCut_Images";
var PROP_FOLDER_ID = "diecut_folder_id";
var PROP_PREFIX = "diecut_img_";
var SHARED_DRIVE_ID = "1mGHfUCBfw4aNv2S9S58gA8o4xD_5Z4YA";
// ── doGet: serve หน้าเว็บ ──────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Die Cut Finder')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── getOrCreateFolder: ดึง/สร้าง folder ใน Drive ─────────────
function getOrCreateFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(PROP_FOLDER_ID);

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // folder ถูกลบ — สร้างใหม่
    }
  }

  // สร้าง folder ใหม่
  var folder = DriveApp.createFolder(FOLDER_NAME);
  // ทำให้ anyone with link สามารถดูได้
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  props.setProperty(PROP_FOLDER_ID, folder.getId());
  return folder;
}

// ── saveImage: รับ base64 → บันทึกลง Drive → คืน fileId ────────
// dieId: รหัส die cut เช่น "PAR-3470"
// base64Data: data URL (data:image/jpeg;base64,xxxx)
function saveImage(dieId, base64Data) {
  try {
    var props = PropertiesService.getScriptProperties();
    var folder = getOrCreateFolder();

    // ลบรูปเก่าของ dieId นี้ก่อน (ถ้ามี)
    var oldFileId = props.getProperty(PROP_PREFIX + dieId);
    if (oldFileId) {
      try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (e) {}
    }

    // แยก base64 กับ mime type
    var match = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    if (!match) return { ok: false, error: 'รูปแบบรูปไม่ถูกต้อง' };

    var mimeType = match[1];
    var base64 = match[2];
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mimeType, 'die_' + dieId + '.jpg');

    // บันทึกลง Drive
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // บันทึก map: dieId → fileId
    props.setProperty(PROP_PREFIX + dieId, file.getId());

    return { ok: true, fileId: file.getId() };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}

// ── loadImage: โหลดรูปจาก Drive → คืน base64 data URL ─────────
function loadImage(dieId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var fileId = props.getProperty(PROP_PREFIX + dieId);
    if (!fileId) return { ok: true, data: null };

    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    var mimeType = blob.getContentType() || 'image/jpeg';

    return { ok: true, data: 'data:' + mimeType + ';base64,' + base64 };
  } catch (e) {
    return { ok: true, data: null }; // ไม่พบรูป — คืน null ไม่ error
  }
}

// ── loadAllImages: โหลดรูปทั้งหมดพร้อมกัน ────────────────────
// คืน object { "PAR-3470": "data:image/jpeg;base64,...", ... }
function loadAllImages() {
  try {
    var props = PropertiesService.getScriptProperties();
    var allProps = props.getProperties();
    var result = {};

    Object.keys(allProps).forEach(function(k) {
      if (k.indexOf(PROP_PREFIX) === 0) {
        var dieId = k.substring(PROP_PREFIX.length);
        var fileId = allProps[k];
        try {
          var file = DriveApp.getFileById(fileId);
          var blob = file.getBlob();
          var base64 = Utilities.base64Encode(blob.getBytes());
          var mimeType = blob.getContentType() || 'image/jpeg';
          result[dieId] = 'data:' + mimeType + ';base64,' + base64;
        } catch (e) {
          // ไฟล์อาจถูกลบ — ข้ามไป
          props.deleteProperty(k);
        }
      }
    });

    return { ok: true, images: result };
  } catch (e) {
    return { ok: false, error: e.toString(), images: {} };
  }
}

// ── deleteImage: ลบรูปของ dieId ──────────────────────────────
function deleteImage(dieId) {
  try {
    var props = PropertiesService.getScriptProperties();
    var fileId = props.getProperty(PROP_PREFIX + dieId);
    if (fileId) {
      try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      props.deleteProperty(PROP_PREFIX + dieId);
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}
function authorizeNow() {
  DriveApp.getRootFolder(); // บังคับขอสิทธิ์ Drive
}
function sendEmailWithImage(params) {
  try {
    var to = params.to_email;
    var subject = params.subject;
    var body = params.message;
    var imageData = params.imageData; // base64 data URL
    
    var attachments = [];
    
    if (imageData && imageData.indexOf('base64,') !== -1) {
      var match = imageData.match(/^data:([^;]+);base64,(.+)$/);
      if (match) {
        var mimeType = match[1];
        var base64 = match[2];
        var bytes = Utilities.base64Decode(base64);
        var blob = Utilities.newBlob(bytes, mimeType, 'รูปหลักฐาน.jpg');
        attachments.push(blob);
      }
    }
    
    var options = {
      name: 'Die Cut Finder',
      htmlBody: body.replace(/\n/g, '<br>'),
    };
    
    if (attachments.length > 0) {
      options.attachments = attachments;
    }
    
    GmailApp.sendEmail(to, subject, body, options);
    return { ok: true };
    
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}
