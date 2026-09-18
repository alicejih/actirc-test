/**
 * =========================================================================
 * 法路祖部落研習時數平臺 - Google Apps Script (GAS) 後端核心程式碼
 * 
 * 試算表 ID: 12lHpaiWIYeuBzTum8KNEfUPKvTm0PDpMrXODrf2sHa8
 * 雲端資料夾 ID: 1AbRkCu7zpt9tTT3tXDX3hjciZ0WOUH7o
 * 
 * 工作表架構說明：
 * 1. 工作表1 (時數填報總表)
 *    Col A (0): 姓名
 *    Col B (1): 身分證字號
 *    Col C (2): 開課單位 / 服務單位
 *    Col D (3): 課程編號
 *    Col E (4): 課程名稱
 *    Col F (5): 活動開始時間
 *    Col G (6): 活動結束時間
 *    Col H (7): 時數
 *    Col I (8): 審核欄位 (待審核 / 通過 / 不通過)
 *    Col J (9): 佐證圖片連結
 *    Col K (10): 紀錄ID (唯一識別碼 UUID)
 *    Col L (11): 管理員標記 (V)
 * 
 * 2. 學員名冊 (會員與管理員權限表)
 *    Col A (0): 姓名
 *    Col B (1): 身分證字號
 *    Col C (2): 出生年月日
 *    Col D (3): 服務單位
 *    Col E (4): 管理員欄位 (若值為 'V' 則具備審核後台權限)
 *    ... (相容 Col L (11) 亦可)
 * 
 * 3. course (課程分類字典檔)
 *    cId, cFace, cType
 * 
 * 4. 系統Log紀錄 (稽核軌跡)
 *    時間戳記, 動作類型, 操作人員, 詳細內容
 * =========================================================================
 */

const SPREADSHEET_ID = "12lHpaiWIYeuBzTum8KNEfUPKvTm0PDpMrXODrf2sHa8";
const UPLOAD_FOLDER_ID = "1AbRkCu7zpt9tTT3tXDX3hjciZ0WOUH7o";

const SHEET_RECORDS = "工作表1";
const SHEET_MEMBERS = "學員名冊";
const SHEET_COURSE  = "course";
const SHEET_LOG     = "系統Log紀錄";

/**
 * POST 請求分派入口
 */
function doPost(e) {
  try {
    const contents = e.postData ? e.postData.contents : "{}";
    const req = JSON.parse(contents);
    const action = req.action;

    // === 學員前台動作 ===
    if (action === "getCourses") {
      return jsonResponse(getCourses());
    } else if (action === "verifyUser") {
      return jsonResponse(verifyUser(req.idNumber, req.birthDate));
    } else if (action === "getDashboard") {
      return jsonResponse(getDashboard(req.idNumber, req.birthDate));
    } else if (action === "apply" || action === "logHoursData") {
      return jsonResponse(applyRecord(req));
    } else if (action === "update") {
      return jsonResponse(updateRecord(req));
    } else if (action === "delete") {
      return jsonResponse(deleteRecord(req));
    }
    
    // === 管理員後台動作 ===
    else if (action === "verifyAdmin") {
      return jsonResponse(verifyAdmin(req.idNumber, req.birthDate));
    } else if (action === "getAdminDashboard") {
      return jsonResponse(getAdminDashboard(req.idNumber, req.birthDate));
    } else if (action === "setRecordStatus") {
      return jsonResponse(setRecordStatus(req));
    }

    return jsonResponse({ status: "error", message: "未知的請求動作: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * GET 請求分派入口 (測試與探測)
 */
function doGet(e) {
  return jsonResponse({
    status: "online",
    system: "法路祖部落研習時數 API 伺服端",
    version: "2.1.0",
    time: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")
  });
}

// =========================================================================
// 核心驗證輔助函式
// =========================================================================

/**
 * 正規化身分證
 */
function cleanId(id) {
  return id ? id.toString().trim().toUpperCase() : "";
}

/**
 * 正規化出生日期 (轉為 YYYYMMDD)
 */
function normalizeBirthDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyyMMdd");
  }
  return val.toString().trim().replace(/[-/\s.]/g, "");
}

/**
 * 檢查是否具有管理員標記 (檢查是否為 'V')
 */
function isVMark(val) {
  if (!val) return false;
  const s = val.toString().trim().toUpperCase();
  return s === "V" || s === "TRUE" || s === "YES" || s === "Y";
}

/**
 * 從儲存格數值或公式中精準提取乾淨的 URL 網址
 * 完美支援純網址、=HYPERLINK("https://...", "標籤")、文字等多種型式
 */
function extractCleanUrl(val, formula) {
  const f = formula ? formula.toString().trim() : "";
  if (f) {
    const m = f.match(/https?:\/\/[^\s"'\)]+/i);
    if (m) return m[0];
  }
  const v = val ? val.toString().trim() : "";
  if (v) {
    const m = v.match(/https?:\/\/[^\s"'\)]+/i);
    if (m) return m[0];
  }
  return "";
}

/**
 * 從「學員名冊」驗證身分並回傳學員資訊
 */
function findMember(idNumber, birthDate) {
  const targetId = cleanId(idNumber);
  const targetDob = normalizeBirthDate(birthDate);

  if (!targetId || !targetDob) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet) throw new Error("找不到工作表：" + SHEET_MEMBERS);

  const values = sheet.getDataRange().getValues();
  // 第一列為表頭，從第二列開始比對 (i = 1)
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = cleanId(row[1]); // Col B: 身分證
    const rowDob = normalizeBirthDate(row[2]); // Col C: 出生日期

    if (rowId === targetId && rowDob === targetDob) {
      // 檢查 Col E (第 5 欄，索引 4) 或 Col L (第 12 欄，索引 11) 是否標記為 V
      const isAdmin = isVMark(row[4]) || isVMark(row[11]);
      return {
        rowIndex: i + 1,
        name: row[0] ? row[0].toString().trim() : "",
        idNumber: rowId,
        unit: row[3] ? row[3].toString().trim() : "",
        isAdmin: isAdmin
      };
    }
  }

  return null;
}

// =========================================================================
// 前台學員功能
// =========================================================================

/**
 * 取得課程字典分類 (依大分類群組化)
 */
function getCourses() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_COURSE);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const cFace = row[1] ? row[1].toString().trim() : "";
    const cType = row[2] ? row[2].toString().trim() : "";
    if (!cFace) continue;

    if (!map[cFace]) {
      map[cFace] = new Set();
    }
    if (cType) {
      map[cFace].add(cType);
    }
  }

  const result = [];
  for (const face in map) {
    result.push({
      mainCategory: face,
      subCategories: Array.from(map[face])
    });
  }
  return result;
}

/**
 * 學員身分驗證
 */
function verifyUser(idNumber, birthDate) {
  const user = findMember(idNumber, birthDate);
  if (!user) {
    return { status: "error", message: "身分證字號或出生年月日核對不符！" };
  }
  return { status: "success", user: user };
}

/**
 * 取得學員個人儀表板 (包含時數與歷史清單)
 */
function getDashboard(idNumber, birthDate) {
  const user = findMember(idNumber, birthDate);
  if (!user) {
    return { status: "error", message: "身分驗證失敗，請確認身分證與出生日期！" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) throw new Error("找不到工作表：" + SHEET_RECORDS);

  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();
  const records = [];
  let totalHours = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = cleanId(row[1]); // Col B
    if (rowId !== user.idNumber) continue;

    const courseName = row[4] ? row[4].toString().trim() : "";
    const categoryFull = row[3] ? row[3].toString().trim() : ""; // 課程分類/編號
    let mainCat = categoryFull;
    let subCat = "";
    if (categoryFull.includes(">")) {
      const parts = categoryFull.split(">");
      mainCat = parts[0].trim();
      subCat = parts[1].trim();
    }

    let sTime = row[5];
    if (sTime instanceof Date) {
      sTime = Utilities.formatDate(sTime, "Asia/Taipei", "yyyy-MM-dd HH:mm");
    } else {
      sTime = sTime ? sTime.toString().trim() : "";
    }

    let eTime = row[6];
    if (eTime instanceof Date) {
      eTime = Utilities.formatDate(eTime, "Asia/Taipei", "yyyy-MM-dd HH:mm");
    } else {
      eTime = eTime ? eTime.toString().trim() : "";
    }

    const h = parseFloat(row[7]) || 0;
    const status = row[8] ? row[8].toString().trim() : "待審核";
    const fileUrl = extractCleanUrl(row[9], formulas[i] ? formulas[i][9] : "");
    const recId = row[10] ? row[10].toString().trim() : ("ROW_" + (i + 1));

    if (status === "通過" || status === "已審核") {
      totalHours += h;
    }

    // 只有「待審核」狀態才允許學員自行修改或刪除
    const canEditOrDelete = (status === "待審核" || status === "");

    records.push({
      recordId: recId,
      courseName: courseName,
      mainCategory: mainCat,
      subCategory: subCat,
      startTime: sTime,
      endTime: eTime,
      hours: h,
      status: status || "待審核",
      canEditOrDelete: canEditOrDelete,
      fileUrl: fileUrl
    });
  }

  // 排序：依活動時間新到舊
  records.sort((a, b) => (b.startTime || "").localeCompare(a.startTime || ""));

  return {
    status: "success",
    user: user,
    totalHours: Math.round(totalHours * 10) / 10,
    records: records
  };
}

/**
 * 申報新時數 (含圖片上傳至 Google Drive)
 */
function applyRecord(req) {
  let user = null;
  if (req.birthDate) {
    user = findMember(req.idNumber, req.birthDate);
  } else if (req.userName && req.idNumber) {
    user = {
      name: req.userName,
      idNumber: cleanId(req.idNumber),
      unit: req.courseUnit || "原資中心"
    };
  }

  if (!user) {
    return { status: "error", message: "身分驗證失敗，無法送出申報！" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) throw new Error("找不到工作表：" + SHEET_RECORDS);

  let fileUrl = "";
  if (req.fileData && req.fileData.base64) {
    try {
      const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
      const decodedBytes = Utilities.base64Decode(req.fileData.base64);
      const fileName = user.name + "_" + (req.courseName || "研習佐證") + "_" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss") + ".jpg";
      const blob = Utilities.newBlob(decodedBytes, req.fileData.mimeType || "image/jpeg", fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    } catch (e) {
      console.warn("上傳照片失敗: " + e.toString());
    }
  }

  // 確保寫入純乾淨的 URL 網址字串，絕不使用 =HYPERLINK 公式
  const pureFileUrl = extractCleanUrl(fileUrl, "");
  const recId = "REC_" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMddHHmmss") + "_" + Math.floor(Math.random() * 1000);
  const hoursNum = parseFloat(req.hours) || 0;

  // 寫入工作表1
  sheet.appendRow([
    user.name,                                 // Col A (0): 姓名
    user.idNumber,                             // Col B (1): 身分證字號
    req.courseUnit || user.unit || "原資中心",  // Col C (2): 開課單位/單位
    req.courseClass || "",                     // Col D (3): 課程分類
    req.courseName || "",                      // Col E (4): 課程名稱
    req.startTime || "",                       // Col F (5): 開始時間
    req.endTime || "",                         // Col G (6): 結束時間
    hoursNum,                                  // Col H (7): 時數
    "待審核",                                  // Col I (8): 審核欄位
    pureFileUrl,                               // Col J (9): 佐證連結 (純 URL 網址字串)
    recId,                                     // Col K (10): 紀錄ID
    ""                                         // Col L (11): 管理員
  ]);

  writeLog("學員申報", user.name + " (" + user.idNumber + ")", `課程：${req.courseName}，時數：${hoursNum}，ID：${recId}`);

  return { status: "success", message: "研習時數已成功申報！請等待管理員審核。" };
}

/**
 * 學員修改待審核紀錄
 */
function updateRecord(req) {
  const user = findMember(req.idNumber, req.birthDate);
  if (!user) return { status: "error", message: "身分驗證失敗" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = cleanId(row[1]);
    const recId = row[10] ? row[10].toString().trim() : ("ROW_" + (i + 1));

    if (rowId === user.idNumber && recId === req.recordId) {
      const status = row[8] ? row[8].toString().trim() : "待審核";
      if (status === "通過" || status === "已審核") {
        return { status: "error", message: "該筆紀錄已核定通過，無法修改！" };
      }

      // 更新欄位
      const rNum = i + 1;
      if (req.courseClass) sheet.getRange(rNum, 4).setValue(req.courseClass);
      if (req.courseName) sheet.getRange(rNum, 5).setValue(req.courseName);
      if (req.startTime) sheet.getRange(rNum, 6).setValue(req.startTime);
      if (req.endTime) sheet.getRange(rNum, 7).setValue(req.endTime);
      if (req.hours) sheet.getRange(rNum, 8).setValue(parseFloat(req.hours) || 0);

      writeLog("學員修改", user.name, `修改紀錄ID：${req.recordId}，課程：${req.courseName}`);
      return { status: "success", message: "紀錄已成功更新！" };
    }
  }

  return { status: "error", message: "找不到指定的紀錄或無權限修改！" };
}

/**
 * 學員刪除待審核紀錄
 */
function deleteRecord(req) {
  const user = findMember(req.idNumber, req.birthDate);
  if (!user) return { status: "error", message: "身分驗證失敗" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = cleanId(row[1]);
    const recId = row[10] ? row[10].toString().trim() : ("ROW_" + (i + 1));

    if (rowId === user.idNumber && recId === req.recordId) {
      const status = row[8] ? row[8].toString().trim() : "待審核";
      if (status === "通過" || status === "已審核") {
        return { status: "error", message: "該筆紀錄已核定通過，不允許刪除！" };
      }

      sheet.deleteRow(i + 1);
      writeLog("學員撤銷", user.name, `刪除紀錄ID：${req.recordId}`);
      return { status: "success", message: "紀錄已成功刪除！" };
    }
  }

  return { status: "error", message: "找不到指定的紀錄！" };
}

// =========================================================================
// 後台管理員專屬功能 (依 Column L 判定權限)
// =========================================================================

/**
 * 驗證管理員身分 (必須有 Column L == 'V')
 */
function verifyAdmin(idNumber, birthDate) {
  const member = findMember(idNumber, birthDate);
  if (!member) {
    return { status: "error", message: "查無此身分證字號或出生年月日不正確！" };
  }

  if (!member.isAdmin) {
    return {
      status: "error",
      message: "抱歉，您不具備審核管理員權限！(試算表管理員欄位需標記為 V)"
    };
  }

  writeLog("管理員登入", member.name, `身分證：${member.idNumber} 登入審核後台`);
  return {
    status: "success",
    admin: {
      name: member.name,
      idNumber: member.idNumber,
      unit: member.unit,
      isAdmin: true
    }
  };
}

/**
 * 取得管理員審核總覽儀表板 (全部申報清單與統計)
 */
function getAdminDashboard(idNumber, birthDate) {
  const verifyRes = verifyAdmin(idNumber, birthDate);
  if (verifyRes.status !== "success") {
    return verifyRes;
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) throw new Error("找不到工作表：" + SHEET_RECORDS);

  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();
  const allList = [];
  let pendingCount = 0;
  let approvedCount = 0;
  let rejectedCount = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const studentName = row[0] ? row[0].toString().trim() : "";
    const studentId = cleanId(row[1]);
    if (!studentName && !studentId) continue;

    const unit = row[2] ? row[2].toString().trim() : "";
    const courseClass = row[3] ? row[3].toString().trim() : "";
    const courseName = row[4] ? row[4].toString().trim() : "";

    let sTime = row[5];
    if (sTime instanceof Date) {
      sTime = Utilities.formatDate(sTime, "Asia/Taipei", "yyyy-MM-dd HH:mm");
    } else {
      sTime = sTime ? sTime.toString().trim() : "";
    }

    let eTime = row[6];
    if (eTime instanceof Date) {
      eTime = Utilities.formatDate(eTime, "Asia/Taipei", "yyyy-MM-dd HH:mm");
    } else {
      eTime = eTime ? eTime.toString().trim() : "";
    }

    const hours = parseFloat(row[7]) || 0;
    const status = row[8] ? row[8].toString().trim() : "待審核";
    const fileUrl = extractCleanUrl(row[9], formulas[i] ? formulas[i][9] : "");
    let recId = row[10] ? row[10].toString().trim() : "";

    // 若舊資料沒有 UUID，自動補建一個方便後台即時更新
    if (!recId) {
      recId = "ROW_" + (i + 1);
      sheet.getRange(i + 1, 11).setValue(recId);
    }

    if (status === "通過" || status === "已審核") {
      approvedCount++;
    } else if (status === "不通過") {
      rejectedCount++;
    } else {
      pendingCount++;
    }

    allList.push({
      recordId: recId,
      name: studentName,
      idNumber: studentId,
      unit: unit,
      courseClass: courseClass,
      courseName: courseName,
      startTime: sTime,
      endTime: eTime,
      hours: hours,
      status: status || "待審核",
      fileUrl: fileUrl
    });
  }

  // 排序：待審核排在最前面，接著依時間由新至舊
  allList.sort((a, b) => {
    const aPending = (a.status === "待審核" || !a.status);
    const bPending = (b.status === "待審核" || !b.status);
    if (aPending && !bPending) return -1;
    if (!aPending && bPending) return 1;
    return (b.startTime || "").localeCompare(a.startTime || "");
  });

  return {
    status: "success",
    admin: verifyRes.admin,
    stats: {
      total: allList.length,
      pending: pendingCount,
      approved: approvedCount,
      rejected: rejectedCount
    },
    records: allList
  };
}

/**
 * 管理員設定審核結果 (通過 / 不通過)
 */
function setRecordStatus(req) {
  const verifyRes = verifyAdmin(req.idNumber, req.birthDate);
  if (verifyRes.status !== "success") {
    return verifyRes;
  }

  const recId = req.recordId;
  const newStatus = req.newStatus; // "通過" 或 "不通過"

  if (!recId || !newStatus) {
    return { status: "error", message: "缺少必要審核參數！" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const currentRowRecId = row[10] ? row[10].toString().trim() : ("ROW_" + (i + 1));

    if (currentRowRecId === recId) {
      const rNum = i + 1;
      // Col I (第 9 欄) 寫入審核結果
      sheet.getRange(rNum, 9).setValue(newStatus);
      SpreadsheetApp.flush();

      const studentName = row[0] || "";
      const courseName = row[4] || "";

      writeLog("管理員審核", verifyRes.admin.name, `紀錄ID：${recId}，學員：${studentName}，課程：${courseName}，審定結果：${newStatus}`);

      return {
        status: "success",
        message: `已成功將「${studentName} - ${courseName}」之審核結果設定為【${newStatus}】！`
      };
    }
  }

  return { status: "error", message: "找不到該筆申報紀錄！" };
}

// =========================================================================
// 稽核與日誌
// =========================================================================

/**
 * 寫入系統日誌
 */
function writeLog(actionType, operator, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName(SHEET_LOG);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_LOG);
      logSheet.appendRow(["時間戳記", "動作類型", "操作人員", "詳細內容"]);
    }
    const nowStr = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
    logSheet.appendRow([nowStr, actionType, operator, details]);
  } catch (err) {
    console.warn("寫入 Log 異常: " + err.toString());
  }
}

/**
 * 產生標準 JSON 輸出
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
