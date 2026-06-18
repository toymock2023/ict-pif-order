/**
 * PIF 難民特賣會 - Google Apps Script
 * 接收網站訂單並寫入 Google Sheets
 *
 * 部署方式請參考「部署說明.md」
 */

// 試算表 ID（從你的 Google Sheets URL 取得，例如：https://docs.google.com/spreadsheets/d/【這串】/edit）
const SPREADSHEET_ID = "請填入你的試算表 ID";

// 主工作表名稱（用來匯整訂單摘要）
const SUMMARY_SHEET = "訂單總覽";

// 明細工作表名稱（每一個訂單品項一列，方便統計）
const DETAIL_SHEET = "訂單明細";

// 推廣人(分潤)名單工作表。欄位順序固定為：
//   推廣代碼 | 姓名 | 分潤% | 查詢金鑰 | 啟用(TRUE/FALSE) | 備註
// 「推廣代碼」是給客人看的網址 ?ref= 值(可好記，如 alice)。
// 「查詢金鑰」是給推廣人查成績用的密碼(不好猜，如 alice-7fK2)。
const SELLER_SHEET = "推廣人名單";

// 網站根網址(結尾要有 /)。用來在「推廣人名單」G 欄自動產生推廣連結。
// 若日後換網域，只要改這一行即可。
const SITE_ROOT_URL = "https://toymock2023.github.io/ict-pif-order/";


/**
 * 接收 POST 請求，處理訂單寫入
 */
function doPost(e) {
  try {
    const order = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ★ 驗證推廣人代碼(白名單)。無效或空白一律記為空字串(視為直營訂單)。
    const validRef = validateRef(ss, order.ref);

    // 1. 寫入「訂單總覽」工作表
    //    基本欄位用固定順序寫(訂單時間 ... 備註，共 16 欄)，
    //    「推廣人」則依標題名稱寫到它實際所在的欄(可能因你自訂欄位而在 AA 等位置)，
    //    這樣即使你在中間插入自己的欄位也不會錯位。
    const summarySheet = getOrCreateSheet(ss, SUMMARY_SHEET, [
      "訂單時間", "訂單編號", "姓名", "公司/部門", "聯絡電話", "Email",
      "取貨方式", "宅配地址", "付款方式", "希望取貨日",
      "購買品項摘要", "總數量",
      "未稅金額", "營業稅 5%", "應付總金額(未含運費)",
      "備註",
      "推廣人"
    ]);

    const orderId = generateOrderId();

    // 先寫入基本 16 欄(到「備註」為止，維持原本固定順序)
    summarySheet.appendRow([
      order.timestamp,
      orderId,
      order.name,
      order.company,
      order.phone,
      order.email,
      order.delivery,
      order.address,
      order.payment,
      order.pickupDate,
      order.itemSummary,
      order.totalQty,
      order.subtotalAmount,
      order.taxAmount,
      order.totalAmount,
      order.note
    ]);

    // ★ 把「推廣人」寫進標題為「推廣人」的那一欄(用名稱定位，不寫死欄號)
    try {
      const summaryRow = summarySheet.getLastRow();
      const sHeaders = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).getValues()[0];
      const refColIdx = sHeaders.indexOf("推廣人") + 1; // indexOf 回傳 0-based，+1 變欄號
      if (refColIdx > 0) {
        summarySheet.getRange(summaryRow, refColIdx).setValue(validRef);
      }
    } catch (refErr) {
      console.error("寫入推廣人欄失敗:", refErr);
    }

    // ★ 強制把「聯絡電話」欄位設為純文字格式，避免開頭 0 被吃掉
    // (用 try 包起來，避免單一格式問題擋下整筆訂單)
    try {
      const summaryRow2 = summarySheet.getLastRow();
      forceTextCell(summarySheet, summaryRow2, ["聯絡電話"], { 聯絡電話: order.phone });
    } catch (fmtErr) {
      console.error("總覽聯絡電話格式化失敗:", fmtErr);
    }

    // 2. 寫入「訂單明細」工作表 (每品項一列)
    const detailSheet = getOrCreateSheet(ss, DETAIL_SHEET, [
      "訂單時間", "訂單編號", "姓名", "公司/部門", "聯絡電話",
      "商品編號", "商品名稱", "規格", "條碼", "未稅單價", "數量", "未稅小計"
    ]);

    order.items.forEach(item => {
      detailSheet.appendRow([
        order.timestamp,
        orderId,
        order.name,
        order.company,
        order.phone,
        item.no,
        item.name,
        item.spec || "",
        item.barcode,
        item.price,
        item.qty,
        item.subtotal
      ]);
      // ★ 把「聯絡電話」「條碼」設為純文字 (用 try 包起來)
      try {
        const detailRow = detailSheet.getLastRow();
        forceTextCell(detailSheet, detailRow, ["聯絡電話", "條碼"], {
          聯絡電話: order.phone,
          條碼: String(item.barcode)
        });
      } catch (fmtErr) {
        console.error("明細格式化失敗:", fmtErr);
      }
    });

    // 3. (選擇性) 發送 email 通知 - 如不需要可刪除以下整段
    try {
      sendOrderNotification(order, orderId);
    } catch (mailErr) {
      console.error("Email 通知失敗:", mailErr);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: "success", orderId: orderId })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error("處理訂單失敗:", err);
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * 提供 GET 請求。
 *  - 不帶參數：回傳系統運作狀態(部署測試用)。
 *  - ?action=stats&key=查詢金鑰：回傳該推廣人的業績彙總(給查詢小網頁用)。
 *
 * 回傳的內容「只有彙總數字」(單數、業績、分潤)，不含任何客人個資，
 * 因此即使金鑰外流也不會洩漏客戶名單。
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) || {};

    if (params.action === "stats") {
      return jsonOut(getSellerStats(params.key));
    }

    // 預設：部署測試
    return jsonOut({
      status: "ok",
      message: "PIF 難民特賣會訂單系統運作中",
      time: new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" })
    });
  } catch (err) {
    return jsonOut({ status: "error", message: err.toString() });
  }
}

/**
 * JSON 輸出 + CORS 標頭(讓查詢小網頁可以跨網域讀取)。
 */
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ───────────────────────────────────────────────
//  推廣人(分潤) 相關
// ───────────────────────────────────────────────

/**
 * 讀取「推廣人名單」工作表，回傳代碼/金鑰對照資料。
 * 回傳物件：{
 *   byCode: { 代碼: {code, name, rate, key, enabled} },
 *   byKey:  { 查詢金鑰: {code, name, rate, key, enabled} }
 * }
 * 欄位順序固定：推廣代碼 | 姓名 | 分潤% | 查詢金鑰 | 啟用 | 備註
 */
function loadSellers(ss) {
  const result = { byCode: {}, byKey: {} };
  const sheet = ss.getSheetByName(SELLER_SHEET);
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  // 第 0 列是標題，從第 1 列開始
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = String(row[0] || "").trim();
    if (!code) continue;
    const name = String(row[1] || "").trim();
    let rate = row[2];
    // 分潤% 容錯：可填 10 或 0.1，統一轉成 0~1 的小數
    rate = parseFloat(rate);
    if (isNaN(rate)) rate = 0.1;
    if (rate > 1) rate = rate / 100;
    const key = String(row[3] || "").trim();
    const enabledRaw = String(row[4]).toUpperCase().trim();
    const enabled = !(enabledRaw === "FALSE" || enabledRaw === "0" || enabledRaw === "否");

    const rec = { code: code, name: name, rate: rate, key: key, enabled: enabled };
    result.byCode[code] = rec;
    if (key) result.byKey[key] = rec;
  }
  return result;
}

/**
 * 驗證下單時帶進來的推廣代碼是否在白名單且啟用中。
 * 有效 → 回傳該代碼；無效/空白/停用 → 回傳 ""(視為直營訂單)。
 */
function validateRef(ss, ref) {
  const code = String(ref || "").trim();
  if (!code) return "";
  try {
    const sellers = loadSellers(ss);
    const rec = sellers.byCode[code];
    if (rec && rec.enabled) return rec.code;
  } catch (err) {
    console.error("驗證推廣代碼失敗:", err);
  }
  return "";
}

/**
 * 依「查詢金鑰」回傳該推廣人的業績彙總。
 * 從「訂單總覽」即時加總，所以你手動修改某筆未稅金額後，這裡的數字會跟著更新。
 */
/**
 * 姓名遮罩：保留第一個字與最後一個字，中間以 * 取代(保護客人隱私)。
 *   王小明  → 王*明
 *   陳大文明 → 陳**明
 *   王明    → 王*   (2 字：遮最後一字)
 *   王      → 王     (1 字：原樣)
 * 使用 Array.from 以正確處理多位元組字元。
 */
function maskName(name) {
  const s = String(name || "").trim();
  const chars = Array.from(s);
  const n = chars.length;
  if (n <= 1) return s;
  if (n === 2) return chars[0] + "*";
  return chars[0] + "*".repeat(n - 2) + chars[n - 1];
}

/**
 * 訂單時間格式化為「2026/6/18 下午 1:53:53」。
 * 儲存格可能是 Date 物件，也可能已是字串(早期訂單)；
 *  - 是 Date：用台北時區格式化成 年/月/日 上午|下午 時:分:秒(12 小時制)
 *  - 是字串：原樣回傳(這些本來就已是想要的格式)
 */
function formatOrderTime(val) {
  if (val instanceof Date) {
    var tz = "Asia/Taipei";
    var y = Utilities.formatDate(val, tz, "yyyy");
    var mo = Number(Utilities.formatDate(val, tz, "MM"));
    var d = Number(Utilities.formatDate(val, tz, "dd"));
    var h24 = Number(Utilities.formatDate(val, tz, "HH"));
    var mi = Utilities.formatDate(val, tz, "mm");
    var se = Utilities.formatDate(val, tz, "ss");
    var ampm = h24 < 12 ? "上午" : "下午";
    var h12 = h24 % 12;
    if (h12 === 0) h12 = 12;
    return y + "/" + mo + "/" + d + " " + ampm + " " + h12 + ":" + mi + ":" + se;
  }
  return String(val || "");
}

function getSellerStats(queryKey) {
  const key = String(queryKey || "").trim();
  if (!key) return { status: "error", message: "缺少查詢金鑰" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sellers = loadSellers(ss);
  const seller = sellers.byKey[key];
  if (!seller) return { status: "error", message: "查詢金鑰不存在" };

  const summarySheet = ss.getSheetByName(SUMMARY_SHEET);
  let orderCount = 0;
  let subtotalSum = 0;
  const orders = []; // 該推廣人的每一筆訂單明細

  if (summarySheet) {
    const data = summarySheet.getDataRange().getValues();
    const header = data[0] || [];
    const idxRef = header.indexOf("推廣人");
    const idxSub = header.indexOf("未稅金額");
    const idxTime = header.indexOf("訂單時間");
    const idxId = header.indexOf("訂單編號");
    const idxName = header.indexOf("姓名");
    // 訂單狀態：優先用標題「狀態」定位；找不到時退回 S 欄(第 19 欄，0-based index 18)
    let idxStatus = header.indexOf("狀態");
    if (idxStatus < 0) idxStatus = 18;

    if (idxRef >= 0 && idxSub >= 0) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idxRef] || "").trim() === seller.code) {
          orderCount++;
          const sub = parseFloat(data[i][idxSub]);
          const subVal = isNaN(sub) ? 0 : sub;
          subtotalSum += subVal;

          // 狀態判定：已出貨 / 已現場取貨 → 已完成；其他或空白 → 處理中
          const rawStatus = String(data[i][idxStatus] || "").trim();
          const done = (rawStatus === "已出貨" || rawStatus === "已現場取貨");

          orders.push({
            time: idxTime >= 0 ? formatOrderTime(data[i][idxTime]) : "",
            orderId: idxId >= 0 ? String(data[i][idxId] || "") : "",
            name: maskName(idxName >= 0 ? String(data[i][idxName] || "") : ""),
            subtotal: Math.round(subVal),
            commission: Math.round(subVal * seller.rate),
            status: done ? "已完成" : "處理中"
          });
        }
      }
    }
  }

  const commission = Math.round(subtotalSum * seller.rate);
  return {
    status: "ok",
    name: seller.name,
    code: seller.code,
    refUrl: SITE_ROOT_URL + "?ref=" + encodeURIComponent(seller.code),
    ratePct: Math.round(seller.rate * 100),
    orderCount: orderCount,
    subtotal: Math.round(subtotalSum),
    commission: commission,
    orders: orders,
    time: new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" })
  };
}


/**
 * 強制把指定欄位的儲存格設為純文字格式並重寫內容
 * 修復 Google Sheets 把開頭 0 的電話/條碼自動吃掉的問題
 *
 * @param sheet      工作表物件
 * @param row        要修改的列號 (1-based)
 * @param colNames   要強制為文字的欄位名稱陣列，例如 ["聯絡電話", "條碼"]
 * @param values     對應的值物件，例如 { 聯絡電話: "0958222911", 條碼: "0123456" }
 */
function forceTextCell(sheet, row, colNames, values) {
  // 只搜尋預期的欄位範圍 (前 20 欄足夠涵蓋訂單欄位)，避免使用者手動擴展工作表導致誤判
  const maxCol = Math.min(sheet.getLastColumn(), 20);
  if (maxCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, maxCol).getValues()[0];
  colNames.forEach(colName => {
    const idx = headers.indexOf(colName);
    if (idx === -1) return;
    try {
      const cell = sheet.getRange(row, idx + 1);
      cell.setNumberFormat("@");
      const v = values[colName];
      if (v !== undefined && v !== null && v !== "") {
        cell.setValue(String(v));
      }
    } catch (e) {
      console.error("forceTextCell 處理 " + colName + " 失敗:", e);
    }
  });
}


/**
 * 取得或建立工作表
 */
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    // 標題列格式
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#d9534f");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
    // 自動調整欄寬
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  } else {
    // 工作表已存在：把缺少的標題補到最後面(例如新版本新增的「推廣人」欄)。
    // 這樣舊試算表升級後，getSellerStats 也找得到「推廣人」欄。
    try {
      const existing = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      headers.forEach(h => {
        if (existing.indexOf(h) === -1) {
          const idx = sheet.getLastColumn() + 1;
          const cell = sheet.getRange(1, idx);
          cell.setValue(h);
          cell.setBackground("#d9534f").setFontColor("#ffffff").setFontWeight("bold");
          existing.push(h); // 同步本地副本，避免同次補兩個欄位時算錯欄號
        }
      });
    } catch (hdrErr) {
      console.error("補標題欄失敗:", hdrErr);
    }
  }
  return sheet;
}


/**
 * 產生訂單編號 (PIF + YYYYMMDD + 4 位流水)
 */
function generateOrderId() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const random = String(Math.floor(Math.random() * 10000)).padStart(4, "0");
  return `PIF${yyyy}${mm}${dd}${random}`;
}


/**
 * (可選) 發送訂單通知 email
 * 收件人：請改成你的 email
 */
function sendOrderNotification(order, orderId) {
  const recipient = "ichewtong.list@gmail.com"; // ← 改成你要接收訂單通知的 email
  const subject = `【PIF 訂單】${orderId} - ${order.name}（${order.totalQty} 件 / NT$${order.totalAmount}）`;

  let body = `收到新訂單！\n\n`;
  body += `訂單編號：${orderId}\n`;
  body += `訂購時間：${order.timestamp}\n`;
  body += `─────────────\n`;
  body += `姓名：${order.name}\n`;
  body += `公司/部門：${order.company}\n`;
  body += `電話：${order.phone}\n`;
  body += `Email：${order.email}\n`;
  body += `取貨方式：${order.delivery}\n`;
  if (order.address) body += `宅配地址：${order.address}\n`;
  body += `付款方式：${order.payment}\n`;
  body += `希望取貨日：${order.pickupDate}\n`;
  body += `備註：${order.note}\n`;
  body += `─────────────\n`;
  body += `購買品項：\n`;
  order.items.forEach(it => {
    const specText = it.spec ? ` (${it.spec})` : "";
    // 贈品判斷：isGift 旗標，或未稅單價/小計為 0（涵蓋滿額贈與福袋大禮包全部 12 樣）
    const isGiftRow = it.isGift === true || (parseInt(it.price) || 0) === 0;
    if (isGiftRow) {
      body += `  ・🎁[贈品] ${it.name}${specText} × ${it.qty} = 免費\n`;
    } else {
      body += `  ・${it.name}${specText} × ${it.qty} = NT$${it.subtotal}\n`;
    }
  });
  body += `─────────────\n`;
  body += `總數量：${order.totalQty} 件\n`;
  body += `未稅金額：NT$ ${order.subtotalAmount}\n`;
  body += `營業稅 5%：NT$ ${order.taxAmount}\n`;
  body += `應付總金額：NT$ ${order.totalAmount}（運費另計）\n`;

  MailApp.sendEmail(recipient, subject, body);
}


// ============================================================
// 📋 帳單預覽工具 - 用於團主對帳/寄送帳單給客人
// ============================================================

/**
 * 在試算表打開時自動建立選單
 * (你打開試算表後上方會出現「📋 帳單工具」選單)
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 帳單工具")
    .addItem("產生帳單預覽（選中那筆訂單）", "showInvoicePreview")
    .addItem("📦 匯出出貨單 PDF（選中那筆訂單）", "exportShippingLabel")
    .addSeparator()
    .addItem("一鍵設定（首次使用）", "setupSheet")
    .addItem("🔧 修復舊訂單電話/條碼（補回開頭 0）", "fixPhoneAndBarcode")
    .addSeparator()
    .addItem("💰 建立/更新 分潤系統工作表", "setupCommissionSheets")
    .addToUi();
}


/**
 * 一鍵建立分潤系統需要的兩張工作表：
 *   1.「推廣人名單」— 你在這裡新增推廣人(白名單)
 *   2.「分潤結算」 — 自動用公式算出每個推廣人的業績與分潤
 * 可重複執行：已存在就不會覆蓋你既有的資料，只補上缺的標題/公式。
 */
function setupCommissionSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 1. 推廣人名單
  let sellerSheet = ss.getSheetByName(SELLER_SHEET);
  if (!sellerSheet) {
    sellerSheet = ss.insertSheet(SELLER_SHEET);
    sellerSheet.appendRow(["推廣代碼", "姓名", "分潤%", "查詢金鑰", "啟用", "備註"]);
    sellerSheet.getRange("1:1").setFontWeight("bold");
    // 範例列(可刪)
    sellerSheet.appendRow(["alice", "王小美", 0.1, "alice-7fK2", true, "範例，可刪除"]);
    // 把「查詢金鑰」「推廣代碼」設成純文字，避免被當公式或數字
    sellerSheet.getRange("A:A").setNumberFormat("@");
    sellerSheet.getRange("D:D").setNumberFormat("@");
    sellerSheet.getRange("C:C").setNumberFormat("0%");
    sellerSheet.setFrozenRows(1);
  }

  // G 欄：自動產生「推廣網址」(無論名單是新建或既有都會補上)。
  // 只要 A 欄有推廣代碼，G 欄就自動組出 根網址?ref=代碼，新增推廣人免手動拼網址。
  sellerSheet.getRange("G1").setValue("推廣網址");
  sellerSheet.getRange("G1").setFontWeight("bold");
  sellerSheet.getRange("G2").setFormula(
    '=ARRAYFORMULA(IF(A2:A="","","' + SITE_ROOT_URL + '?ref="&A2:A))'
  );
  sellerSheet.getRange("G:G").setNumberFormat("@"); // 純文字，避免被當連結縮短或公式

  // H 欄：自動產生「查成績連結」。用 D 欄查詢金鑰組出 根網址seller.html?key=金鑰，
  // 私下提供給該推廣人查自己的業績/分潤(請勿公開)。
  sellerSheet.getRange("H1").setValue("查成績連結");
  sellerSheet.getRange("H1").setFontWeight("bold");
  sellerSheet.getRange("H2").setFormula(
    '=ARRAYFORMULA(IF(D2:D="","","' + SITE_ROOT_URL + 'seller.html?key="&D2:D))'
  );
  sellerSheet.getRange("H:H").setNumberFormat("@");

  // 2. 分潤結算(用公式即時從訂單總覽加總；改訂單金額會自動更新)
  let calcSheet = ss.getSheetByName("分潤結算");
  if (!calcSheet) {
    calcSheet = ss.insertSheet("分潤結算");
  }
  calcSheet.clearContents();
  calcSheet.appendRow(["推廣代碼", "姓名", "分潤%", "成交訂單數", "推廣業績(未稅)", "預估分潤"]);
  calcSheet.getRange("1:1").setFontWeight("bold");
  calcSheet.setFrozenRows(1);

  // 從第 2 列起，對應「推廣人名單」每一位推廣人帶出公式。
  // - A/B/C：用 FILTER + VLOOKUP 帶出代碼、姓名、分潤%(會隨名單自動增減列)。
  // - D/E：成交訂單數、業績。改用 BYROW + SUMPRODUCT，因為 SUMIF/COUNTIF 不吃
  //        INDEX 取出的「陣列」(會報「引數必須是一個範圍」)，SUMPRODUCT 則可以吃陣列。
  // - 欄位用 MATCH 動態定位「推廣人」「未稅金額」在訂單總覽的第幾欄，所以你日後
  //   移動欄位或新增自訂欄都不會壞。
  const refCol =
    'INDEX(訂單總覽!$A:$ZZ,0,MATCH("推廣人",訂單總覽!$1:$1,0))';
  const subCol =
    'INDEX(訂單總覽!$A:$ZZ,0,MATCH("未稅金額",訂單總覽!$1:$1,0))';

  calcSheet.getRange("A2").setFormula(
    '=IFERROR(FILTER(' + SELLER_SHEET + '!A2:A, ' + SELLER_SHEET + '!A2:A<>""),"")'
  );
  calcSheet.getRange("B2").setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,' + SELLER_SHEET + '!A:B,2,FALSE),"")))'
  );
  calcSheet.getRange("C2").setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(VLOOKUP(A2:A,' + SELLER_SHEET + '!A:C,3,FALSE),"")))'
  );
  // 成交訂單數：該推廣代碼在訂單總覽「推廣人」欄出現幾次
  calcSheet.getRange("D2").setFormula(
    '=BYROW(A2:A, LAMBDA(code, IF(code="","", SUMPRODUCT((' + refCol + '=code)*1))))'
  );
  // 推廣業績(未稅)：把該代碼對應的「未稅金額」加總
  calcSheet.getRange("E2").setFormula(
    '=BYROW(A2:A, LAMBDA(code, IF(code="","", SUMPRODUCT((' + refCol + '=code)*N(' + subCol + ')))))'
  );
  // 預估分潤：業績 × 分潤%
  calcSheet.getRange("F2").setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",ROUND(E2:E*C2:C)))'
  );
  calcSheet.getRange("C:C").setNumberFormat("0%");
  calcSheet.getRange("E:E").setNumberFormat("$#,##0");
  calcSheet.getRange("F:F").setNumberFormat("$#,##0");

  ui.alert(
    "✅ 分潤系統工作表已建立/更新完成！\n\n" +
    "1.「推廣人名單」：在這裡新增推廣人(一列一位)。\n" +
    "   - 推廣代碼：給客人的網址 ?ref= 值(好記，如 alice)\n" +
    "   - 分潤%：可填 10 或 0.1\n" +
    "   - 查詢金鑰：給推廣人查成績的密碼(不好猜)\n" +
    "   - 啟用：打 TRUE 才生效\n\n" +
    "2.「分潤結算」：自動算好每人的業績與分潤，免手動。\n" +
    "   你修改訂單總覽的未稅金額後，這裡會自動重算。"
  );
}


/**
 * 修復舊訂單的電話與條碼：
 * - 把「聯絡電話」「條碼」欄位設為純文字格式
 * - 對於失去開頭 0 的電話（9 碼且開頭非 0），自動補回 0
 * - 對於失去開頭 0 的條碼（少於 13 碼），自動補零到 13 碼
 */
function fixPhoneAndBarcode() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalFixed = 0;
  let totalFormatted = 0;
  const reportLines = [];

  const sheetsToFix = [
    { name: SUMMARY_SHEET, cols: ["聯絡電話"] },
    { name: DETAIL_SHEET, cols: ["聯絡電話", "條碼"] }
  ];

  sheetsToFix.forEach(cfg => {
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      reportLines.push(`⚠️ 找不到工作表「${cfg.name}」，已跳過`);
      return;
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    cfg.cols.forEach(colName => {
      const idx = headers.indexOf(colName);
      if (idx === -1) return;
      const colNum = idx + 1;

      // 先把整欄設為純文字格式
      const fullRange = sheet.getRange(2, colNum, lastRow - 1, 1);
      fullRange.setNumberFormat("@");
      totalFormatted += (lastRow - 1);

      // 讀取所有值並嘗試補零
      const values = fullRange.getValues();
      let changedInCol = 0;
      const newValues = values.map(r => {
        let v = r[0];
        if (v === "" || v === null || v === undefined) return [""];
        v = String(v).trim();

        if (colName === "聯絡電話") {
          // 移除非數字字元做檢查 (但保留原本的 - 等格式)
          const digits = v.replace(/\D/g, "");
          // 台灣手機 9 碼 (例如 958222911) → 補 0 變 0958222911
          if (digits.length === 9 && !v.startsWith("0")) {
            const fixed = "0" + v;
            changedInCol++;
            return [fixed];
          }
          // 台灣市話 9 碼 (例如 282838333) → 補 0 變 0282838333
          // 已包含在上方判斷中
        }

        if (colName === "條碼") {
          const digits = v.replace(/\D/g, "");
          // 常見條碼長度 8/12/13/14；若少 1 碼通常是開頭 0 被吃掉
          if (digits.length === 12 && !v.startsWith("0")) {
            const fixed = "0" + v;
            changedInCol++;
            return [fixed];
          }
          if (digits.length === 7 && !v.startsWith("0")) {
            const fixed = "0" + v;
            changedInCol++;
            return [fixed];
          }
        }

        return [v];
      });

      // 寫回試算表
      fullRange.setValues(newValues);
      if (changedInCol > 0) {
        totalFixed += changedInCol;
        reportLines.push(`✓ ${cfg.name} / ${colName}：補回 ${changedInCol} 筆開頭 0`);
      }
    });
  });

  const message = [
    `✅ 修復完成！`,
    ``,
    `• 已將電話/條碼欄位設為純文字格式：${totalFormatted} 個儲存格`,
    `• 已自動補回開頭 0：${totalFixed} 筆`,
    ``
  ].concat(reportLines).concat([
    ``,
    `※ 之後新進的訂單會自動套用純文字格式，不會再有開頭 0 消失的問題。`
  ]).join("\n");

  ui.alert(message);
}

/**
 * 首次使用：在訂單總覽自動補上「運費」「截止日」「狀態」欄位
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SUMMARY_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("找不到「訂單總覽」工作表，請先確認有客人下過至少一筆訂單。");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newCols = [
    { name: "運費(元)", default: "" },
    { name: "繳款截止日", default: "" },
    { name: "狀態", default: "待確認" }
  ];

  let added = 0;
  newCols.forEach(col => {
    if (headers.indexOf(col.name) === -1) {
      const newColIdx = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColIdx).setValue(col.name);
      sheet.getRange(1, newColIdx)
        .setBackground("#d9534f").setFontColor("#ffffff").setFontWeight("bold");
      added++;
    }
  });

  // 狀態欄加入下拉選單
  const allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusColIdx = allHeaders.indexOf("狀態") + 1;
  if (statusColIdx > 0 && sheet.getLastRow() > 1) {
    const range = sheet.getRange(2, statusColIdx, sheet.getLastRow() - 1, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["待確認", "已發帳單", "已匯款", "已出貨", "取消"])
      .setAllowInvalid(true)
      .build();
    range.setDataValidation(rule);
  }

  SpreadsheetApp.getUi().alert(`✅ 設定完成！新增 ${added} 個欄位。\n\n下次更新時可重複執行不會重複新增。`);
}

/**
 * 主要功能：顯示帳單預覽
 * 流程：點選任一訂單列 → 執行 → 跳出對話框輸入運費 → 顯示完整帳單可複製
 */
function showInvoicePreview() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET);
  const detailSheet = ss.getSheetByName(DETAIL_SHEET);

  if (!summarySheet || !detailSheet) {
    ui.alert("找不到工作表，請確認「訂單總覽」與「訂單明細」存在。");
    return;
  }

  // 1. 取得使用者選的列
  const activeSheet = SpreadsheetApp.getActiveSheet();
  if (activeSheet.getName() !== SUMMARY_SHEET) {
    ui.alert("請先切換到「訂單總覽」工作表，並點擊任一筆訂單的任何儲存格後再執行。");
    return;
  }

  const activeRow = activeSheet.getActiveRange().getRow();
  if (activeRow < 2) {
    ui.alert("請點選一筆訂單列後再執行（不能是標題列）。");
    return;
  }

  // 2. 取得該筆訂單資料
  const headers = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).getValues()[0];
  const rowData = summarySheet.getRange(activeRow, 1, 1, summarySheet.getLastColumn()).getValues()[0];
  const order = {};
  headers.forEach((h, i) => { order[h] = rowData[i]; });

  // 3. 取得對應的商品明細
  const detailHeaders = detailSheet.getRange(1, 1, 1, detailSheet.getLastColumn()).getValues()[0];
  const detailData = detailSheet.getRange(2, 1, detailSheet.getLastRow() - 1, detailSheet.getLastColumn()).getValues();
  const orderIdIdx = detailHeaders.indexOf("訂單編號");

  const orderItems = detailData
    .filter(r => r[orderIdIdx] === order["訂單編號"])
    .map(r => {
      const item = {};
      detailHeaders.forEach((h, i) => { item[h] = r[i]; });
      return item;
    });

  if (orderItems.length === 0) {
    ui.alert("找不到此訂單的商品明細，請確認訂單編號正確。");
    return;
  }

  // 4. 詢問運費（如果還沒填）
  let shippingFee = order["運費(元)"];
  if (shippingFee === "" || shippingFee === undefined || shippingFee === null) {
    const isPickup = String(order["取貨方式"]).indexOf("自取") >= 0;
    const defaultFee = isPickup ? "0" : "130";
    const resp = ui.prompt(
      "輸入運費",
      `請輸入運費金額（${isPickup ? "客人選擇自取，預設 0" : "客人選擇宅配，每箱 130 元"}）：\n\n例：130 或 260 或 0`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    shippingFee = parseInt(resp.getResponseText()) || 0;
    // 寫回試算表
    const colIdx = headers.indexOf("運費(元)") + 1;
    if (colIdx > 0) {
      summarySheet.getRange(activeRow, colIdx).setValue(shippingFee);
    }
  }
  shippingFee = parseInt(shippingFee) || 0;

  // 5. 詢問繳款截止日（如果還沒填）
  let dueDate = order["繳款截止日"];
  if (!dueDate || dueDate === "") {
    const today = new Date();
    const defaultDue = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 5);
    const defaultStr = Utilities.formatDate(defaultDue, "Asia/Taipei", "yyyy/MM/dd");
    const resp = ui.prompt(
      "輸入繳款截止日",
      `請輸入繳款截止日期（建議從今天起 3~5 天，預設 ${defaultStr}）：\n\n例：2026/05/20`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    dueDate = resp.getResponseText().trim() || defaultStr;
    const colIdx = headers.indexOf("繳款截止日") + 1;
    if (colIdx > 0) {
      summarySheet.getRange(activeRow, colIdx).setValue(dueDate);
    }
  } else if (dueDate instanceof Date) {
    dueDate = Utilities.formatDate(dueDate, "Asia/Taipei", "yyyy/MM/dd");
  }

  // 6. 計算金額
  const subtotal = parseInt(order["未稅金額"]) || 0;
  const tax = parseInt(order["營業稅 5%"]) || 0;
  const goodsTotal = parseInt(order["應付總金額(未含運費)"]) || (subtotal + tax);
  const grandTotal = goodsTotal + shippingFee;

  // 7. 產生帳單 HTML
  const customerName = order["姓名"] || "";
  const company = order["公司/部門"] || "";
  const phone = order["聯絡電話"] || "";
  const email = order["Email"] || "";
  const delivery = order["取貨方式"] || "";
  const address = order["宅配地址"] || "";
  const orderId = order["訂單編號"] || "";
  const orderTime = order["訂單時間"] || "";
  const customerNote = order["備註"] || "";

  let itemsHtml = "";
  orderItems.forEach((it, i) => {
    const specText = it["規格"] ? `<br><span style="color:#888;font-size:12px;">${it["規格"]}</span>` : "";
    // 識別贈品：未稅單價為 0 即視為贈品（涵蓋滿額贈與福袋大禮包全部 12 樣）
    const isGiftRow = (parseInt(it["未稅單價"]) || 0) === 0;
    const rowBg = isGiftRow ? "background:#fff8dc;" : "";
    const giftBadge = isGiftRow ? '<span style="background:#d9534f;color:#fff;font-size:10px;padding:1px 6px;border-radius:8px;margin-right:6px;font-weight:700;">🎁 贈品</span>' : '';
    const priceText = isGiftRow ? '<span style="color:#28a745;">免費</span>' : it["未稅單價"];
    const subtotalText = isGiftRow ? '<span style="color:#28a745;">免費</span>' : parseInt(it["未稅小計"]).toLocaleString();
    itemsHtml += `
      <tr style="${rowBg}">
        <td style="padding:8px 6px;border-bottom:1px solid #eee;text-align:center;color:#888;">${i + 1}</td>
        <td style="padding:8px 6px;border-bottom:1px solid #eee;">${giftBadge}${it["商品名稱"]}${specText}</td>
        <td style="padding:8px 6px;border-bottom:1px solid #eee;text-align:right;">${priceText}</td>
        <td style="padding:8px 6px;border-bottom:1px solid #eee;text-align:center;">${it["數量"]}</td>
        <td style="padding:8px 6px;border-bottom:1px solid #eee;text-align:right;font-weight:600;">${subtotalText}</td>
      </tr>`;
  });

  const invoiceHtml = `
<style>
  body { font-family: "Noto Sans TC","Microsoft JhengHei",Arial,sans-serif; padding:0; margin:0; color:#333; }
  .wrap { max-width: 720px; margin: 0 auto; padding: 20px; }
  .toolbar { display:flex; gap:8px; margin-bottom:16px; padding-bottom:12px; border-bottom:1px solid #eee; }
  .toolbar button { padding:8px 16px; border:1px solid #ddd; background:#fff; cursor:pointer; border-radius:4px; font-size:13px; }
  .toolbar button.primary { background:#d9534f; color:#fff; border-color:#d9534f; font-weight:600; }
  .toolbar button:hover { background:#f5f5f5; }
  .toolbar button.primary:hover { background:#c9302c; }
  pre { white-space: pre-wrap; word-wrap: break-word; font-family: inherit; font-size: 14px; line-height: 1.7; background:#fafafa; padding: 16px; border-radius:6px; border:1px solid #eee; max-height:480px; overflow:auto; }
  h2 { font-size:16px; color:#d9534f; margin:0 0 10px; }
  .hint { font-size:12px; color:#888; margin-top:8px; }
</style>
<div class="wrap">
  <div class="toolbar">
    <button class="primary" onclick="copyText()">📋 一鍵複製全文</button>
    <button onclick="copyHtml()">複製為 HTML（保留格式貼到 Gmail）</button>
    <button onclick="google.script.host.close()">關閉</button>
  </div>
  <h2>📧 帳單純文字版（可貼到任何 email）</h2>
  <pre id="plain">${buildPlainTextInvoice({
    orderId, orderTime, customerName, company, phone, email,
    delivery, address, items: orderItems,
    subtotal, tax, goodsTotal, shippingFee, grandTotal,
    dueDate, customerNote
  })}</pre>
  <div class="hint">提示：「一鍵複製全文」會把上面文字複製到剪貼簿，你直接到 Gmail 貼上即可。</div>

  <h2 style="margin-top:24px;">📄 HTML 格式版（保留表格樣式）</h2>
  <div id="rich" style="border:1px solid #eee; padding:16px; border-radius:6px; background:#fff;">
    ${buildHtmlInvoice({
      orderId, customerName, company, phone, items: orderItems, itemsHtml,
      subtotal, tax, goodsTotal, shippingFee, grandTotal,
      dueDate, customerNote, delivery, address
    })}
  </div>

  <script>
    function copyText() {
      const text = document.getElementById("plain").innerText;
      navigator.clipboard.writeText(text).then(() => {
        alert("✅ 已複製！可以貼到 Gmail / Outlook 了");
      }).catch(() => {
        // 後備方案
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        alert("✅ 已複製！");
      });
    }
    function copyHtml() {
      const rich = document.getElementById("rich");
      const range = document.createRange();
      range.selectNode(rich);
      window.getSelection().removeAllRanges();
      window.getSelection().addRange(range);
      document.execCommand("copy");
      window.getSelection().removeAllRanges();
      alert("✅ 已複製 HTML 格式！直接到 Gmail 貼上會保留表格樣式");
    }
  </script>
</div>`;

  const htmlOutput = HtmlService.createHtmlOutput(invoiceHtml)
    .setWidth(800).setHeight(720)
    .setTitle("帳單預覽 - " + orderId);
  ui.showModalDialog(htmlOutput, `帳單預覽 - ${customerName}（${orderId}）`);

  // 更新狀態為「已發帳單」
  const statusColIdx = headers.indexOf("狀態") + 1;
  if (statusColIdx > 0) {
    const currentStatus = summarySheet.getRange(activeRow, statusColIdx).getValue();
    if (currentStatus === "待確認" || currentStatus === "") {
      summarySheet.getRange(activeRow, statusColIdx).setValue("已發帳單");
    }
  }
}

/**
 * 產生純文字版帳單
 */
function buildPlainTextInvoice(data) {
  const line = "═══════════════════════════════";
  const sub = "───────────────────────────────";
  let txt = "";
  txt += `${line}\n`;
  txt += `        PIF 難民特賣會 - 訂單帳單\n`;
  txt += `${line}\n\n`;
  txt += `親愛的 ${data.customerName} 您好，\n\n`;
  txt += `感謝您訂購 PIF 難民特賣會商品！\n`;
  txt += `以下為您本次訂單的帳單明細，請於繳款截止日前完成匯款。\n\n`;
  txt += `📋 訂單資訊\n`;
  txt += `${sub}\n`;
  txt += `訂單編號：${data.orderId}\n`;
  txt += `訂購時間：${data.orderTime}\n`;
  txt += `訂購人　：${data.customerName}` + (data.company ? `（${data.company}）` : "") + `\n`;
  txt += `聯絡電話：${data.phone}\n`;
  if (data.email) txt += `Email　 ：${data.email}\n`;
  txt += `取貨方式：${data.delivery}\n`;
  if (data.address) txt += `宅配地址：${data.address}\n`;
  txt += `\n`;
  txt += `🛍️ 商品明細\n`;
  txt += `${sub}\n`;
  data.items.forEach((it, i) => {
    // 贈品判斷：未稅單價為 0（涵蓋滿額贈與福袋大禮包全部 12 樣）
    const isGiftRow = (parseInt(it["未稅單價"]) || 0) === 0;
    txt += `${String(i + 1).padStart(2, " ")}. ${isGiftRow ? "🎁[贈品] " : ""}${it["商品名稱"]}\n`;
    if (it["規格"]) txt += `    規格：${it["規格"]}\n`;
    if (isGiftRow) {
      txt += `    免費贈送 × ${it["數量"]}\n`;
    } else {
      txt += `    NT$ ${it["未稅單價"]} × ${it["數量"]} = NT$ ${parseInt(it["未稅小計"]).toLocaleString()}\n`;
    }
  });
  txt += `\n`;
  txt += `💰 金額明細\n`;
  txt += `${sub}\n`;
  txt += `商品未稅金額　 NT$ ${parseInt(data.subtotal).toLocaleString()}\n`;
  txt += `營業稅 5%　　　NT$ ${parseInt(data.tax).toLocaleString()}\n`;
  txt += `商品合計　　　 NT$ ${parseInt(data.goodsTotal).toLocaleString()}\n`;
  txt += `運費　　　　　 NT$ ${parseInt(data.shippingFee).toLocaleString()}\n`;
  txt += `${sub}\n`;
  txt += `★ 應付總金額　 NT$ ${parseInt(data.grandTotal).toLocaleString()}\n\n`;
  txt += `🏦 匯款資訊\n`;
  txt += `${sub}\n`;
  txt += `銀行：永豐銀行（807）延平分行\n`;
  txt += `戶名：毓秀堂貿易有限公司\n`;
  txt += `帳號：(807) 10900100312867\n\n`;
  txt += `⏰ 繳款截止日：${data.dueDate}\n\n`;
  if (data.customerNote) {
    txt += `📝 您的備註\n`;
    txt += `${sub}\n`;
    txt += `${data.customerNote}\n\n`;
  }
  txt += `※ 注意事項\n`;
  txt += `${sub}\n`;
  txt += `1. 完成匯款後，請回覆此 email 並提供「匯款帳號末 5 碼」以加速對帳\n`;
  txt += `2. 若於截止日前未完成匯款，訂單將自動取消\n`;
  txt += `3. 對帳完成後，我們將盡速為您安排出貨\n\n`;
  txt += `如有任何問題，歡迎來電或回信詢問。\n\n`;
  txt += `毓秀堂貿易有限公司\n`;
  txt += `📧 ichewtong.list@gmail.com\n`;
  txt += `☎️ 02-8283-8333\n`;
  txt += `${line}\n`;
  return txt;
}

/**
 * 產生 HTML 格式版帳單
 */
function buildHtmlInvoice(data) {
  return `
<div style="font-family:'Noto Sans TC','Microsoft JhengHei',Arial,sans-serif;color:#333;max-width:680px;">
  <div style="text-align:center;padding:20px 0;border-bottom:3px solid #d9534f;margin-bottom:24px;">
    <h1 style="font-size:22px;color:#d9534f;margin:0;letter-spacing:2px;">PIF 難民特賣會</h1>
    <p style="font-size:14px;color:#888;margin:8px 0 0;">訂單帳單</p>
  </div>

  <p style="font-size:15px;line-height:1.8;">親愛的 <strong>${data.customerName}</strong> 您好，<br>感謝您訂購 PIF 難民特賣會商品！以下為您本次訂單的帳單明細：</p>

  <h3 style="font-size:15px;color:#d9534f;border-left:4px solid #d9534f;padding-left:10px;margin:24px 0 12px;">📋 訂單資訊</h3>
  <table style="width:100%;font-size:14px;line-height:1.8;border-collapse:collapse;">
    <tr><td style="width:100px;color:#888;">訂單編號</td><td><strong>${data.orderId}</strong></td></tr>
    <tr><td style="color:#888;">訂購人</td><td>${data.customerName}${data.company ? "（" + data.company + "）" : ""}</td></tr>
    <tr><td style="color:#888;">聯絡電話</td><td>${data.phone}</td></tr>
    <tr><td style="color:#888;">取貨方式</td><td>${data.delivery}</td></tr>
    ${data.address ? `<tr><td style="color:#888;">宅配地址</td><td>${data.address}</td></tr>` : ""}
  </table>

  <h3 style="font-size:15px;color:#d9534f;border-left:4px solid #d9534f;padding-left:10px;margin:24px 0 12px;">🛍️ 商品明細</h3>
  <table style="width:100%;font-size:13px;border-collapse:collapse;">
    <thead>
      <tr style="background:#fff5f5;">
        <th style="padding:10px 6px;text-align:center;color:#d9534f;border-bottom:2px solid #d9534f;width:40px;">#</th>
        <th style="padding:10px 6px;text-align:left;color:#d9534f;border-bottom:2px solid #d9534f;">品名</th>
        <th style="padding:10px 6px;text-align:right;color:#d9534f;border-bottom:2px solid #d9534f;width:80px;">未稅單價</th>
        <th style="padding:10px 6px;text-align:center;color:#d9534f;border-bottom:2px solid #d9534f;width:60px;">數量</th>
        <th style="padding:10px 6px;text-align:right;color:#d9534f;border-bottom:2px solid #d9534f;width:90px;">未稅小計</th>
      </tr>
    </thead>
    <tbody>${data.itemsHtml}</tbody>
  </table>

  <h3 style="font-size:15px;color:#d9534f;border-left:4px solid #d9534f;padding-left:10px;margin:24px 0 12px;">💰 金額明細</h3>
  <table style="width:100%;font-size:14px;line-height:1.9;border-collapse:collapse;">
    <tr><td style="color:#888;">商品未稅金額</td><td style="text-align:right;">NT$ ${parseInt(data.subtotal).toLocaleString()}</td></tr>
    <tr><td style="color:#888;">營業稅 5%</td><td style="text-align:right;">NT$ ${parseInt(data.tax).toLocaleString()}</td></tr>
    <tr><td style="color:#888;border-bottom:1px dashed #ddd;padding-bottom:6px;">商品合計</td><td style="text-align:right;border-bottom:1px dashed #ddd;">NT$ ${parseInt(data.goodsTotal).toLocaleString()}</td></tr>
    <tr><td style="color:#888;">運費</td><td style="text-align:right;">NT$ ${parseInt(data.shippingFee).toLocaleString()}</td></tr>
    <tr style="background:#fff5f5;"><td style="padding:10px 8px;font-size:16px;font-weight:700;">★ 應付總金額</td><td style="text-align:right;padding:10px 8px;font-size:20px;color:#d9534f;font-weight:700;">NT$ ${parseInt(data.grandTotal).toLocaleString()}</td></tr>
  </table>

  <div style="background:#fff8dc;border:2px solid #d4a017;border-radius:6px;padding:16px;margin:24px 0;">
    <h3 style="font-size:15px;color:#8b5a00;margin:0 0 10px;">⏰ 繳款截止日：<span style="color:#d9534f;">${data.dueDate}</span></h3>
    <p style="margin:0;font-size:13px;color:#6b4500;">若於截止日前未完成匯款，訂單將自動取消，敬請留意。</p>
  </div>

  <h3 style="font-size:15px;color:#d9534f;border-left:4px solid #d9534f;padding-left:10px;margin:24px 0 12px;">🏦 匯款資訊</h3>
  <table style="width:100%;font-size:14px;line-height:1.9;background:#fafafa;padding:12px;border-radius:6px;">
    <tr><td style="width:80px;color:#888;padding-left:12px;">銀行</td><td>永豐銀行（807）延平分行</td></tr>
    <tr><td style="color:#888;padding-left:12px;">戶名</td><td>毓秀堂貿易有限公司</td></tr>
    <tr><td style="color:#888;padding-left:12px;">帳號</td><td><strong>(807) 10900100312867</strong></td></tr>
  </table>

  ${data.customerNote ? `
  <h3 style="font-size:15px;color:#d9534f;border-left:4px solid #d9534f;padding-left:10px;margin:24px 0 12px;">📝 您的備註</h3>
  <p style="background:#fafafa;padding:12px 16px;border-radius:6px;font-size:14px;color:#555;">${data.customerNote}</p>` : ""}

  <div style="margin-top:30px;padding:16px;background:#f9f9f9;border-radius:6px;font-size:13px;color:#666;line-height:1.8;">
    <strong style="color:#333;">※ 注意事項</strong><br>
    1. 完成匯款後，請回覆此 email 並提供「匯款帳號末 5 碼」以加速對帳<br>
    2. 若於截止日前未完成匯款，訂單將自動取消<br>
    3. 對帳完成後，我們將盡速為您安排出貨
  </div>

  <div style="text-align:center;margin-top:30px;padding-top:20px;border-top:1px solid #eee;font-size:13px;color:#999;line-height:1.8;">
    毓秀堂貿易有限公司<br>
    📧 ichewtong.list@gmail.com　 ☎️ 02-8283-8333
  </div>
</div>`;
}


// ============================================================
// 📦 出貨單 PDF 匯出工具
// ============================================================

/**
 * 匯出出貨單 PDF
 * 流程：選中訂單列 → 抓資料(含 V欄官網訂單號碼) → 生成 HTML → 轉 PDF → 跳下載連結
 */
function exportShippingLabel() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET);
  const detailSheet = ss.getSheetByName(DETAIL_SHEET);

  if (!summarySheet || !detailSheet) {
    ui.alert("找不到工作表，請確認「訂單總覽」與「訂單明細」存在。");
    return;
  }

  const activeSheet = SpreadsheetApp.getActiveSheet();
  if (activeSheet.getName() !== SUMMARY_SHEET) {
    ui.alert("請先切換到「訂單總覽」工作表，並點擊任一筆訂單的任何儲存格後再執行。");
    return;
  }

  const activeRow = activeSheet.getActiveRange().getRow();
  if (activeRow < 2) {
    ui.alert("請點選一筆訂單列後再執行（不能是標題列）。");
    return;
  }

  // 取得該筆訂單資料 (前面欄位用 header 對應，V 欄手動指定)
  const lastCol = Math.max(summarySheet.getLastColumn(), 22);
  const headers = summarySheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowData = summarySheet.getRange(activeRow, 1, 1, lastCol).getValues()[0];
  const order = {};
  headers.forEach((h, i) => { if (h) order[h] = rowData[i]; });

  // V 欄 (第 22 欄) 為官網訂單號碼 (使用者手動填的)
  const webOrderNo = rowData[21] || "";

  // 取得商品明細
  const detailHeaders = detailSheet.getRange(1, 1, 1, detailSheet.getLastColumn()).getValues()[0];
  const detailLastRow = detailSheet.getLastRow();
  if (detailLastRow < 2) {
    ui.alert("訂單明細工作表沒有資料。");
    return;
  }
  const detailData = detailSheet.getRange(2, 1, detailLastRow - 1, detailSheet.getLastColumn()).getValues();
  const orderIdIdx = detailHeaders.indexOf("訂單編號");

  const orderItems = detailData
    .filter(r => r[orderIdIdx] === order["訂單編號"])
    .map(r => {
      const item = {};
      detailHeaders.forEach((h, i) => { item[h] = r[i]; });
      return item;
    });

  if (orderItems.length === 0) {
    ui.alert("找不到此訂單的商品明細。");
    return;
  }

  // 取貨方式 (判斷是否需要顯示地址)
  const delivery = String(order["取貨方式"] || "");
  const isPickup = delivery.indexOf("自取") >= 0;
  const address = order["宅配地址"] || "";

  // 計算金額
  const subtotal = parseInt(order["未稅金額"]) || 0;
  const tax = parseInt(order["營業稅 5%"]) || 0;
  const goodsTotal = parseInt(order["應付總金額(未含運費)"]) || (subtotal + tax);

  // 商品列 HTML
  let itemsHtml = "";
  let totalQty = 0;
  orderItems.forEach((it, i) => {
    // 贈品判斷：未稅單價為 0（涵蓋滿額贈與福袋大禮包全部 12 樣）
    const price = parseInt(it["未稅單價"]) || 0;
    const isGift = price === 0;
    const giftMark = isGift ? '<span style="background:#d9534f;color:#fff;font-size:9px;padding:1px 5px;border-radius:8px;margin-right:4px;">🎁贈</span>' : '';

    // 數量處理：支援 "2→1" 這種「下訂→實際」格式
    // 顯示時保留原字串；計算金額/加總則用箭頭後的「實際出貨數量」
    const qtyRaw = String(it["數量"] == null ? "" : it["數量"]).trim();
    const hasArrow = /→|->/.test(qtyRaw);
    const arrowParts = qtyRaw.split(/→|->/);
    const actualQtyStr = hasArrow ? arrowParts[arrowParts.length - 1] : qtyRaw;
    const actualQty = parseInt(actualQtyStr, 10) || 0;
    // 出貨單顯示用：有箭頭就保留原字串；否則顯示實際數字
    const qtyDisplay = hasArrow
      ? qtyRaw.replace(/->/g, "→")  // 統一箭頭符號
      : String(actualQty);

    totalQty += actualQty;
    const subTxt = isGift ? '<span style="color:#28a745;">免費</span>' : (price * actualQty).toLocaleString();
    itemsHtml += `
      <tr>
        <td style="padding:6px 4px;border:1px solid #ccc;text-align:center;font-size:11px;">${i + 1}</td>
        <td style="padding:6px 4px;border:1px solid #ccc;font-size:11px;">${it["條碼"] || ""}</td>
        <td style="padding:6px 4px;border:1px solid #ccc;font-size:11px;">${giftMark}${it["商品名稱"] || ""}</td>
        <td style="padding:6px 4px;border:1px solid #ccc;text-align:right;font-size:11px;">${price}</td>
        <td style="padding:6px 4px;border:1px solid #ccc;text-align:center;font-size:11px;font-weight:600;">${qtyDisplay}</td>
        <td style="padding:6px 4px;border:1px solid #ccc;text-align:right;font-size:11px;">${subTxt}</td>
      </tr>`;
  });

  const customerNote = order["備註"] || "";

  // 出貨單 HTML
  const labelHtml = `
<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>
  body { font-family: "Noto Sans TC", "Microsoft JhengHei", Arial, sans-serif; padding: 20px; color: #333; }
  .header { text-align: center; border-bottom: 3px solid #d9534f; padding-bottom: 12px; margin-bottom: 20px; }
  .header h1 { font-size: 22px; color: #d9534f; margin: 0; letter-spacing: 2px; }
  .header .sub { font-size: 13px; color: #888; margin-top: 6px; }
  .section { margin-bottom: 16px; }
  .section-title { font-size: 13px; color: #d9534f; border-left: 4px solid #d9534f; padding-left: 8px; margin-bottom: 8px; font-weight: 700; }
  table.info { width: 100%; border-collapse: collapse; font-size: 12px; }
  table.info td { padding: 4px 8px; border: 1px solid #ddd; }
  table.info td.label { background: #f5f5f5; color: #555; width: 90px; font-weight: 600; }
  table.items { width: 100%; border-collapse: collapse; margin-top: 6px; }
  table.items th { background: #d9534f; color: #fff; padding: 8px 4px; font-size: 11px; border: 1px solid #ccc; }
  .summary { margin-top: 12px; }
  .summary table { width: 100%; border-collapse: collapse; }
  .summary td { padding: 5px 8px; font-size: 12px; }
  .summary td.label { text-align: right; color: #666; width: 70%; }
  .summary td.value { text-align: right; font-weight: 600; }
  .summary tr.total td { background: #fff5f5; color: #d9534f; font-size: 14px; font-weight: 700; border-top: 2px solid #d9534f; }
  .note-box { background: #fff8dc; border: 1px dashed #d4a017; padding: 8px 12px; font-size: 11px; color: #6b4500; margin-top: 8px; border-radius: 4px; }
  .footer { text-align: center; margin-top: 24px; padding-top: 12px; border-top: 1px solid #eee; font-size: 11px; color: #999; }
</style></head>
<body>
  <div class="header">
    <h1>📦 出貨單</h1>
    <div class="sub">PIF 難民特賣會 ／ 毓秀堂貿易有限公司</div>
  </div>

  <div class="section">
    <div class="section-title">訂單資訊</div>
    <table class="info">
      <tr>
        <td class="label">訂單時間</td><td>${order["訂單時間"] || ""}</td>
        <td class="label">訂單編號</td><td>${order["訂單編號"] || ""}</td>
      </tr>
      <tr>
        <td class="label">官網訂單號</td><td><strong>${webOrderNo}</strong></td>
        <td class="label">姓名</td><td><strong>${order["姓名"] || ""}</strong></td>
      </tr>
      <tr>
        <td class="label">聯絡電話</td><td>${order["聯絡電話"] || ""}</td>
        <td class="label">取貨方式</td><td>${delivery}</td>
      </tr>
      <tr>
        <td class="label">付款方式</td><td>${order["付款方式"] || ""}</td>
        <td class="label">希望取貨日</td><td>${order["希望取貨日"] || ""}</td>
      </tr>
      ${!isPickup && address ? `<tr><td class="label">宅配地址</td><td colspan="3">${address}</td></tr>` : ''}
    </table>
  </div>

  <div class="section">
    <div class="section-title">商品明細</div>
    <table class="items">
      <thead>
        <tr>
          <th style="width:36px;">#</th>
          <th style="width:110px;">條碼</th>
          <th>商品名稱</th>
          <th style="width:60px;">未稅單價</th>
          <th style="width:50px;">數量</th>
          <th style="width:70px;">小計</th>
        </tr>
      </thead>
      <tbody>${itemsHtml}</tbody>
    </table>
  </div>

  <div class="summary">
    <table>
      <tr><td class="label">商品未稅金額</td><td class="value">NT$ ${subtotal.toLocaleString()}</td></tr>
      <tr><td class="label">營業稅 5%</td><td class="value">NT$ ${tax.toLocaleString()}</td></tr>
      <tr class="total"><td class="label">應付總金額（未含運費）</td><td class="value">NT$ ${goodsTotal.toLocaleString()}</td></tr>
      <tr><td class="label" style="font-size:11px;color:#999;">總數量</td><td class="value" style="font-size:11px;color:#999;">${totalQty} 件</td></tr>
    </table>
  </div>

  ${customerNote ? `
  <div class="section" style="margin-top:14px;">
    <div class="section-title">客戶備註</div>
    <div class="note-box">${customerNote}</div>
  </div>` : ''}

  <div class="footer">
    出貨單列印日期：${Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm")} ／ 毓秀堂貿易有限公司 ☎️ 02-8283-8333
  </div>
</body></html>`;

  // 用 Drive API 轉 PDF
  const customerName = order["姓名"] || "客戶";
  const fileName = `出貨單_${webOrderNo || order["訂單編號"] || "未知"}_${customerName}.pdf`;
  const blob = Utilities.newBlob(labelHtml, "text/html", fileName.replace(".pdf", ".html"))
    .getAs("application/pdf")
    .setName(fileName);

  // 暫存到 Drive
  const pdfFile = DriveApp.createFile(blob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const downloadUrl = pdfFile.getDownloadUrl();
  const viewUrl = pdfFile.getUrl();
  const fileId = pdfFile.getId();

  // 出貨單下載對話框(精簡版)
  var ST = "<scr" + "ipt>";
  var ET = "</scr" + "ipt>";
  var dialogHtml =
    '<div style="font-family:sans-serif;padding:20px;text-align:center;">'
    + '<h2 style="color:#28a745;font-size:17px;">出貨單已生成</h2>'
    + '<div><strong>' + customerName + '</strong></div>'
    + '<div style="font-size:12px;color:#999;margin:6px 0 16px;">' + fileName + '</div>'
    + '<a href="' + downloadUrl + '" target="_blank" download style="display:inline-block;padding:12px 24px;margin:6px;background:#d9534f;color:#fff;text-decoration:none;border-radius:6px;">下載 PDF</a>'
    + '<a href="' + viewUrl + '" target="_blank" style="display:inline-block;padding:12px 24px;margin:6px;background:#6c757d;color:#fff;text-decoration:none;border-radius:6px;">在新分頁開啟</a>'
    + '<br><button onclick="cleanup()" style="margin-top:10px;padding:10px 20px;background:#6c757d;color:#fff;border:none;border-radius:6px;cursor:pointer;">刪除暫存 PDF</button>'
    + '</div>'
    + ST
    + 'function cleanup(){'
    + 'if(!confirm("確定刪除暫存 PDF？"))return;'
    + 'google.script.run.withSuccessHandler(function(){alert("已刪除");google.script.host.close();}).withFailureHandler(function(e){alert("失敗："+e);}).deleteShippingLabelPdf("' + fileId + '");'
    + '}'
    + ET;

  var dialog = HtmlService.createHtmlOutput(dialogHtml).setWidth(420).setHeight(320).setTitle("出貨單下載");
  ui.showModalDialog(dialog, "出貨單 PDF");
}


/**
 * 刪除暫存的出貨單 PDF (供 exportShippingLabel 內部呼叫)
 */
function deleteShippingLabelPdf(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
    return true;
  } catch (e) {
    console.error("刪除暫存 PDF 失敗:", e);
    throw e;
  }
}
