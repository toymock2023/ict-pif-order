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


/**
 * 接收 POST 請求，處理訂單寫入
 */
function doPost(e) {
  try {
    const order = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. 寫入「訂單總覽」工作表
    const summarySheet = getOrCreateSheet(ss, SUMMARY_SHEET, [
      "訂單時間", "訂單編號", "姓名", "公司/部門", "聯絡電話", "Email",
      "取貨方式", "宅配地址", "付款方式", "希望取貨日",
      "購買品項摘要", "總數量",
      "未稅金額", "營業稅 5%", "應付總金額(未含運費)",
      "備註"
    ]);

    const orderId = generateOrderId();

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
 * 提供 GET 請求測試 (測試 Web App 是否成功部署)
 */
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({
      status: "ok",
      message: "PIF 難民特賣會訂單系統運作中",
      time: new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" })
    })
  ).setMimeType(ContentService.MimeType.JSON);
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
    body += `  ・${it.name}${specText} × ${it.qty} = NT$${it.subtotal}\n`;
  });
  body += `─────────────\n`;
  body += `總數量：${order.totalQty} 件\n`;
  body += `未稅金額：NT$ ${order.subtotalAmount}\n`;
  body += `營業稅 5%：NT$ ${order.taxAmount}\n`;
  body += `應付總金額：NT$ ${order.totalAmount}（運費另計）\n`;

  MailApp.sendEmail(recipient, subject, body);
}
