/**
 * Shonandai Kitchen Studio 予約・問い合わせ受付用 Google Apps Script (GAS)
 * 
 * 【使い方】
 * 1. Google スプレッドシートを新規作成（例:「湘南台キッチンスタジオ_予約問い合わせ台帳」）
 * 2. 1行目にヘッダーを記入：
 *    [受信日時, 管理ID, お名前/貴社名, メールアドレス, 電話番号, 希望プラン, 利用希望日, 利用用途, お問い合わせ内容]
 * 3. メニュー「拡張機能」>「Apps Script」を開く。
 * 4. 本コードを貼り付けて保存。
 * 5. 右上「デプロイ」>「新しいデプロイ」> 種類の選択「ウェブアプリ」
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員（Anyone）
 * 6. 発行された「ウェブアプリURL」を Webサイトの送信先に指定。
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    var timestamp = new Date();
    var id = "REQ-" + Utilities.formatDate(timestamp, "JST", "yyyyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
    
    // スプレッドシートに追記
    sheet.appendRow([
      timestamp,
      id,
      data.name || "",
      data.email || "",
      data.tel || "",
      data.plan || "",
      data.date || "",
      data.usage || "",
      data.message || ""
    ]);
    
    // 管理者宛てにGmail通知（任意）
    var adminEmail = "o.naoki0503@intakingresources.com"; // または通知したいメール
    var subject = "【湘南台キッチンスタジオ】新規予約・問い合わせ受付：" + (data.name || "お客様");
    var body = "Shonandai Kitchen Studio のWebサイトより新しいお問い合わせを受信しました。\n\n"
             + "■ 管理ID: " + id + "\n"
             + "■ 受信日時: " + Utilities.formatDate(timestamp, "JST", "yyyy/MM/dd HH:mm") + "\n"
             + "■ お名前: " + data.name + "\n"
             + "■ メール: " + data.email + "\n"
             + "■ 電話番号: " + data.tel + "\n"
             + "■ 希望プラン: " + data.plan + "\n"
             + "■ 利用希望日: " + data.date + "\n"
             + "■ 利用用途: " + data.usage + "\n"
             + "■ お問い合わせ・ご要望:\n" + data.message + "\n\n"
             + "スプレッドシートを確認してください。";
             
    MailApp.sendEmail(adminEmail, subject, body);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", id: id }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
