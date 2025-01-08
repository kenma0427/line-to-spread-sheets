// LINE developersのメッセージ送受信設定に記載のアクセストークン
const LINE_TOKEN = 'チャネルアクセストークン（ロングターム）'; // Messaging API設定のアクセストークン
const LINE_URL = 'https://api.line.me/v2/bot/message/reply';

function doPost(e) {
  // 応答用Tokenを取得
  const replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  const userMessage = JSON.parse(e.postData.contents).events[0].message.text;

  // メッセージを現場名として取得
  const siteName = userMessage.trim(); // トリムして余分な空白を除去

  // スプレッドシート設定
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = sheet.getSheetByName("シート1"); 

  // データ範囲と最終行を取得
  const data = listSheet.getDataRange().getValues();
  const headers = data[0]; // ヘッダー行
  const siteNameColumnIndex = headers.indexOf("現場名");
  const unitPriceColumnIndex = headers.indexOf("単価");
  const laborColumnIndex = headers.indexOf("人工");
  const totalPriceColumnIndex = headers.indexOf("取引金額");

  if (siteNameColumnIndex === -1 || unitPriceColumnIndex === -1 || laborColumnIndex === -1 || totalPriceColumnIndex === -1) {
    throw new Error("必要な列名（現場名、単価、人工、取引金額）が見つかりません");
  }

  // 現場名の存在確認
  let siteRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][siteNameColumnIndex] === siteName) {
      siteRowIndex = i; // 現場名が見つかった行インデックス
      break;
    }
  }

  if (siteRowIndex === -1) {
    // 新規現場名を追加
    listSheet.appendRow([siteName, 12500, 1, 12500]); // 初期値：単価=12500, 人工=1, 取引金額=12500
  } else {
    // 人工を1カウントアップ
    const currentLabor = listSheet.getRange(siteRowIndex + 1, laborColumnIndex + 1).getValue();
    const updatedLabor = currentLabor + 1;
    listSheet.getRange(siteRowIndex + 1, laborColumnIndex + 1).setValue(updatedLabor);

    // 取引金額を更新
    const updatedTotalPrice = 12500 * updatedLabor;
    listSheet.getRange(siteRowIndex + 1, totalPriceColumnIndex + 1).setValue(updatedTotalPrice);
  }

  // LINE返信メッセージ
  const messages = [
    {
      'type': 'text',
      'text': `現場名「${siteName}」のデータを更新しました。`
    }
  ];

  // LINEで返信
  UrlFetchApp.fetch(LINE_URL, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': `Bearer ${LINE_TOKEN}`,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': messages,
    }),
  });

  return ContentService.createTextOutput(
    JSON.stringify({ 'content': 'post ok' })
  ).setMimeType(ContentService.MimeType.JSON);
}
