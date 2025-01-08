// LINE developersのメッセージ送受信設定に記載のアクセストークン
const LINE_TOKEN = 'チャネルアクセストークン（ロングターム）'; // Messaging API設定のアクセストークン
const LINE_URL = 'https://api.line.me/v2/bot/message/reply';

function doPost(e) {
  try {
    // POSTデータのパース
    const contents = JSON.parse(e.postData.contents);
    const events = contents.events || [];
    if (events.length === 0) {
      throw new Error('イベントデータが空です');
    }

    const replyToken = events[0].replyToken;
    const userMessage = events[0].message.text;

    if (!replyToken || !userMessage) {
      throw new Error('replyTokenまたはuserMessageが取得できません');
    }

    // メッセージを現場名として取得
    const siteName = userMessage.trim();

    // スプレッドシート設定
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const listSheet = sheet.getSheetByName("シート1");

    if (!listSheet) {
      throw new Error('シート1が見つかりません');
    }

    const data = listSheet.getDataRange().getValues();
    const headers = data[0]; // ヘッダー行
    const siteNameColumnIndex = headers.indexOf("現場名");
    const unitPriceColumnIndex = headers.indexOf("単価");
    const laborColumnIndex = headers.indexOf("人工");
    const totalPriceColumnIndex = headers.indexOf("取引金額");

    if ([siteNameColumnIndex, unitPriceColumnIndex, laborColumnIndex, totalPriceColumnIndex].includes(-1)) {
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
      const defaultUnitPrice = 12500; // 初期単価
      listSheet.appendRow([siteName, defaultUnitPrice, 1, defaultUnitPrice]);
    } else {
      // 人工を1カウントアップ
      const currentLabor = listSheet.getRange(siteRowIndex + 1, laborColumnIndex + 1).getValue();
      const updatedLabor = currentLabor + 1;
      listSheet.getRange(siteRowIndex + 1, laborColumnIndex + 1).setValue(updatedLabor);

      // 取引金額を更新
      const unitPrice = listSheet.getRange(siteRowIndex + 1, unitPriceColumnIndex + 1).getValue();
      const updatedTotalPrice = unitPrice * updatedLabor;
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
    const response = UrlFetchApp.fetch(LINE_URL, {
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

    if (response.getResponseCode() !== 200) {
      throw new Error(`LINE APIの呼び出しに失敗しました: ${response.getContentText()}`);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ 'content': 'post ok' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log(error.message);
    return ContentService.createTextOutput(
      JSON.stringify({ 'error': error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
