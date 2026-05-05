// ===================================================
// Google Sheets × Gemini AI 売上自動分析レポート
// ===================================================

const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY';

function analyzeAndReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('売上データ');
  const reportSheet = ss.getSheetByName('AIレポート');

  // 売上データを取得（2行目以降、最大30行）
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('売上データが入力されていません。');
    return;
  }

  const range = dataSheet.getRange(2, 1, lastRow - 1, 4); // A〜D列
  const values = range.getValues();

  // データを整形してプロンプト用テキストに変換
  const dataText = values
    .filter(row => row[0] !== '')
    .map(row => {
      const date = Utilities.formatDate(new Date(row[0]), 'Asia/Tokyo', 'yyyy/MM/dd');
      return `${date}：売上 ${row[1].toLocaleString()}円、客数 ${row[2]}人、天候 ${row[3]}`;
    })
    .join('\n');

  // Gemini APIへのプロンプト
  const prompt = `以下は店舗の日次売上データです。
このデータをもとに以下の3点を分析してください。

【売上データ】
${dataText}

【分析してほしい内容】
1. 売上トレンドの要約（2〜3文）
2. 気になる数字や特徴（良い点・悪い点）
3. 来週に向けた改善提案（具体的に2点）

日本語で、店舗オーナーにわかりやすく回答してください。`;

  // Gemini API呼び出し
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      maxOutputTokens: 800,
      thinkingConfig: { thinkingBudget: 0 }
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });

  const result = JSON.parse(response.getContentText());
  const aiText = result.candidates[0].content.parts[0].text;

  // AIレポートシートに書き込み
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  reportSheet.clearContents();
  reportSheet.getRange('A1').setValue('AI売上分析レポート');
  reportSheet.getRange('A2').setValue(`生成日時：${now}`);
  reportSheet.getRange('A4').setValue(aiText);

  // 書式調整
  reportSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  reportSheet.getRange('A4').setWrap(true);
  reportSheet.setColumnWidth(1, 600);

  SpreadsheetApp.getUi().alert('AIレポートを生成しました！「AIレポート」シートを確認してください。');
}

// メニューにボタンを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI分析')
    .addItem('売上レポートを生成', 'analyzeAndReport')
    .addToUi();
}
