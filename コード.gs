/**
 * 週報自動作成システム (Rev: Error-Handling-Update)
 * 修正内容: エラー発生時のUIメッセージを変更（開発者への問い合わせを誘導）
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('週報自動作成')
    .addItem('実行', 'startReportGeneration')
    .addToUi();
}

// --- 定数定義 ---
const SETTINGS_SHEET_NAME = "設定シート";
const PROMPT_DOC_ID_CELL = "B7";
const OUTPUT_FOLDER_ID_CELL = "B8";
const LOG_FOLDER_ID_CELL = "B9";
const START_DATE_CELL = "B3";
const END_DATE_CELL = "B4";
const MASTER_SHEET_NAME = "FOCusユーザマスタ";
const DATA_SHEET_NAME = "週報データ抽出";
const AI_MODEL = "models/gemini-2.5-pro"; 
const MAX_PROMPT_SIZE_BYTES = 9 * 1024 * 1024; // 9MB
// ---------------------------------------------------------------------------------

/**
 * 週報自動作成のメイン処理
 */
function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  let logMessage = "処理開始 (ボタン押下) \n";
  let logFolderId = null;

  try {
    // 0. 設定読み込み
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`ERR-999: シート '${SETTINGS_SHEET_NAME}' が見つかりません。`);

    logMessage += "0. 設定シート読み込み...\n";
    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
      startDate: settingsSheet.getRange(START_DATE_CELL).getValue(),
      endDate: settingsSheet.getRange(END_DATE_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. 知識データ取得
    logMessage += "1. 知識データ取得...\n";
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error(`ERR-999: シート '${MASTER_SHEET_NAME}' が見つかりません。`);
    const masterData = masterSheet.getDataRange().getValues();

    // 2. 部署Map作成
    logMessage += "2. 部署Map作成...\n";
    if (masterData.length <= 1) throw new Error("ERR-003: 部署マスターデータがありません。");
    const departmentMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][1]) departmentMap.set(masterData[i][1], masterData[i][2]);
    }
    logMessage += `  -> Map作成完了 (${departmentMap.size}件)\n`;

    // 3. データ取得
    logMessage += "3. データ取得...\n";
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    if (!dataSheet) throw new Error(`ERR-999: シート '${DATA_SHEET_NAME}' が見つかりません。`);
    const reportData = dataSheet.getDataRange().getValues();

    // 3-1. データ存在チェック
    if (reportData.length <= 1) throw new Error("ERR-001: 週報データがありません。抽出条件を確認してください。");
    logMessage += `  -> データ取得完了 (${reportData.length - 1}件)\n`;

    // 4. データ前処理
    logMessage += "4. データ前処理...\n";
    const processedData = reportData.map((row, index) => {
      if (index === 0) return [...row, "部署名"];
      return [...row, departmentMap.get(row[1]) || "不明"];
    });

    // 5. プロンプト雛形取得
    logMessage += "5. プロンプト雛形取得...\n";
    let promptTemplate;
    try {
      promptTemplate = DocumentApp.openById(settings.promptDocId).getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: プロンプト雛形(ID: ${settings.promptDocId})が読み込めません。`);
    }

    // 6. プロンプト結合
    logMessage += "6. プロンプト結合...\n";
    const inputDataText = _createGroupedInputText(processedData);
    const finalPrompt = `${promptTemplate}\n\n---\n入力データ:\n${inputDataText}\n`;

    // 6-1. サイズチェック
    const promptSize = Utilities.newBlob(finalPrompt).getBytes().length;
    logMessage += `  -> サイズ: ${promptSize} byte\n`;
    if (promptSize > MAX_PROMPT_SIZE_BYTES) throw new Error(`ERR-002: プロンプトサイズ超過 (${promptSize} byte)。`);

    // 7. API呼び出し
    logMessage += "7. Gemini API 呼び出し...\n";
    let aiResponseText;
    try {
      const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!API_KEY) throw new Error("スクリプトプロパティ 'GEMINI_API_KEY' が未設定です。");

      const API_URL = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${API_KEY}`;
      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": finalPrompt }] }] }),
        'muteHttpExceptions': true,
        'timeoutSeconds': 300
      };

      const res = UrlFetchApp.fetch(API_URL, options);
      const resCode = res.getResponseCode();
      const resBody = res.getContentText();

      if (resCode === 200) {
        const json = JSON.parse(resBody);
        if (json.candidates?.[0]?.content?.parts?.[0]?.text) {
          aiResponseText = json.candidates[0].content.parts[0].text;
        } else {
          throw new Error("AI応答形式エラー。");
        }
      } else {
        throw new Error(`APIエラー (HTTP ${resCode}): ${resBody.substring(0, 200)}...`);
      }
    } catch (e) {
      if (e.message.includes("Timeout") || e.message.includes("deadline")) {
        throw new Error(`ERR-101: タイムアウト(5分超過)。`);
      }
      throw new Error(`ERR-100: AI通信エラー。 ${e.message}`);
    }
    logMessage += "  -> 応答取得完了。\n";

    // 8. 出力処理
    logMessage += "8. 出力処理...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    _applyMarkdownStyles(doc.getBody(), aiResponseText);
    doc.saveAndClose();
    
    const file = DriveApp.getFileById(doc.getId());
    outputFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    logMessage += "P. 処理成功\n";
    ui.alert("週報の自動作成が完了しました。");

  } catch (e) {
    // ★★★ エラーハンドリング変更箇所 ★★★
    logMessage += `Z. エラー発生: ${e.message}\nStack: ${e.stack}\n`;
    
    let userMessage = e.message;
    // システム内部エラーの場合のプレフィックス付与
    if (!userMessage.startsWith("ERR-")) {
      userMessage = `ERR-999: 予期せぬエラー\n(${userMessage})`;
    }

    // ユーザー向けアラート文言の構築
    const alertText = `処理を中断しました。

【エラー内容】
${userMessage}

--------------------------------------------------
この画面のスクリーンショットを撮り、
システム開発者へお問い合わせください。
--------------------------------------------------
※詳細はログフォルダのテキストファイルをご確認ください。`;

    ui.alert(alertText);

  } finally {
    if (logFolderId) _writeLog(logMessage, logFolderId);
    else Logger.log(logMessage);
  }
}

// --- ヘルパー関数 (変更なし) ---
function _writeLog(msg, fid) {
  try { DriveApp.getFolderById(fid).createFile(`log_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss")}.txt`, msg); } catch(e){}
}
function _escapeCsvCell(c) {
  let s = (c == null) ? "" : String(c);
  s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n{3,}/g, "\n\n");
  if (s.includes('"') || s.includes(',') || s.includes('\n')) s = `"${s.replace(/"/g, '""')}"`;
  return s;
}
function _createGroupedInputText(data) {
  const map = new Map();
  const h = data[0];
  const hMap = new Map();
  h.forEach((v, i) => hMap.set(v, i));
  const targets = ["部署名", "活動日", "顧客名", "活動目的", "予定及び活動結果", "社外同行者"];
  const idxs = targets.map(t => hMap.get(t));
  
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const s = r[1];
    if (!map.has(s)) map.set(s, []);
    map.get(s).push(r);
  }
  
  let txt = "";
  for (const [s, rows] of map.entries()) {
    txt += `[担当者: ${s}]\n` + targets.join(",") + "\n";
    for (const r of rows) txt += idxs.map(i => _escapeCsvCell(r[i])).join(',') + "\n";
    txt += "\n";
  }
  return txt;
}
/**
 * AIが生成したマークダウンテキストを解析し、Google Docのスタイルを適用します。
 * 修正: "####" (担当者名) を「太字の本文」ではなく「見出し3」として扱い、以前の書式に戻しました。
 */
function _applyMarkdownStyles(body, rawAiText) {

  if (!rawAiText) {
    return;
  }

  const cleanedText = rawAiText.replace(/^\uFEFF/, "").trim();
  const lines = cleanedText.split('\n');

  lines.forEach(line => {

    try {
      const trimmedLine = line.trim();

      if (trimmedLine.startsWith("---") && trimmedLine.length < 5) {
        body.appendHorizontalRule();
        return;
      }
      if (trimmedLine === "") {
        body.appendParagraph("");
        return;
      }

      let plainText = trimmedLine;
      let headingType = null;
      let isBold = false; // 見出しのみ太字
      let isListItem = false;
      let indentLevel = 0;

      // スタイルの判定 (見出し)
      // Googleドキュメントの階層構造に合わせてマッピングします
      if (plainText.startsWith("# ")) {
        headingType = DocumentApp.ParagraphHeading.TITLE; // タイトル
        plainText = plainText.substring(2);
        isBold = true;
      } else if (plainText.startsWith("## ")) {
        headingType = DocumentApp.ParagraphHeading.HEADING1; // 見出し1 (セクション)
        plainText = plainText.substring(3);
        isBold = true;
      } else if (plainText.startsWith("### ")) {
        headingType = DocumentApp.ParagraphHeading.HEADING2; // 見出し2 (部署名)
        plainText = plainText.substring(4);
        isBold = true;
      } else if (plainText.startsWith("#### ")) {
        // ★修正箇所: ここを「見出し3」に戻します
        headingType = DocumentApp.ParagraphHeading.HEADING3; // 見出し3 (担当者名)
        plainText = plainText.substring(5);
        isBold = true;
      }
      // リスト (-) の判定
      else if (rawLineContent = line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        isListItem = true;
        const indentMatch = line.match(/^\s*/);
        indentLevel = indentMatch ? Math.floor(indentMatch[0].length / 2) : 0;
        plainText = line.replace(/^\s*[-*] /, "").trim();
      }

      // ドキュメントに書き込む
      let paragraph;

      if (isListItem) {
        const listItem = body.appendListItem(plainText);
        listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
        if (indentLevel > 0) {
          listItem.setIndentFirstLine(18 * indentLevel);
          listItem.setIndentStart(36 * indentLevel);
        }
        paragraph = listItem;
      } else {
        paragraph = body.appendParagraph(plainText);
        if (headingType) {
          paragraph.setHeading(headingType);
        }
      }

      // 行全体の太字を適用 (見出しの場合のみ)
      if (isBold && plainText.length > 0 && paragraph) {
        const textElement = paragraph.editAsText();
        const textLength = textElement.getText().length;
        if (textLength > 0) {
          textElement.setBold(0, textLength - 1, true);
        }
      }

    } catch (e) {
      Logger.log(`スタイル適用エラー: ${e.message}`);
    }
  });
}