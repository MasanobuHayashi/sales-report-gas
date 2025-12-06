/**
 * 週報自動作成システム (Rev: Limit-Release)
 * 修正内容: 
 * 1. MAX_PROMPT_SIZE_BYTES を 9MB に緩和
 * 2. テキスト圧縮ロジック（改行圧縮）の追加
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
const AI_MODEL = "models/gemini-2.5-pro"; // ★ ユーザー指定モデル
// ★修正: 2MB制限を撤廃し、GASのPayload上限(約10MB)近くまで緩和
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
    // 0. スプレッドシートと設定シートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      throw new Error(`ERR-999: スクリプトエラー。 '${SETTINGS_SHEET_NAME}' が見つかりません。`);
    }

    // 0. 設定シートから各種IDと日付を読み込み
    logMessage += "0. 設定シート読み込み (ID類)... \n";
    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
      startDate: settingsSheet.getRange(START_DATE_CELL).getValue(),
      endDate: settingsSheet.getRange(END_DATE_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. 知識(部署)データ取得 (FOCusユーザマスタシート)
    logMessage += "1. 知識(部署)データ取得 (FOCusユーザマスタシート)... \n";
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) {
      throw new Error(`ERR-999: スクリプトエラー。 '${MASTER_SHEET_NAME}' が見つかりません。`);
    }
    const masterData = masterSheet.getDataRange().getValues();

    // 2. 部署マスターMap作成
    logMessage += "2. 部署マスターMap作成... \n";
    if (masterData.length <= 1) {
      throw new Error("ERR-003: 部署マスター未検出。 'FOCusユーザマスタ'シートにデータがありません。");
    }
    const departmentMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      const staffName = masterData[i][1]; // B列: 担当者
      const departmentName = masterData[i][2]; // C列: 部署名
      if (staffName) {
        departmentMap.set(staffName, departmentName);
      }
    }
    logMessage += `  -> Map作成完了 ( ${departmentMap.size} 件) \n`;

    // 3. データ取得 (週報データ抽出)
    logMessage += "3. データ取得 (週報データ抽出)... \n";
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    if (!dataSheet) {
      throw new Error(`ERR-999: スクリプトエラー。 '${DATA_SHEET_NAME}' が見つかりません。`);
    }
    const reportData = dataSheet.getDataRange().getValues();

    // 3-1. データ存在チェック
    if (reportData.length <= 1) {
      throw new Error("ERR-001: データ入力エラー。 '週報データ抽出'にヘッダー以外のデータがありません。(Query関数で除外されすぎた可能性)");
    }
    logMessage += `  -> データ取得完了 ( ${reportData.length - 1} 件) \n`;

    // 4. データ前処理 (GASが部署名を付与)
    logMessage += "4. データ前処理 (GASが部署名を付与)... \n";
    const processedData = reportData.map((row, index) => {
      if (index === 0) {
        return [...row, "部署名"]; // ヘッダー
      }
      const staffName = row[1]; // B列: 担当者
      const department = departmentMap.get(staffName) || "不明";
      return [...row, department];
    });

    // 5. プロンプト雛形取得 (Google Docs)
    logMessage += "5. プロンプト雛形取得 (Google Docs)... \n";
    let promptTemplate;
    try {
      const promptDoc = DocumentApp.openById(settings.promptDocId);
      promptTemplate = promptDoc.getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: Google Docs エラー。 プロンプト雛形(ID:  ${settings.promptDocId} )が読み込めません。  ${e.message}`);
    }
    logMessage += "  -> 雛形取得完了。 \n";

    // 6. プロンプト結合 (雛形 + 担当者ごとグループ化データ)
    logMessage += "6. プロンプト結合 (担当者ごとグループ化)... \n";
    const inputDataText = _createGroupedInputText(processedData);

    const finalPrompt = `${promptTemplate}

---
入力データ:
${inputDataText}
`;

    // 6-1. オーバーフロー検知 (サイズチェック)
    logMessage += "6-1. オーバーフロー検知 (サイズチェック)... \n";
    const promptSize = Utilities.newBlob(finalPrompt).getBytes().length;
    logMessage += `  -> プロンプトサイズ:  ${promptSize}  バイト (上限: ${MAX_PROMPT_SIZE_BYTES})\n`;
    
    if (promptSize > MAX_PROMPT_SIZE_BYTES) {
      throw new Error(`ERR-002: データ入力エラー。 プロンプトサイズ超過 ( ${promptSize}  バイト)。これ以上のデータ量は分割処理が必要です。`);
    }

    // 7. UrlFetchApp 呼び出し (Gemini API)
    logMessage += "7. UrlFetchApp 呼び出し (Gemini API)... \n";
    let aiResponseText;

    try {
      const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
      if (!API_KEY) {
        throw new Error("スクリプトプロパティ 'GEMINI_API_KEY' が設定されていません。");
      }

      const API_URL = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${API_KEY}`;

      const payload = {
        "contents": [{
          "parts": [{
            "text": finalPrompt
          }]
        }]
      };

      const options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true,
        // timeoutはGASの仕様上、UrlFetchApp.fetchの実行時間上限に従います
      };

      const httpResponse = UrlFetchApp.fetch(API_URL, options);
      const responseCode = httpResponse.getResponseCode();
      const responseBody = httpResponse.getContentText();

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts && jsonResponse.candidates[0].content.parts[0] && jsonResponse.candidates[0].content.parts[0].text) {
          aiResponseText = jsonResponse.candidates[0].content.parts[0].text;
        } else if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].finishReason) {
          throw new Error(`AIが応答を生成できませんでした (Finish Reason:  ${jsonResponse.candidates[0].finishReason} )。 ResponseBody:  ${JSON.stringify(jsonResponse)}`);
        } else {
          throw new Error(`AIからの応答が予期した形式ではありません。 ResponseBody:  ${JSON.stringify(jsonResponse)}`);
        }
        if (!aiResponseText) {
          throw new Error("AIからの応答が空です。");
        }
      } else {
        throw new Error(`Gemini API エラー (HTTP  ${responseCode}): ${responseBody}`);
      }

    } catch (e) {
      if (e.message.includes("Timeout") || e.message.includes("deadline") || e.message.includes("Exceeded maximum execution")) {
        throw new Error(`ERR-101: Gemini AI サービスタイムアウト。 処理が5分以内に完了しませんでした。  ${e.message}`);
      }
      throw new Error(`ERR-100: Gemini AI サービスエラー (UrlFetchApp)。  ${e.message}`);
    }
    logMessage += "  -> AI応答取得完了。 \n";

    // 8. 週報Google Doc出力

    let outputFolder;
    try {
      outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    } catch (e) {
      throw new Error(`ERR-200: Google Drive エラー。 週報出力フォルダ(ID:  ${settings.outputFolderId} )にアクセスできません。  ${e.message}`);
    }

    const fileName = "週報_" + Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // 8-1. 既存ファイル検索・削除
    logMessage += "8-1. 既存ファイル検索・削除... \n";
    const existingFiles = outputFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }

    // 8-2. 週報Google Doc出力 (スタイル適用)
    logMessage += "8-2. 週報Google Doc出力 (スタイル適用)... \n";
    const newDoc = DocumentApp.create(fileName);
    const body = newDoc.getBody();

    _applyMarkdownStyles(body, aiResponseText);

    newDoc.saveAndClose();

    const newFile = DriveApp.getFileById(newDoc.getId());
    outputFolder.addFile(newFile);
    DriveApp.getRootFolder().removeFile(newFile);

    logMessage += `  -> 週報作成完了 (File ID:  ${newDoc.getId()})\n`;
    logMessage += "P. 処理成功 (ユーザーに通知) \n";

    ui.alert("週報の自動作成が完了しました。");

  } catch (e) {
    logMessage += `Z. エラー処理 \n`;
    let errorMessage = e.message;
    if (!errorMessage.startsWith("ERR-")) {
      errorMessage = `ERR-999: 予期せぬスクリプトエラー。  ${errorMessage}`;
    }
    logMessage += `  -> ${errorMessage}\n`;
    if (e.stack) {
      logMessage += `  -> Stack: ${e.stack}\n`;
    }

    ui.alert(`エラーが発生しました。 \n${errorMessage}\n\n 詳細はログファイルを確認してください。`);

  } finally {
    if (logFolderId) {
      _writeLog(logMessage, logFolderId);
    } else {
      Logger.log("[FATAL] ログフォルダIDが取得できなかったため、ログファイル書き込み不可。");
      Logger.log(logMessage);
    }
  }
}

// --- ヘルパー関数 ---

function _writeLog(message, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
    const fileName = `log_${timestamp}.txt`;
    folder.createFile(fileName, message, MimeType.PLAIN_TEXT);
  } catch (e) {
    Logger.log(`[CRITICAL] ログファイル書き込み失敗 (FolderID:  ${folderId}): ${e.message}`);
    Logger.log(`[CRITICAL] 記録失敗したログ: \n${message}`);
  }
}

/**
 * 1つのセルをCSV形式で安全にエスケープし、無駄な空白を圧縮します。
 * ★ 修正: 文字列のカット（Truncate）は行いません。
 */
function _escapeCsvCell(cell) {
  let strCell = (cell === null || cell === undefined) ? "" : String(cell);
  
  // ★追加: 改行コードの正規化と、過剰な改行（3つ以上）の圧縮
  strCell = strCell.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  strCell = strCell.replace(/\n{3,}/g, "\n\n"); // 3つ以上の改行は2つにする

  if (strCell.includes('"')) {
    strCell = strCell.replace(/"/g, '""');
  }
  if (strCell.includes(',') || strCell.includes('\n') || strCell.includes('"')) {
    strCell = `"${strCell}"`;
  }
  return strCell;
}

function _arrayToCsv(data) {
  return data.map(row =>
    row.map(cell => _escapeCsvCell(cell)).join(',')
  ).join('\n');
}

function _createGroupedInputText(processedData) {
  const groupedData = new Map();
  const header = processedData[0];

  const headerMap = new Map();
  header.forEach((h, idx) => headerMap.set(h, idx));

  const idxStaffName = 1; 
  const idxActivityDate = headerMap.get("活動日");
  const idxCustomerName = headerMap.get("顧客名");
  const idxPurpose = headerMap.get("活動目的");
  const idxResult = headerMap.get("予定及び活動結果");
  const idxDoko = headerMap.get("社外同行者");
  const idxDepartment = headerMap.get("部署名");

  const headersForAi = ["部署名", "活動日", "顧客名", "活動目的", "予定及び活動結果", "社外同行者"];
  const indicesToExtract = [idxDepartment, idxActivityDate, idxCustomerName, idxPurpose, idxResult, idxDoko];

  for (let i = 1; i < processedData.length; i++) {
    const row = processedData[i];
    const staffName = row[idxStaffName];
    if (!groupedData.has(staffName)) {
      groupedData.set(staffName, []);
    }
    groupedData.get(staffName).push(row);
  }

  let inputDataText = "";
  for (const [staffName, rows] of groupedData.entries()) {
    inputDataText += `[担当者:  ${staffName}]\n`;
    inputDataText += headersForAi.join(",") + "\n";

    for (const row of rows) {
      const extractedRow = indicesToExtract.map(idx => {
        return (idx !== undefined) ? row[idx] : "";
      });

      // _escapeCsvCell 側で文字数カットせず改行圧縮のみ実施
      const csvRow = extractedRow.map(cell => _escapeCsvCell(cell)).join(',');
      inputDataText += csvRow + "\n";
    }
    inputDataText += "\n";
  }

  return inputDataText;
}

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
      let isBold = false;
      let isListItem = false;
      let indentLevel = 0;

      if (plainText.startsWith("# ")) {
        headingType = DocumentApp.ParagraphHeading.TITLE;
        plainText = plainText.substring(2);
        isBold = true;
      } else if (plainText.startsWith("## ")) {
        headingType = DocumentApp.ParagraphHeading.HEADING1;
        plainText = plainText.substring(3);
        isBold = true;
      } else if (plainText.startsWith("### ")) {
        headingType = DocumentApp.ParagraphHeading.HEADING2;
        plainText = plainText.substring(4);
        isBold = true;
      } else if (plainText.startsWith("#### ")) {
        headingType = null;
        plainText = plainText.substring(5);
        isBold = true;
      } else if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        isListItem = true;
        const indentMatch = line.match(/^\s*/);
        indentLevel = indentMatch ? Math.floor(indentMatch[0].length / 2) : 0;
        plainText = line.replace(/^\s*[-*] /, "").trim();
      }

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
      if (isBold && plainText.length > 0 && paragraph) {
        const textElement = paragraph.editAsText();
        const textLength = textElement.getText().length;
        if (textLength > 0) {
          textElement.setBold(0, textLength - 1, true);
        }
      }
    } catch (e) {
      Logger.log(`スタイル適用エラー:  ${e.message}`);
    }
  });
}