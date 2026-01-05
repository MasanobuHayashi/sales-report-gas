/**
 * 週報自動作成システム (Rev: Multi-Stage-Synthesis-Update)
 * 修正内容: 
 * 1. プロンプト指示通りの「全体サマリー」および「組織課題」を生成する2段階ステップを実装
 * 2. 部署別詳細レポートからエグゼクティブ・サマリーを抽出するロジックを追加
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
const AI_MODEL = "models/gemini-2.5-flash"; 
const MAX_PROMPT_SIZE_BYTES = 9 * 1024 * 1024; // 9MB
// ---------------------------------------------------------------------------------

/**
 * 週報自動作成のメイン処理
 */
function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  let logMessage = "処理開始 (ボタン押下) \n";
  let logFolderId = null;
  let currentApiKey = null;

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

    // 1. マスターおよびデータ取得
    logMessage += "1. データ取得...\n";
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    if (!masterSheet || !dataSheet) throw new Error("ERR-999: 必要なシートが見つかりません。");
    
    const masterData = masterSheet.getDataRange().getValues();
    const reportData = dataSheet.getDataRange().getValues();
    if (reportData.length <= 1) throw new Error("ERR-001: 週報データがありません。");

    const departmentMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][1]) departmentMap.set(masterData[i][1], masterData[i][2]);
    }

    // 4. 部署ごとのグループ化
    const header = reportData[0];
    const deptGroups = new Map();
    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const deptName = departmentMap.get(row[1]) || "不明";
      if (!deptGroups.has(deptName)) deptGroups.set(deptName, [header]);
      deptGroups.get(deptName).push([...row, deptName]);
    }

    // 5. プロンプト取得
    logMessage += "5. プロンプト雛形取得...\n";
    let promptTemplate;
    try {
      promptTemplate = DocumentApp.openById(settings.promptDocId).getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: プロンプト雛形読込エラー。`);
    }

    currentApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!currentApiKey) throw new Error("APIキー未設定。");

    // --- STEP 1: 部署別詳細レポートの生成 ---
    logMessage += "7-1. 部署別詳細生成ステップ...\n";
    let detailContent = "";
    const sortedDepts = Array.from(deptGroups.keys()).sort();

    for (const deptName of sortedDepts) {
      logMessage += `  -> ${deptName} 生成中...\n`;
      const deptData = deptGroups.get(deptName);
      const inputDataText = _createGroupedInputText(deptData);
      
      const detailPrompt = `${promptTemplate}
---
【部署別レポート指示】
現在は「${deptName}」のセクションを執筆しています。
全体サマリーや組織課題は書かず、この部署の「### 部署名」から始まる詳細（担当者別報告）のみを、1人も漏らさず出力してください。

入力データ:
${inputDataText}
`;
      detailContent += _callGeminiApi(detailPrompt, currentApiKey) + "\n\n";
    }

    // --- STEP 2: 全体サマリーおよび組織課題の生成 ---
    logMessage += "7-2. 全体サマリー・組織課題生成ステップ...\n";
    const analysisPrompt = `${promptTemplate}
---
【全体分析指示】
上記の指示に基づき、全部署の詳細レポート（以下に添付）を分析し、
プロンプト構成の「全体サマリー（冒頭）」および「組織全体の課題と次週の重点事項（文末）」のみを、
シニアセールスマネージャーの視点で執筆してください。

※部署別・担当者別の詳細は、以下の入力データに含まれているため、ここでは出力しないでください。

入力データ（全部署のレポート詳細）:
${detailContent}
`;
    const analysisResult = _callGeminiApi(analysisPrompt, currentApiKey);

    // 8. ドキュメント出力と構成
    logMessage += "8. ドキュメント出力と統合...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();
    const titleDate = Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // 構成の統合： [タイトル] -> [全体サマリー] -> [部署別詳細] -> [組織課題]
    body.appendParagraph(`営業週報（${titleDate} 〜）`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    // 分析結果（冒頭サマリーと末尾課題を含む）をパースして適用
    _applyMarkdownStyles(body, analysisResult + "\n\n" + detailContent);

    doc.saveAndClose();
    const file = DriveApp.getFileById(doc.getId());
    outputFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    logMessage += "P. 処理成功\n";
    ui.alert("週報の自動作成が完了しました。サマリーと組織課題を含む全ての指示内容が反映されました。");

  } catch (e) {
    let safeErrorMessage = currentApiKey ? e.message.split(currentApiKey).join("********") : e.message;
    logMessage += `Z. エラー発生: ${safeErrorMessage}\n`;
    ui.alert(`処理を中断しました。\n\n【エラー内容】\n${safeErrorMessage}`);
  } finally {
    if (logFolderId) _writeLog(logMessage, logFolderId);
    else Logger.log(logMessage);
  }
}

/**
 * Gemini APIを呼び出す内部関数
 */
function _callGeminiApi(prompt, apiKey) {
  const API_URL = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true,
    'timeoutSeconds': 300
  };

  const res = UrlFetchApp.fetch(API_URL, options);
  const resCode = res.getResponseCode();
  const resBody = res.getContentText();

  if (resCode === 200) {
    const json = JSON.parse(resBody);
    return json.candidates?.[0]?.content?.parts?.[0]?.text || "";
  } else {
    throw new Error(`APIエラー (HTTP ${resCode})`);
  }
}

// --- ヘルパー関数 ---
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

function _applyMarkdownStyles(body, rawAiText) {
  if (!rawAiText) return;
  const lines = rawAiText.replace(/^\uFEFF/, "").split('\n');
  lines.forEach(line => {
    try {
      const trimmedLine = line.trim();
      if (trimmedLine.startsWith("---")) { body.appendHorizontalRule(); return; }
      if (trimmedLine === "") { body.appendParagraph(""); return; }

      let plainText = trimmedLine;
      let headingType = null;
      let isBold = false;

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
        headingType = DocumentApp.ParagraphHeading.HEADING3;
        plainText = plainText.substring(5);
        isBold = true;
      }
      else if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        const p = body.appendListItem(line.replace(/^\s*[-*] /, "").trim());
        p.setGlyphType(DocumentApp.GlyphType.BULLET);
        return;
      }

      const paragraph = body.appendParagraph(plainText);
      if (headingType) paragraph.setHeading(headingType);
      if (isBold && plainText.length > 0) paragraph.editAsText().setBold(true);
    } catch (e) {
      Logger.log(`スタイル適用エラー: ${e.message}`);
    }
  });
}