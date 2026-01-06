/**
 * 週報自動作成システム (Rev: 2026-Stable-Final)
 * 準拠仕様書: Rev 1.3 + Parallel Processing Update
 */

// --- 定数定義 (変更なし) ---
const SETTINGS_SHEET_NAME = "設定シート";
const PROMPT_DOC_ID_CELL = "B7";
const OUTPUT_FOLDER_ID_CELL = "B8";
const LOG_FOLDER_ID_CELL = "B9";
const MASTER_SHEET_NAME = "FOCusユーザマスタ";
const DATA_SHEET_NAME = "週報データ抽出";
const AI_MODEL = "models/gemini-2.5-flash"; 
const MAX_PROMPT_SIZE_BYTES = 9 * 1024 * 1024;
const MAX_EXECUTION_TIME_MS = 340 * 1000; 

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('週報自動作成')
    .addItem('実行', 'startReportGeneration')
    .addToUi();
}

function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date().getTime();
  let logMessage = `処理開始: ${new Date().toLocaleString()}\n`;
  let logFolderId = null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush();

    // 0. 設定読み込み
    logMessage += "0. 設定シート読み込み...\n";
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error("ERR-999: 設定シートが見つかりません。");

    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. マスター順序の読み込み (【改善】除外フラグの反映)
    logMessage += "1. マスターデータ解析...\n";
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const masterData = masterSheet.getDataRange().getValues();
    const masterOrder = []; 
    const orderedDepts = []; 
    for (let i = 1; i < masterData.length; i++) {
      const name = masterData[i][1];
      const dept = masterData[i][2];
      const isExcluded = masterData[i][3]; // D列: 対象外フラグ
      if (name && dept && !isExcluded) {
        masterOrder.push({ name, dept });
        if (!orderedDepts.includes(dept)) orderedDepts.push(dept);
      }
    }

    // 2. 週次データの取得
    logMessage += "2. 週次データ取得...\n";
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const rawReportData = dataSheet.getDataRange().getValues();
    if (rawReportData.length <= 1) throw new Error("ERR-001: 週報データがありません。");

    let minDate = null;
    let maxDate = null;
    const dataByStaff = new Map();

    for (let i = 1; i < rawReportData.length; i++) {
      const row = rawReportData[i];
      const dateVal = row[0];
      const staff = row[1];
      if (dateVal instanceof Date) {
        if (!minDate || dateVal < minDate) minDate = dateVal;
        if (!maxDate || dateVal > maxDate) maxDate = dateVal;
      }
      if (!dataByStaff.has(staff)) dataByStaff.set(staff, []);
      dataByStaff.get(staff).push(row);
    }

    const tz = Session.getScriptTimeZone();
    const dateRangeStr = (minDate && maxDate) 
      ? `${Utilities.formatDate(minDate, tz, "yyyy年MM月dd日")}～${Utilities.formatDate(maxDate, tz, "yyyy年MM月dd日")}`
      : "期間未特定";

    // 3. プロンプト読み込み
    logMessage += "3. プロンプト雛形取得...\n";
    let promptFull;
    try {
      promptFull = DocumentApp.openById(settings.promptDocId).getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: プロンプト雛形読込エラー。`);
    }
    const promptTemplate = promptFull.split("▽ メンテナンス担当者様へ")[0].trim();

    const currentApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!currentApiKey) throw new Error("ERR-100: APIキー未設定。");

    // --- STEP 1: 並列処理 (変更なし) ---
    logMessage += "4. AI生成リクエスト構築 (並列準備)...\n";
    const requestPayloads = [];
    const requestDepts = [];
    const API_URL = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${currentApiKey}`;

    for (const deptName of orderedDepts) {
      const deptStaffList = masterOrder.filter(m => m.dept === deptName);
      let deptDataForAi = "";
      for (const staff of deptStaffList) {
        if (dataByStaff.has(staff.name)) {
          deptDataForAi += _createStaffText(staff.name, deptName, dataByStaff.get(staff.name)) + "\n";
        } else {
          deptDataForAi += `[担当者: ${staff.name} (部署: ${deptName})] 活動データなし\n`;
        }
      }

      if (deptDataForAi) {
        const detailPrompt = `${promptTemplate}\n---\n【部署別セクション生成】\n部署: ${deptName}\n分析用タグ: 【DEPT_SUMMARY】\n\n入力データ:\n${deptDataForAi}`;
        if (Utilities.newBlob(detailPrompt).getBytes().length > MAX_PROMPT_SIZE_BYTES) continue; 
        requestPayloads.push({
          'url': API_URL, 'method': 'post', 'contentType': 'application/json',
          'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": detailPrompt }] }] }),
          'muteHttpExceptions': true
        });
        requestDepts.push(deptName);
      }
    }

    logMessage += `  -> ${requestPayloads.length}部署分を一括並列送信中...\n`;
    let detailContent = "";
    let analysisSummaries = "";
    const responses = UrlFetchApp.fetchAll(requestPayloads);

    for (let i = 0; i < responses.length; i++) {
      const deptName = requestDepts[i];
      const res = responses[i];
      if (res.getResponseCode() === 200) {
        const aiText = JSON.parse(res.getContentText()).candidates?.[0]?.content?.parts?.[0]?.text || "";
        const parts = aiText.split("【DEPT_SUMMARY】");
        detailContent += parts[0].trim() + "\n\n";
        analysisSummaries += `■部署: ${deptName}\n${parts[1] || ""}\n\n`;
      }
    }

    // --- STEP 2: 全体統合 ---
    logMessage += "5. 全体要約生成 (統合)...\n";
    const analysisPrompt = `${promptTemplate}\n---\n【全体統合指示】\n1. タイトル日付を「${dateRangeStr}」としてください。\n2. {{DETAIL_PLACEHOLDER}} の位置に詳細を結合します。\n\n分析用インプット:\n${analysisSummaries}`;
    
    const finalShell = _callGeminiApiSingle(analysisPrompt, currentApiKey);
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 4. 出力処理 (ファイル名は maxDate を維持)
    logMessage += "6. ファイル出力...\n";
    const fileName = "週報_" + Utilities.formatDate(maxDate || new Date(), tz, "yyyy-MM-dd");
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const existingFiles = outputFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);
    
    const doc = DocumentApp.create(fileName);
    _applyMarkdownStyles(doc.getBody(), finalFullText);
    doc.saveAndClose();
    
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    logMessage += "処理完了: 成功\n";
    ui.alert("作成完了", "週報の自動作成が完了しました。", ui.ButtonSet.OK);

  } catch (e) {
    logMessage += `\n❌ 異常終了: ${e.message}\n`;
    ui.alert("エラー発生", `処理を中断しました。\n${e.message}`, ui.ButtonSet.OK);
  } finally {
    if (logFolderId) DriveApp.getFolderById(logFolderId).createFile(`log_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss")}.txt`, logMessage);
  }
}

function _callGeminiApiSingle(prompt, apiKey) {
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }), 'muteHttpExceptions': true };
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`, options);
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  throw new Error(`API Error: ${res.getResponseCode()}`);
}

function _createStaffText(staff, dept, rows) {
  let txt = `[担当者: ${staff} (部署: ${dept})]\n`;
  rows.forEach(r => {
    let dStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd") : String(r[0]).substring(0, 10);
    txt += `- ${dStr} ${r[2]} / ${r[4]}\n`;
  });
  return txt;
}

/**
 * 修正版：見出しスタイルの適用時に太字を強制する
 */
function _applyMarkdownStyles(body, rawAiText) {
  if (!rawAiText) return;
  const lines = rawAiText.replace(/^\uFEFF/, "").split('\n');
  lines.forEach(line => {
    let plain = line.trim();
    if (plain === "") { body.appendParagraph(""); return; }

    let head = null;
    if (plain.startsWith("# ")) { head = DocumentApp.ParagraphHeading.TITLE; plain = plain.substring(2); }
    else if (plain.startsWith("## ")) { head = DocumentApp.ParagraphHeading.HEADING1; plain = plain.substring(3); }
    else if (plain.startsWith("### ")) { head = DocumentApp.ParagraphHeading.HEADING2; plain = plain.substring(4); }
    else if (plain.startsWith("#### ")) { head = DocumentApp.ParagraphHeading.HEADING3; plain = plain.substring(5); }

    if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
      body.appendListItem(plain.replace(/^[-*]\s+/, "")).setGlyphType(DocumentApp.GlyphType.BULLET);
    } else {
      const p = body.appendParagraph(plain);
      // 【修正】スタイル適用と同時に太字(Bold)を設定する
      if (head) p.setHeading(head).setBold(true);
    }
  });
}