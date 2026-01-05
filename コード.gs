/**
 * 週報自動作成システム (Rev: 2024-Compliance-Optimized)
 * モデル: Gemini 2.5 Flash
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
const MASTER_SHEET_NAME = "FOCusユーザマスタ";
const DATA_SHEET_NAME = "週報データ抽出";
const AI_MODEL = "models/gemini-2.5-flash"; // 最新のFlashモデルを指定
// ---------------------------------------------------------------------------------

function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  let logMessage = "処理開始\n";
  let logFolderId = null;
  let currentApiKey = null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SpreadsheetApp.flush();

    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. マスター順序の読み込み（FOCusユーザマスタ準拠）
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const masterData = masterSheet.getDataRange().getValues();
    const masterOrder = []; 
    const orderedDepts = []; 
    for (let i = 1; i < masterData.length; i++) {
      const name = masterData[i][1];
      const dept = masterData[i][2];
      if (name && dept) {
        masterOrder.push({ name, dept });
        if (!orderedDepts.includes(dept)) orderedDepts.push(dept);
      }
    }

    // 2. 週次データの取得と日付範囲の自動算出 (Point 4対応)
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

      // 日付の最小・最大を計算
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

    // 3. プロンプトの読み込みとメンテナンスセクションの除外 (Point 2対応)
    let promptFull;
    try {
      promptFull = DocumentApp.openById(settings.promptDocId).getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: プロンプト雛形読込エラー。`);
    }
    // メンテナンス用テキストを除外
    const promptTemplate = promptFull.split("▽ メンテナンス担当者様へ")[0].trim();

    currentApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!currentApiKey) throw new Error("APIキー未設定。");

    // --- STEP 1: 部署別詳細生成（マスタ順） ---
    logMessage += "部署別詳細生成...\n";
    let detailContent = "";
    let analysisSummaries = "";

    for (const deptName of orderedDepts) {
      const deptStaffList = masterOrder.filter(m => m.dept === deptName);
      let deptDataForAi = "";
      for (const staff of deptStaffList) {
        if (dataByStaff.has(staff.name)) {
          deptDataForAi += _createStaffText(staff.name, deptName, dataByStaff.get(staff.name)) + "\n";
        } else {
          // 活動なしの場合もマスタにある担当者は出力対象（プロンプト側の「活動がない場合」の指示に従う）
          deptDataForAi += `[担当者: ${staff.name} (部署: ${deptName})] 活動データなし\n`;
        }
      }

      if (deptDataForAi) {
        const detailPrompt = `${promptTemplate}\n---\n【部署別セクション生成】\n部署: ${deptName}\n分析用タグ: 【DEPT_SUMMARY】\n\n入力データ:\n${deptDataForAi}`;
        const res = _callGeminiApi(detailPrompt, currentApiKey);
        const parts = res.split("【DEPT_SUMMARY】");
        detailContent += parts[0].trim() + "\n\n";
        analysisSummaries += `■部署: ${deptName}\n${parts[1] || ""}\n\n`;
      }
    }

    // --- STEP 2: 全体統合 ---
    logMessage += "全体要約生成...\n";
    const analysisPrompt = `${promptTemplate}\n---\n【全体統合指示】\n1. タイトル日付を「${dateRangeStr}」としてください。\n2. {{DETAIL_PLACEHOLDER}} の位置に詳細を結合します。\n\n分析用インプット:\n${analysisSummaries}`;
    const finalShell = _callGeminiApi(analysisPrompt, currentApiKey);
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 4. 出力
    const doc = DocumentApp.create("週報_" + Utilities.formatDate(minDate || new Date(), tz, "yyyy-MM-dd"));
    _applyMarkdownStyles(doc.getBody(), finalFullText);
    doc.saveAndClose();
    
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    ui.alert("週報が完成しました。");
  } catch (e) {
    ui.alert(`エラー: ${e.message}`);
  }
}

function _callGeminiApi(prompt, apiKey) {
  const options = {
    'method': 'post', 'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true
  };
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`, options);
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  throw new Error(`API Error: ${res.getResponseCode()}`);
}

function _createStaffText(staff, dept, rows) {
  let txt = `[担当者: ${staff} (部署: ${dept})]\n`;
  rows.forEach(r => {
    const dStr = r[0] instanceof Date ? Utilities.formatDate(r[0], "JST", "MM/dd") : String(r[0]);
    txt += `- ${dStr} ${r[2]} / ${r[4]}\n`;
  });
  return txt;
}

function _applyMarkdownStyles(body, rawAiText) {
  if (!rawAiText) return;
  const lines = rawAiText.split('\n');
  lines.forEach(line => {
    let plain = line.trim();
    if (plain === "") return;

    let head = null;
    if (plain.startsWith("# ")) { head = DocumentApp.ParagraphHeading.TITLE; plain = plain.substring(2); }
    else if (plain.startsWith("## ")) { head = DocumentApp.ParagraphHeading.HEADING1; plain = plain.substring(3); }
    else if (plain.startsWith("### ")) { head = DocumentApp.ParagraphHeading.HEADING2; plain = plain.substring(4); }
    else if (plain.startsWith("#### ")) { head = DocumentApp.ParagraphHeading.HEADING3; plain = plain.substring(5); }

    let p;
    if (line.match(/^(\s*)- /)) {
      p = body.appendListItem(line.replace(/^\s*[-*] /, "").trim()).setGlyphType(DocumentApp.GlyphType.BULLET);
    } else {
      p = body.appendParagraph(plain);
      if (head) p.setHeading(head).setBold(true);
    }
  });
}