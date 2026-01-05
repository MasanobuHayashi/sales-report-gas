/**
 * 週報自動作成システム (Rev: Signature-and-Format-Strict-Update)
 * 修正内容: 
 * 1. Utilities.formatDate の引数エラーを解消（Dateオブジェクトへの明示的変換）
 * 2. タイトル二重出力の解消（AIにタイトルの執筆と日付置換を完全に委ねる）
 * 3. 文中太字のパース精度向上とハードコーディングの徹底排除
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

/**
 * 週報自動作成のメイン処理
 */
function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  let logMessage = "処理開始 (ボタン押下) \n";
  let logFolderId = null;
  let currentApiKey = null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`ERR-999: 設定シートが見つかりません。`);

    // 日付取得の安全性を確保
    const rawStart = settingsSheet.getRange(START_DATE_CELL).getValue();
    const rawEnd = settingsSheet.getRange(END_DATE_CELL).getValue();
    if (!rawStart || !rawEnd) throw new Error("ERR-001: 設定シートの日付が空です。");
    
    // 明示的にDateオブジェクトに変換（シグネチャエラー対策）
    const startDate = new Date(rawStart);
    const endDate = new Date(rawEnd);
    const tz = Session.getScriptTimeZone();
    const dateRangeStr = `${Utilities.formatDate(startDate, tz, "yyyy年MM月dd日")}～${Utilities.formatDate(endDate, tz, "yyyy年MM月dd日")}`;

    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. マスタ順序の読み込み
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

    // 2. データ取得
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const rawReportData = dataSheet.getDataRange().getValues();
    if (rawReportData.length <= 1) throw new Error("ERR-001: 週報データがありません。");

    const dataByStaff = new Map();
    for (let i = 1; i < rawReportData.length; i++) {
      const staff = rawReportData[i][1];
      if (!dataByStaff.has(staff)) dataByStaff.set(staff, []);
      dataByStaff.get(staff).push(rawReportData[i]);
    }

    let promptTemplate;
    try {
      promptTemplate = DocumentApp.openById(settings.promptDocId).getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: プロンプト雛形読込エラー。`);
    }

    currentApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!currentApiKey) throw new Error("APIキー未設定。");

    // --- STEP 1: 部署別詳細生成 ---
    logMessage += "7-1. 部署別詳細生成（マスタ順）...\n";
    let detailContent = "";
    let analysisSummaries = "";

    for (const deptName of orderedDepts) {
      const deptStaffList = masterOrder.filter(m => m.dept === deptName);
      let deptDataForAi = "";
      for (const staff of deptStaffList) {
        if (dataByStaff.has(staff.name)) {
          deptDataForAi += _createStaffText(staff.name, deptName, dataByStaff.get(staff.name)) + "\n";
        }
      }

      if (deptDataForAi) {
        logMessage += `  -> ${deptName} 生成中...\n`;
        const detailPrompt = `${promptTemplate}
---
【部署別セクション生成指示】
1. 現在は「${deptName}」の詳細報告セクションを執筆しています。
2. 指示書の構成案に基づき、担当者別の活動詳細をマークダウンで出力してください。
3. 文末に必ず【DEPT_SUMMARY】というタグを入れ、続けてこの部署の分析用要約を記述してください。
4. メタ発言は厳禁です。

入力データ:
${deptDataForAi}
`;
        const res = _callGeminiApi(detailPrompt, currentApiKey);
        const parts = res.split("【DEPT_SUMMARY】");
        detailContent += parts[0].trim() + "\n\n";
        analysisSummaries += `■部署: ${deptName}\n${parts[1] || ""}\n\n`;
      }
    }

    // --- STEP 2: 全体統合および構成生成 ---
    logMessage += "7-2. 全体構成生成（タイトル日付置換）...\n";
    const analysisPrompt = `${promptTemplate}
---
【全体分析・構成生成指示】
1. プロンプトの構成定義に基づき、「タイトル」「全体サマリー」「組織全体の課題」セクションを執筆してください。
2. タイトル内の日付プレースホルダ（yyyy年...）は必ず「${dateRangeStr}」に書き換えてください。
3. 詳細セクション（部署別報告）を挿入すべき場所に、単独行で「{{DETAIL_PLACEHOLDER}}」と記述してください。
4. メタ発言を排除し、マークダウン書式（#～####, **太字**）を厳格に適用してください。

分析用インプット:
${analysisSummaries}
`;
    const finalShell = _callGeminiApi(analysisPrompt, currentApiKey);
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 8. ドキュメント出力
    logMessage += "8. ドキュメント出力統合...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(startDate, tz, "yyyy-MM-dd");
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();
    
    // AIが生成したテキストをスタイリングしながら流し込む（ハードコードされたタイトルは無し）
    _applyMarkdownStyles(body, finalFullText);

    doc.saveAndClose();
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    logMessage += "P. 処理成功\n";
    ui.alert("週報が完成しました。");

  } catch (e) {
    let safeErr = currentApiKey ? e.message.split(currentApiKey).join("********") : e.message;
    logMessage += `Z. エラー: ${safeErr}\n`;
    ui.alert(`エラーが発生しました。\n${safeErr}`);
  } finally {
    if (logFolderId) _writeLog(logMessage, logFolderId);
  }
}

/** Gemini API呼び出し */
function _callGeminiApi(prompt, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true,
    'timeoutSeconds': 300
  };
  const res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  throw new Error(`APIエラー: ${res.getResponseCode()}`);
}

function _writeLog(msg, fid) {
  try { DriveApp.getFolderById(fid).createFile(`log_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss")}.txt`, msg); } catch(e){}
}

function _createStaffText(staff, dept, rows) {
  const targets = ["部署名", "活動日", "顧客名", "活動目的", "予定及び活動結果", "社外同行者"];
  let txt = `[担当者: ${staff} (部署: ${dept})]\n${targets.join(",")}\n`;
  rows.forEach(r => {
    const dStr = r[0] instanceof Date ? Utilities.formatDate(r[0], "JST", "MM/dd") : String(r[0]);
    const vals = [dept, dStr, r[2], r[3], r[4], r[5]];
    txt += vals.map(v => _escCsv(v)).join(',') + "\n";
  });
  return txt;
}

function _escCsv(c) {
  let s = (c == null) ? "" : String(c).replace(/\n/g, " ");
  return (s.includes('"') || s.includes(',')) ? `"${s.replace(/"/g, '""')}"` : s;
}

/** マークダウンパース */
function _applyMarkdownStyles(body, rawAiText) {
  if (!rawAiText) return;
  const lines = rawAiText.replace(/^\uFEFF/, "").split('\n');
  lines.forEach(line => {
    try {
      const trim = line.trim();
      if (trim.startsWith("---") || trim === "") {
        if (trim.startsWith("---")) body.appendHorizontalRule();
        else body.appendParagraph("");
        return;
      }
      
      let plain = trim;
      let head = null;
      if (plain.startsWith("# ")) { head = DocumentApp.ParagraphHeading.TITLE; plain = plain.substring(2); }
      else if (plain.startsWith("## ")) { head = DocumentApp.ParagraphHeading.HEADING1; plain = plain.substring(3); }
      else if (plain.startsWith("### ")) { head = DocumentApp.ParagraphHeading.HEADING2; plain = plain.substring(4); }
      else if (plain.startsWith("#### ")) { head = DocumentApp.ParagraphHeading.HEADING3; plain = plain.substring(5); }
      
      let p;
      if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        p = body.appendListItem(line.replace(/^\s*[-*] /, "").trim()).setGlyphType(DocumentApp.GlyphType.BULLET);
      } else {
        p = body.appendParagraph(plain);
        if (head) p.setHeading(head).editAsText().setBold(true);
      }
      
      // 文中の **太字** を処理
      _renderBold(p);
    } catch (e) {}
  });
}

function _renderBold(paragraph) {
  const text = paragraph.editAsText();
  let content = text.getText();
  const regex = /\*\*(.*?)\*\*/g;
  let match;
  let offset = 0;

  while ((match = regex.exec(content)) !== null) {
    const full = match[0];
    const inner = match[1];
    const start = match.index - offset;
    
    text.deleteText(start, start + full.length - 1);
    text.insertText(start, inner);
    text.setBold(start, start + inner.length - 1, true);
    
    offset += (full.length - inner.length);
    content = text.getText();
    regex.lastIndex = start + inner.length;
  }
}