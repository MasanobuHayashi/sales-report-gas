/**
 * 週報自動作成システム (Rev: Performance-and-Order-Optimization)
 * 修正内容: 
 * 1. 6分制限回避のため、第2段階への入力データを「部署別要約のみ」に軽量化（タイムアウト対策）
 * 2. プロンプト指示に合わせ、出力順序を「タイトル > サマリー > 詳細 > 組織課題」に固定
 * 3. 2段階生成のフローを最適化し、AIの推論負荷を軽減
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`ERR-999: シート '${SETTINGS_SHEET_NAME}' が見つかりません。`);

    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
      startDate: settingsSheet.getRange(START_DATE_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const masterData = masterSheet.getDataRange().getValues();
    const reportData = dataSheet.getDataRange().getValues();
    if (reportData.length <= 1) throw new Error("ERR-001: 週報データがありません。");

    const departmentMap = new Map();
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][1]) departmentMap.set(masterData[i][1], masterData[i][2]);
    }

    const deptGroups = new Map();
    for (let i = 1; i < reportData.length; i++) {
      const row = reportData[i];
      const deptName = departmentMap.get(row[1]) || "不明";
      if (!deptGroups.has(deptName)) deptGroups.set(deptName, [reportData[0]]);
      deptGroups.get(deptName).push([...row, deptName]);
    }

    let promptTemplate;
    try {
      promptTemplate = DocumentApp.openById(settings.promptDocId).getBody().getText();
    } catch (e) {
      throw new Error(`ERR-201: プロンプト雛形読込エラー。`);
    }

    currentApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!currentApiKey) throw new Error("APIキー未設定。");

    // --- STEP 1: 部署別詳細レポートの生成（要約ポイントも同時に抽出） ---
    logMessage += "7-1. 部署別詳細生成（タイムアウト対策：軽量化）...\n";
    let detailContent = "";
    let summaryForAnalysis = "";
    const sortedDepts = Array.from(deptGroups.keys()).sort();

    for (const deptName of sortedDepts) {
      logMessage += `  -> ${deptName} 生成中...\n`;
      const deptData = deptGroups.get(deptName);
      const inputDataText = _createGroupedInputText(deptData);
      
      const detailPrompt = `${promptTemplate}
---
【重要：部署別レポート指示】
1. 現在は「${deptName}」のセクションを執筆しています。
2. タイトルや全体サマリーは含めず、「### ${deptName}」から始まる担当者別の活動詳細を漏れなく記述してください。
3. 文末に、次の分析ステップのために、この部署の活動のハイライトを【要約: ${deptName}】というタグに続けて、3行程度で箇条書きしてください。

入力データ:
${inputDataText}
`;
      const fullResponse = _callGeminiApi(detailPrompt, currentApiKey);
      
      // 要約タグの部分を抽出してサマリー用変数へ、それ以外を詳細用変数へ
      const parts = fullResponse.split(/【要約: .*】/);
      detailContent += parts[0] + "\n\n";
      if (parts.length > 1) {
        summaryForAnalysis += `■${deptName}のハイライト:\n${parts[1].trim()}\n\n`;
      }
    }

    // --- STEP 2: 全体要約と組織課題の生成（軽量化されたデータを使用） ---
    logMessage += "7-2. 全体サマリー・組織課題生成（高速モード）...\n";
    const analysisPrompt = `${promptTemplate}
---
【重要：全体分析指示】
1. 各部署から届いた以下の「ハイライト」に基づき、プロンプト指示にある「全体サマリー」と「組織全体の課題と次週の重点事項」の2セクションのみを執筆してください。
2. シニアセールスマネージャーの視点で深く分析してください。
3. 担当者別の詳細は含めないでください。

入力データ（全部署のハイライト）:
${summaryForAnalysis}
`;
    const analysisResult = _callGeminiApi(analysisPrompt, currentApiKey);

    // 8. ドキュメント出力と構成（順序の適正化）
    logMessage += "8. ドキュメント出力（順序適正化）...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();
    const titleDate = Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // 構成の再定義： 1.タイトル > 2.サマリー（分析結果の前半） > 3.詳細 > 4.組織課題（分析結果の後半）
    body.appendParagraph(`営業週報（${titleDate} 〜）`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    
    // 分析結果を「サマリー」と「課題」に分割して配置
    const analysisParts = analysisResult.split(/## 組織全体の課題と次週の重点事項/);
    const summaryText = analysisParts[0];
    const issuesText = analysisParts.length > 1 ? "## 組織全体の課題と次週の重点事項\n" + analysisParts[1] : "";

    _applyMarkdownStyles(body, summaryText);   // 2. 全体サマリー
    _applyMarkdownStyles(body, detailContent); // 3. 部署別詳細（138名分）
    _applyMarkdownStyles(body, issuesText);   // 4. 組織全体の課題

    doc.saveAndClose();
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    logMessage += "P. 処理成功\n";
    ui.alert("週報が完成しました。6分制限を回避し、正しい順序で構成されています。");

  } catch (e) {
    let safeErrorMessage = currentApiKey ? e.message.split(currentApiKey).join("********") : e.message;
    logMessage += `Z. エラー: ${safeErrorMessage}\n`;
    ui.alert(`エラーが発生しました。ログを確認してください。\n${safeErrorMessage}`);
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
  throw new Error(`APIエラー (HTTP ${res.getResponseCode()})`);
}

/** ログ書き出し */
function _writeLog(msg, fid) {
  try { DriveApp.getFolderById(fid).createFile(`log_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss")}.txt`, msg); } catch(e){}
}

/** CSVセルエスケープ */
function _escapeCsvCell(c) {
  let s = (c == null) ? "" : String(c);
  s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n").replace(/\n{3,}/g, "\n\n");
  if (s.includes('"') || s.includes(',') || s.includes('\n')) s = `"${s.replace(/"/g, '""')}"`;
  return s;
}

/** グループ化テキスト生成 */
function _createGroupedInputText(data) {
  const map = new Map();
  const targets = ["部署名", "活動日", "顧客名", "活動目的", "予定及び活動結果", "社外同行者"];
  const hMap = new Map();
  data[0].forEach((v, i) => hMap.set(v, i));
  const idxs = targets.map(t => hMap.get(t));
  for (let i = 1; i < data.length; i++) {
    const s = data[i][1];
    if (!map.has(s)) map.set(s, []);
    map.get(s).push(data[i]);
  }
  let txt = "";
  for (const [s, rows] of map.entries()) {
    txt += `[担当者: ${s}]\n${targets.join(",")}\n`;
    for (const r of rows) txt += idxs.map(i => _escapeCsvCell(r[i])).join(',') + "\n";
    txt += "\n";
  }
  return txt;
}

/** スタイル適用 */
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

      if (plainText.startsWith("# ")) { headingType = DocumentApp.ParagraphHeading.TITLE; plainText = plainText.substring(2); isBold = true; }
      else if (plainText.startsWith("## ")) { headingType = DocumentApp.ParagraphHeading.HEADING1; plainText = plainText.substring(3); isBold = true; }
      else if (plainText.startsWith("### ")) { headingType = DocumentApp.ParagraphHeading.HEADING2; plainText = plainText.substring(4); isBold = true; }
      else if (plainText.startsWith("#### ")) { headingType = DocumentApp.ParagraphHeading.HEADING3; plainText = plainText.substring(5); isBold = true; }
      else if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        const p = body.appendListItem(line.replace(/^\s*[-*] /, "").trim());
        p.setGlyphType(DocumentApp.GlyphType.BULLET);
        return;
      }
      const paragraph = body.appendParagraph(plainText);
      if (headingType) paragraph.setHeading(headingType);
      if (isBold && plainText.length > 0) paragraph.editAsText().setBold(true);
    } catch (e) {}
  });
}