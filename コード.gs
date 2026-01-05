/**
 * 週報自動作成システム (Rev: Final-Structure-Optimization)
 * 修正内容: 
 * 1. 出力順序をプロンプト通りに固定（サマリー > 詳細 > 組織課題）
 * 2. AIのメタ発言やタグの混入を徹底排除
 * 3. 2段階生成のデータ連携をより強固に改善
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

    // --- STEP 1: 部署別詳細生成 ---
    logMessage += "7-1. 部署別詳細生成（詳細＋分析用要約の抽出）...\n";
    let detailContent = "";
    let summaryForAnalysis = "";
    const sortedDepts = Array.from(deptGroups.keys()).sort();

    for (const deptName of sortedDepts) {
      logMessage += `  -> ${deptName} 生成中...\n`;
      const deptData = deptGroups.get(deptName);
      const inputDataText = _createGroupedInputText(deptData);
      
      const detailPrompt = `${promptTemplate}
---
【部署別指示（厳守）】
1. 現在は「${deptName}」の詳細報告セクションを作成しています。
2. 「### ${deptName}」の見出しから始め、担当者別の活動を全て漏れなく詳細に記述してください。
3. 文末に必ず【ANALYSIS_START】というタグを入れ、続けてこの部署の「成果・傾向・共通の課題」を組織分析用に5行程度で抽出してください。
4. メタ発言（「はい」「作成します」等）は一切禁止します。

入力データ:
${inputDataText}
`;
      const response = _callGeminiApi(detailPrompt, currentApiKey);
      const resParts = response.split("【ANALYSIS_START】");
      detailContent += resParts[0].trim() + "\n\n";
      if (resParts.length > 1) {
        summaryForAnalysis += `■部署: ${deptName}\n${resParts[1].trim()}\n\n`;
      }
    }

    // --- STEP 2: 全体サマリーと組織課題の生成 ---
    logMessage += "7-2. 全体サマリー・組織課題生成（エグゼクティブ・ビュー）...\n";
    const summaryPrompt = `${promptTemplate}
---
【全体要約指示（厳守）】
1. プロンプトの構成定義のうち、「## 全体サマリー」セクションのみを執筆してください。
2. 他のセクション（詳細や課題）は含めないでください。
3. 以下の全部署のハイライトを基に、シニアマネージャーとして俯瞰的な洞察を述べてください。

入力データ:
${summaryForAnalysis}
`;

    const issuesPrompt = `${promptTemplate}
---
【組織課題指示（厳守）】
1. プロンプトの構成定義のうち、「## 組織全体の課題と次週の重点事項」セクションのみを執筆してください。
2. 他のセクションは含めないでください。
3. 具体的なボトルネックや解決策を提示してください。

入力データ:
${summaryForAnalysis}
`;

    const summarySection = _callGeminiApi(summaryPrompt, currentApiKey);
    const issuesSection = _callGeminiApi(issuesPrompt, currentApiKey);

    // 8. ドキュメント出力（順序を完全に制御）
    logMessage += "8. ドキュメント統合出力...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();
    const titleDate = Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // 統合構成: 1.タイトル > 2.サマリー > 3.部署別詳細 > 4.組織課題
    body.appendParagraph(`営業週報（${titleDate} 〜）`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    _applyMarkdownStyles(body, summarySection); // 2. 全体サマリー
    _applyMarkdownStyles(body, detailContent);  // 3. 部署別詳細
    _applyMarkdownStyles(body, issuesSection);  // 4. 組織全体の課題

    doc.saveAndClose();
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    logMessage += "P. 処理成功\n";
    ui.alert("週報が完成しました。順序・品質ともに指示通りに調整済みです。");

  } catch (e) {
    let safeErr = currentApiKey ? e.message.split(currentApiKey).join("********") : e.message;
    logMessage += `Z. エラー: ${safeErr}\n`;
    ui.alert(`エラー: ${safeErr}`);
  } finally {
    if (logFolderId) _writeLog(logMessage, logFolderId);
  }
}

/** Gemini API呼び出し */
function _callGeminiApi(prompt, apiKey) {
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true,
    'timeoutSeconds': 300
  };
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`, options);
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  throw new Error(`APIエラー: ${res.getResponseCode()}`);
}

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
      else if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        body.appendListItem(line.replace(/^\s*[-*] /, "").trim()).setGlyphType(DocumentApp.GlyphType.BULLET);
        return;
      }
      const p = body.appendParagraph(plain);
      if (head) p.setHeading(head).editAsText().setBold(true);
    } catch (e) {}
  });
}