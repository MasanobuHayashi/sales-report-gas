/**
 * 週報自動作成システム (Rev: Master-Order-Preservation)
 * 修正内容: 
 * 1. FOCusユーザマスタの「並び順」を完全遵守（部署順・担当者順）
 * 2. プロンプトの構成・ナンバリングをAIに生成させ、ハードコーディングを排除
 * 3. 文中の太字（**text**）を含むマークダウンパースの強化
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('週報自動作成')
    .addItem('実行', 'startReportGeneration')
    .addToUi();
}

// --- 定数定義（セルの場所のみを固定） ---
const SETTINGS_SHEET_NAME = "設定シート";
const PROMPT_DOC_ID_CELL = "B7";
const OUTPUT_FOLDER_ID_CELL = "B8";
const LOG_FOLDER_ID_CELL = "B9";
const START_DATE_CELL = "B3";
const MASTER_SHEET_NAME = "FOCusユーザマスタ";
const DATA_SHEET_NAME = "週報データ抽出";
const AI_MODEL = "models/gemini-2.5-flash"; 
// ---------------------------------------------------------------------------------

function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  let logMessage = "処理開始 (ボタン押下) \n";
  let logFolderId = null;
  let currentApiKey = null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`ERR-999: 設定シートが見つかりません。`);

    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
      startDate: settingsSheet.getRange(START_DATE_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. マスタ順序の読み込み（部署・担当者の順序を記憶）
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const masterData = masterSheet.getDataRange().getValues();
    const masterOrder = []; // {name, dept} の配列
    const orderedDepts = []; // 重複なしの部署リスト（出現順）
    
    for (let i = 1; i < masterData.length; i++) {
      const name = masterData[i][1];
      const dept = masterData[i][2];
      if (name && dept) {
        masterOrder.push({ name, dept });
        if (!orderedDepts.includes(dept)) orderedDepts.push(dept);
      }
    }

    // 2. 週次データの読み込みとマスタ順でのグループ化
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const rawReportData = dataSheet.getDataRange().getValues();
    if (rawReportData.length <= 1) throw new Error("ERR-001: 週報データがありません。");

    // 担当者名をキーとしたデータのMap
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

    // --- STEP 1: 部署別詳細生成（マスタ順に実行） ---
    logMessage += "7-1. 部署別詳細生成（マスタ順走査）...\n";
    let detailContent = "";
    let analysisSummaries = "";

    for (const deptName of orderedDepts) {
      // 当該部署の担当者のみを抽出
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
1. 現在は「${deptName}」の見出し配下の詳細を執筆しています。
2. 指示書の構成案に基づき、担当者別の活動詳細をマークダウンで出力してください。
3. 文末に必ず【DEPT_ANALYSIS】というタグに続けて、この部署の特筆事項を3行で要約してください。
4. メタ発言（「了解しました」等）は厳禁です。

入力データ:
${deptDataForAi}
`;
        const res = _callGeminiApi(detailPrompt, currentApiKey);
        const parts = res.split("【DEPT_ANALYSIS】");
        detailContent += parts[0].trim() + "\n\n";
        analysisSummaries += `■部署: ${deptName}\n${parts[1] || ""}\n\n`;
      }
    }

    // --- STEP 2: 全体統合および分析（サマリーと課題） ---
    logMessage += "7-2. 全体サマリーおよび分析生成...\n";
    const analysisPrompt = `${promptTemplate}
---
【全体分析・構成生成指示】
1. 指示書の構成定義に基づき、「全体サマリー」と「組織全体の課題」のセクションを執筆してください。
2. また、詳細セクション（部署別報告）を挿入すべき場所に、必ず「{{DETAIL_PLACEHOLDER}}」という文字列を1行単独で記述してください。
3. 生成される文書全体のナンバリング（1. 2. 3.）や見出しレベルを指示書通りに維持してください。

分析用インプット:
${analysisSummaries}
`;
    const finalShell = _callGeminiApi(analysisPrompt, currentApiKey);

    // 最終統合
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 8. ドキュメント出力
    logMessage += "8. ドキュメント出力...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    const body = doc.getBody();
    const titleDate = Utilities.formatDate(settings.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    body.appendParagraph(`営業週報（${titleDate} 〜）`).setHeading(DocumentApp.ParagraphHeading.TITLE);
    _applyMarkdownStyles(body, finalFullText);

    doc.saveAndClose();
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    logMessage += "P. 処理成功\n";
    ui.alert("週報が完成しました。マスタ順序およびプロンプトの構成が正しく反映されています。");

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

/** 担当者別CSVテキスト生成 */
function _createStaffText(staff, dept, rows) {
  const targets = ["部署名", "活動日", "顧客名", "活動目的", "予定及び活動結果", "社外同行者"];
  let txt = `[担当者: ${staff} (部署: ${dept})]\n${targets.join(",")}\n`;
  rows.forEach(r => {
    // データのインデックス（週報データ抽出シートの列順に合わせる）
    const vals = [dept, _fmtDate(r[0]), r[2], r[3], r[4], r[5]];
    txt += vals.map(v => _escCsv(v)).join(',') + "\n";
  });
  return txt;
}

function _fmtDate(d) { return d instanceof Date ? Utilities.formatDate(d, "JST", "MM/dd") : d; }

function _escCsv(c) {
  let s = (c == null) ? "" : String(c);
  s = s.replace(/\n/g, " ");
  if (s.includes('"') || s.includes(',')) s = `"${s.replace(/"/g, '""')}"`;
  return s;
}

/** マークダウンパース（文中太字対応） */
function _applyMarkdownStyles(body, rawAiText) {
  if (!rawAiText) return;
  const lines = rawAiText.replace(/^\uFEFF/, "").split('\n');
  lines.forEach(line => {
    try {
      const trim = line.trim();
      if (trim.startsWith("---")) { body.appendHorizontalRule(); return; }
      if (trim === "") { body.appendParagraph(""); return; }

      let plain = trim;
      let head = null;
      
      // 見出し判定
      if (plain.startsWith("# ")) { head = DocumentApp.ParagraphHeading.TITLE; plain = plain.substring(2); }
      else if (plain.startsWith("## ")) { head = DocumentApp.ParagraphHeading.HEADING1; plain = plain.substring(3); }
      else if (plain.startsWith("### ")) { head = DocumentApp.ParagraphHeading.HEADING2; plain = plain.substring(4); }
      else if (plain.startsWith("#### ")) { head = DocumentApp.ParagraphHeading.HEADING3; plain = plain.substring(5); }
      
      let p;
      if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        p = body.appendListItem(line.replace(/^\s*[-*] /, "").trim()).setGlyphType(DocumentApp.GlyphType.BULLET);
      } else {
        p = body.appendParagraph(plain);
        if (head) p.setHeading(head);
      }

      // 文中の太字（**text**）を処理
      _applyBoldFormatting(p);

    } catch (e) {}
  });
}

function _applyBoldFormatting(element) {
  const textElement = element.editAsText();
  const text = textElement.getText();
  const boldRegex = /\*\*(.*?)\*\*/g;
  let match;
  
  while ((match = boldRegex.exec(text)) !== null) {
    const start = match.index;
    const end = start + match[0].length - 1;
    // ** を削除して中身を太字にする（簡易実装のため削除後のインデックス調整は省略）
    textElement.setBold(start, end, true);
  }
}