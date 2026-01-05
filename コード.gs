/**
 * 週報自動作成システム (Rev: Final-Stability-and-Order-Optimization)
 * 修正内容: 
 * 1. SpreadsheetApp.flush() の導入により数式計算待ちによる ERR-001 を解消
 * 2. プロンプト指示の日付プレースホルダ置換とタイトル出力をAIに完全委譲
 * 3. FOCusユーザマスタの「行の並び順」を絶対基準として部署・担当者を出力
 * 4. 文中太字(**...**)のパースおよびスタイリングの不整合を修正
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('週報自動作成')
    .addItem('実行', 'startReportGeneration')
    .addToUi();
}

// --- 定数定義 ---
const SETTINGS_SHEET_NAME = "設定シート";
const PROMPT_DOC_ID_CELL = "B6";   // B7からB6へ修正
const OUTPUT_FOLDER_ID_CELL = "B7"; // B8からB7へ修正
const LOG_FOLDER_ID_CELL = "B8";   // B9からB8へ修正
const START_DATE_CELL = "B2";      // B3からB2へ修正
const END_DATE_CELL = "B3";        // B4からB3へ修正
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
    // 0. スプレッドシートの数式計算を強制同期 (ERR-001対策)
    SpreadsheetApp.flush();

    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`ERR-999: シート '${SETTINGS_SHEET_NAME}' が見つかりません。`);

    // 日付取得の堅牢化
    logMessage += "0. 日付情報の検証...\n";
    const startDateVal = settingsSheet.getRange(START_DATE_CELL).getValue();
    const endDateVal = settingsSheet.getRange(END_DATE_CELL).getValue();
    
    // Dateオブジェクトへの変換
    const startDate = (startDateVal instanceof Date) ? startDateVal : new Date(settingsSheet.getRange(START_DATE_CELL).getDisplayValue());
    const endDate = (endDateVal instanceof Date) ? endDateVal : new Date(settingsSheet.getRange(END_DATE_CELL).getDisplayValue());

    if (!startDate || isNaN(startDate.getTime()) || !endDate || isNaN(endDate.getTime())) {
      throw new Error(`ERR-001: 設定シートの日付 [${START_DATE_CELL}/${END_DATE_CELL}] が取得できません。スプレッドシートの再計算を確認してください。`);
    }

    const tz = Session.getScriptTimeZone();
    const dateRangeStr = `${Utilities.formatDate(startDate, tz, "yyyy年MM月dd日")}～${Utilities.formatDate(endDate, tz, "yyyy年MM月dd日")}`;

    const settings = {
      promptDocId: settingsSheet.getRange(PROMPT_DOC_ID_CELL).getValue(),
      outputFolderId: settingsSheet.getRange(OUTPUT_FOLDER_ID_CELL).getValue(),
      logFolderId: settingsSheet.getRange(LOG_FOLDER_ID_CELL).getValue(),
    };
    logFolderId = settings.logFolderId;

    // 1. マスタ順序（部署・担当者）の読み込み
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    const masterData = masterSheet.getDataRange().getValues();
    const masterOrder = []; // {name, dept} の配列
    const orderedDepts = []; // 出現順の部署リスト
    for (let i = 1; i < masterData.length; i++) {
      const name = masterData[i][1];
      const dept = masterData[i][2];
      if (name && dept) {
        masterOrder.push({ name, dept });
        if (!orderedDepts.includes(dept)) orderedDepts.push(dept);
      }
    }

    // 2. 週次データの抽出
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

    // --- STEP 1: 部署別詳細生成（マスタ順） ---
    logMessage += "7-1. 部署別詳細生成（マスタ順走査）...\n";
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
2. 指示書の構成定義に基づき、担当者別の活動詳細を出力してください。
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

    // --- STEP 2: 全体要約および構成生成 ---
    logMessage += "7-2. 全体構成および分析生成（日付連携）...\n";
    const analysisPrompt = `${promptTemplate}
---
【全体要約・構成生成指示】
1. プロンプトの構成定義に従い、「タイトル」「全体サマリー」「組織全体の課題」セクションを執筆してください。
2. タイトル内の日付プレースホルダは、必ず「${dateRangeStr}」に書き換えてください。
3. 詳細セクション（部署別報告）の挿入位置に、単独行で「{{DETAIL_PLACEHOLDER}}」と記述してください。
4. ナンバリング（1. 2. 3.）や書式を指示書通りに維持してください。

分析用インプット:
${analysisSummaries}
`;
    const finalShell = _callGeminiApi(analysisPrompt, currentApiKey);
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 8. ドキュメント出力
    logMessage += "8. ドキュメント出力...\n";
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const fileName = "週報_" + Utilities.formatDate(startDate, tz, "yyyy-MM-dd");
    const existing = outputFolder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const doc = DocumentApp.create(fileName);
    _applyMarkdownStyles(doc.getBody(), finalFullText);

    doc.saveAndClose();
    outputFolder.addFile(DriveApp.getFileById(doc.getId()));
    DriveApp.getRootFolder().removeFile(DriveApp.getFileById(doc.getId()));

    logMessage += "P. 処理成功\n";
    ui.alert("週報が完成しました。");

  } catch (e) {
    let safeErr = currentApiKey ? e.message.split(currentApiKey).join("********") : e.message;
    logMessage += `Z. エラー: ${safeErr}\n`;
    ui.alert(`エラーが発生しました。\n\n${safeErr}`);
  } finally {
    if (logFolderId) _writeLog(logMessage, logFolderId);
  }
}

/** Gemini API呼び出し */
function _callGeminiApi(prompt, apiKey) {
  const options = {
    'method': 'post', 'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true, 'timeoutSeconds': 300
  };
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`, options);
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
    txt += [dept, dStr, r[2], r[3], r[4], r[5]].map(v => {
      let s = (v == null) ? "" : String(v).replace(/\n/g, " ");
      return (s.includes('"') || s.includes(',')) ? `"${s.replace(/"/g, '""')}"` : s;
    }).join(',') + "\n";
  });
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
      
      let p;
      if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
        p = body.appendListItem(line.replace(/^\s*[-*] /, "").trim()).setGlyphType(DocumentApp.GlyphType.BULLET);
      } else {
        p = body.appendParagraph(plain);
        if (head) p.setHeading(head).editAsText().setBold(true);
      }
      
      // 太字(**...**)のレンダリング
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
    const full = match[0]; const inner = match[1];
    const start = match.index - offset;
    text.deleteText(start, start + full.length - 1);
    text.insertText(start, inner);
    text.setBold(start, start + inner.length - 1, true);
    offset += (full.length - inner.length);
    content = text.getText();
    regex.lastIndex = start + inner.length;
  }
}