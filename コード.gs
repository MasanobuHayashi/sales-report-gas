/**
 * 週報自動作成システム (Rev: 2026-Parallel-Architecture)
 * 準拠仕様書: Rev 1.3 + Parallel Processing Update
 * モデル: Gemini 2.5 Flash
 * [主な変更点]
 * - UrlFetchApp.fetchAll を使用した部署別データの並列生成（タイムアウト回避の切り札）
 */

// --- 定数定義 ---
const SETTINGS_SHEET_NAME = "設定シート";
const PROMPT_DOC_ID_CELL = "B7";
const OUTPUT_FOLDER_ID_CELL = "B8";
const LOG_FOLDER_ID_CELL = "B9";
const MASTER_SHEET_NAME = "FOCusユーザマスタ";
const DATA_SHEET_NAME = "週報データ抽出";
const AI_MODEL = "models/gemini-2.5-flash"; 
const MAX_PROMPT_SIZE_BYTES = 9 * 1024 * 1024;
// 並列処理にするため、タイムアウトマージンは不要ですが、念のため残します
const MAX_EXECUTION_TIME_MS = 340 * 1000; 
// ---------------------------------------------------------------------------------

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

    // 1. マスター順序の読み込み
    logMessage += "1. マスターデータ解析(順序定義)...\n";
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
    logMessage += `  -> 対象データ: ${rawReportData.length - 1}件\n`;

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

    // --- STEP 1: 並列処理の準備 (リクエスト構築) ---
    logMessage += "4. AI生成リクエスト構築 (並列準備)...\n";
    
    // リクエストを格納する配列
    const requestPayloads = [];
    // どのリクエストがどの部署かを紐付ける配列
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
        
        // サイズチェック
        if (Utilities.newBlob(detailPrompt).getBytes().length > MAX_PROMPT_SIZE_BYTES) {
           logMessage += `  WARN: ${deptName} サイズ超過のためスキップ\n`;
           continue; 
        }

        // fetchAll用のリクエストオブジェクトを作成
        requestPayloads.push({
          'url': API_URL,
          'method': 'post',
          'contentType': 'application/json',
          'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": detailPrompt }] }] }),
          'muteHttpExceptions': true
        });
        requestDepts.push(deptName);
      }
    }

    // --- 並列実行 (The Critical Moment) ---
    logMessage += `  -> ${requestPayloads.length}部署分を一括並列送信中...\n`;
    let detailContent = "";
    let analysisSummaries = "";

    // ★ここで一斉送信 (fetchAll)★
    const responses = UrlFetchApp.fetchAll(requestPayloads);

    logMessage += "  -> 全レスポンス受信完了。解析開始...\n";

    // レスポンスの解析
    for (let i = 0; i < responses.length; i++) {
      const deptName = requestDepts[i];
      const res = responses[i];
      const resCode = res.getResponseCode();

      if (resCode === 200) {
        try {
          const json = JSON.parse(res.getContentText());
          const aiText = json.candidates?.[0]?.content?.parts?.[0]?.text;
          if (aiText) {
            const parts = aiText.split("【DEPT_SUMMARY】");
            detailContent += parts[0].trim() + "\n\n";
            analysisSummaries += `■部署: ${deptName}\n${parts[1] || "(サマリーなし)"}\n\n`;
          } else {
             logMessage += `  ERR: ${deptName} 応答空\n`;
             detailContent += `\n### ${deptName}\n(AI生成応答なし)\n\n`;
          }
        } catch (e) {
          logMessage += `  ERR: ${deptName} JSONパース失敗\n`;
          detailContent += `\n### ${deptName}\n(システムエラー)\n\n`;
        }
      } else {
        logMessage += `  ERR: ${deptName} API Error ${resCode}\n`;
        detailContent += `\n### ${deptName}\n(AI通信エラー: ${resCode})\n\n`;
      }
    }

    // --- STEP 2: 全体統合 (ここは単発実行) ---
    logMessage += "5. 全体要約生成 (統合)...\n";
    const analysisPrompt = `${promptTemplate}\n---\n【全体統合指示】\n1. タイトル日付を「${dateRangeStr}」としてください。\n2. {{DETAIL_PLACEHOLDER}} の位置に詳細を結合します。\n\n分析用インプット:\n${analysisSummaries}`;
    
    // 単発呼び出し用の関数を使用
    const finalShell = _callGeminiApiSingle(analysisPrompt, currentApiKey);
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 4. 出力
    logMessage += "6. ファイル出力...\n";
    
    // 【修正】ファイル名の日付を minDate から maxDate に変更
    const fileName = "週報_" + Utilities.formatDate(maxDate || new Date(), tz, "yyyy-MM-dd");
  
    const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
    const existingFiles = outputFolder.getFilesByName(fileName);
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);
    
    const doc = DocumentApp.create(fileName);
    _applyMarkdownStyles(doc.getBody(), finalFullText);
    doc.saveAndClose();
    
    const file = DriveApp.getFileById(doc.getId());
    outputFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    logMessage += "処理完了: 成功\n";
    ui.alert("作成完了", "週報の自動作成が完了しました。", ui.ButtonSet.OK);

  } catch (e) {
    logMessage += `\n❌ 異常終了: ${e.message}\nStack: ${e.stack}\n`;
    const userMsg = e.message.startsWith("ERR-") ? e.message : `ERR-999: 予期せぬエラー\n(${e.message})`;
    ui.alert("エラー発生", `処理を中断しました。\n${userMsg}\n\nスクリーンショットを開発者へ送信してください。`, ui.ButtonSet.OK);
  } finally {
    if (logFolderId) {
      try {
        DriveApp.getFolderById(logFolderId).createFile(`log_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss")}.txt`, logMessage);
      } catch(e) { console.log(logMessage); }
    }
  }
}

// 単発実行用 (全体要約で使用)
function _callGeminiApiSingle(prompt, apiKey) {
  const API_URL = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`;
  const options = {
    'method': 'post', 'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true
  };
  const res = UrlFetchApp.fetch(API_URL, options); // 単発は fetch のまま
  if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
  throw new Error(`API Error: ${res.getResponseCode()}`);
}

// ヘルパー関数 (変更なし)
function _createStaffText(staff, dept, rows) {
  let txt = `[担当者: ${staff} (部署: ${dept})]\n`;
  rows.forEach(r => {
    let dStr = (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd") : String(r[0]).substring(0, 10);
    txt += `- ${dStr} ${r[2]} / ${r[4]}\n`;
  });
  return txt;
}

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
      if (head) p.setHeading(head);
    }
  });
}