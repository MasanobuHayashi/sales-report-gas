/**
 * 週報自動作成システム (Rev: 2026-Stable-Refactored)
 * 準拠仕様書: Rev 1.3
 * モデル: Gemini 2.5 Flash
 * * [主な変更点]
 * - ログ保存機能の実装 (Driveへの出力)
 * - 独自エラーコード (ERR-xxx) による例外処理の標準化
 * - プロンプトサイズチェック (9MB制限)
 * - API呼び出しのリトライロジック追加
 * - ユーザー向けアラートUIの改善
 */

// --- 定数定義 ---
const SETTINGS_SHEET_NAME = "設定シート";
const PROMPT_DOC_ID_CELL = "B7";
const OUTPUT_FOLDER_ID_CELL = "B8";
const LOG_FOLDER_ID_CELL = "B9";
const MASTER_SHEET_NAME = "FOCusユーザマスタ";
const DATA_SHEET_NAME = "週報データ抽出";
const AI_MODEL = "models/gemini-2.5-flash"; 
const MAX_PROMPT_SIZE_BYTES = 9 * 1024 * 1024; // 9MB制限
const MAX_EXECUTION_TIME_MS = 340 * 1000; // GAS 6分制限に対する安全マージン(5分40秒)
// ---------------------------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('週報自動作成')
    .addItem('実行', 'startReportGeneration')
    .addToUi();
}

/**
 * メイン処理
 */
function startReportGeneration() {
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date().getTime();
  let logMessage = `処理開始: ${new Date().toLocaleString()}\n`;
  let logFolderId = null;
  let currentApiKey = null;
  let isSuccess = false;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // シート読み込み待機時間を考慮してFlush
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

    if (!settings.promptDocId || !settings.outputFolderId || !settings.logFolderId) {
       throw new Error("ERR-200: 設定シートのID指定に不備があります。");
    }

    // 1. マスター順序の読み込み
    logMessage += "1. マスターデータ解析(順序定義)...\n";
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
    if (!masterSheet) throw new Error(`ERR-999: シート '${MASTER_SHEET_NAME}' が見つかりません。`);
    
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
    logMessage += `  -> 部署数: ${orderedDepts.length}, 担当者数: ${masterOrder.length}\n`;

    // 2. 週次データの取得と日付範囲算出
    logMessage += "2. 週次データ取得...\n";
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    if (!dataSheet) throw new Error(`ERR-999: シート '${DATA_SHEET_NAME}' が見つかりません。`);

    const rawReportData = dataSheet.getDataRange().getValues();
    // ヘッダーのみの場合はデータなしとみなす
    if (rawReportData.length <= 1) throw new Error("ERR-001: 週報データがありません。抽出条件またはQUERY関数を確認してください。");

    let minDate = null;
    let maxDate = null;
    const dataByStaff = new Map();

    for (let i = 1; i < rawReportData.length; i++) {
      const row = rawReportData[i];
      const dateVal = row[0]; // A列: 活動日
      const staff = row[1];   // B列: 担当者

      if (dateVal instanceof Date) {
        if (!minDate || dateVal < minDate) minDate = dateVal;
        if (!maxDate || dateVal > maxDate) maxDate = dateVal;
      }

      if (!dataByStaff.has(staff)) dataByStaff.set(staff, []);
      dataByStaff.get(staff).push(row);
    }
    logMessage += `  -> データ件数: ${rawReportData.length - 1}件\n`;

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
      throw new Error(`ERR-201: プロンプト雛形読込エラー (ID: ${settings.promptDocId})。アクセス権限を確認してください。`);
    }
    const promptTemplate = promptFull.split("▽ メンテナンス担当者様へ")[0].trim();

    // APIキー取得
    currentApiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!currentApiKey) throw new Error("ERR-100: APIキー未設定。スクリプトプロパティ 'GEMINI_API_KEY' を設定してください。");

    // --- STEP 1: 部署別詳細生成 ---
    logMessage += "4. AI生成開始 (部署別)...\n";
    let detailContent = "";
    let analysisSummaries = "";

    for (const deptName of orderedDepts) {
      // タイムアウト安全装置
      if (new Date().getTime() - startTime > MAX_EXECUTION_TIME_MS) {
        throw new Error("ERR-101: 処理時間が上限(約6分)に達したため、安全に中断しました。対象期間を短くして再実行してください。");
      }

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
        const blobSize = Utilities.newBlob(detailPrompt).getBytes().length;
        if (blobSize > MAX_PROMPT_SIZE_BYTES) {
           logMessage += `  WARN: ${deptName}のデータサイズ過大 (${blobSize} bytes). スキップします。\n`;
           detailContent += `\n\n### ${deptName}\n(データ量超過のため生成スキップ)\n\n`;
           continue; 
        }

        logMessage += `  Generating: ${deptName}...\n`;
        try {
          const res = _callGeminiApiWithRetry(detailPrompt, currentApiKey);
          const parts = res.split("【DEPT_SUMMARY】");
          detailContent += parts[0].trim() + "\n\n";
          analysisSummaries += `■部署: ${deptName}\n${parts[1] || "(サマリー生成なし)"}\n\n`;
        } catch (apiErr) {
          logMessage += `  ERROR in ${deptName}: ${apiErr.message}\n`;
          detailContent += `\n\n### ${deptName}\n(AI生成エラー: ${apiErr.message})\n\n`;
        }
      }
    }

    // --- STEP 2: 全体統合 ---
    logMessage += "5. AI生成 (全体要約)...\n";
    const analysisPrompt = `${promptTemplate}\n---\n【全体統合指示】\n1. タイトル日付を「${dateRangeStr}」としてください。\n2. {{DETAIL_PLACEHOLDER}} の位置に詳細を結合します。\n\n分析用インプット:\n${analysisSummaries}`;
    
    // サイズチェック (全体)
    if (Utilities.newBlob(analysisPrompt).getBytes().length > MAX_PROMPT_SIZE_BYTES) {
      throw new Error("ERR-002: 全体要約プロンプトのサイズが超過しました。");
    }

    const finalShell = _callGeminiApiWithRetry(analysisPrompt, currentApiKey);
    const finalFullText = finalShell.replace("{{DETAIL_PLACEHOLDER}}", detailContent);

    // 4. 出力
    logMessage += "6. ファイル出力処理...\n";
    const fileName = "週報_" + Utilities.formatDate(minDate || new Date(), tz, "yyyy-MM-dd");
    
    // 既存ファイルのクリーンアップ（同名ファイルはゴミ箱へ）
    try {
      const outputFolder = DriveApp.getFolderById(settings.outputFolderId);
      const existingFiles = outputFolder.getFilesByName(fileName);
      while (existingFiles.hasNext()) {
        existingFiles.next().setTrashed(true);
      }
      
      const doc = DocumentApp.create(fileName);
      _applyMarkdownStyles(doc.getBody(), finalFullText);
      doc.saveAndClose();
      
      const file = DriveApp.getFileById(doc.getId());
      outputFolder.addFile(file);
      DriveApp.getRootFolder().removeFile(file); // ルートから削除
      
      logMessage += `  -> ファイル生成成功: ${fileName}\n`;
    } catch (driveErr) {
      throw new Error(`ERR-200: ファイル出力エラー。フォルダIDを確認してください。(${driveErr.message})`);
    }

    isSuccess = true;
    logMessage += "処理完了: 成功\n";
    ui.alert("作成完了", "週報の自動作成が完了しました。", ui.ButtonSet.OK);

  } catch (e) {
    // エラーハンドリング
    logMessage += `\n❌ 異常終了: ${e.message}\nStack: ${e.stack}\n`;
    
    let userMsg = e.message;
    if (!userMsg.startsWith("ERR-")) {
      userMsg = `ERR-999: 予期せぬエラー\n(${userMsg})`;
    }

    const alertText = `処理を中断しました。\n\n【エラー内容】\n${userMsg}\n\n--------------------------------------------------\nこの画面のスクリーンショットを撮り、\nシステム開発者へお問い合わせください。\n--------------------------------------------------\n※詳細はログフォルダをご確認ください。`;
    ui.alert("エラー発生", alertText, ui.ButtonSet.OK);

  } finally {
    // ログ出力 (必ず実行)
    if (logFolderId) {
      try {
        const logFileName = `log_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss")}.txt`;
        DriveApp.getFolderById(logFolderId).createFile(logFileName, logMessage);
      } catch (logErr) {
        console.error("ログ保存失敗: " + logErr.message);
        console.log(logMessage); // 最低限コンソールには残す
      }
    } else {
      console.log(logMessage);
    }
  }
}

/**
 * Gemini API呼び出し (リトライロジック付き)
 */
function _callGeminiApiWithRetry(prompt, apiKey, maxRetries = 3) {
  const API_URL = `https://generativelanguage.googleapis.com/v1beta/${AI_MODEL}:generateContent?key=${apiKey}`;
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({ "contents": [{ "parts": [{ "text": prompt }] }] }),
    'muteHttpExceptions': true
  };

  for (let i = 0; i < maxRetries; i++) {
    try {
      const res = UrlFetchApp.fetch(API_URL, options);
      const resCode = res.getResponseCode();
      
      if (resCode === 200) {
        const json = JSON.parse(res.getContentText());
        return json.candidates?.[0]?.content?.parts?.[0]?.text || "";
      } else if (resCode === 429 || resCode >= 500) {
        // レート制限(429)またはサーバーエラー(5xx)の場合はリトライ
        Utilities.sleep(1000 * Math.pow(2, i)); // 指数バックオフ
        continue;
      } else {
        throw new Error(`API Error (HTTP ${resCode}): ${res.getContentText().substring(0, 200)}`);
      }
    } catch (e) {
      if (i === maxRetries - 1) throw new Error(`ERR-100: AI通信エラー (Retry limit exceeded). ${e.message}`);
      Utilities.sleep(1000);
    }
  }
}

// --- ヘルパー関数 ---

function _createStaffText(staff, dept, rows) {
  let txt = `[担当者: ${staff} (部署: ${dept})]\n`;
  rows.forEach(r => {
    // 日付フォーマット
    let dStr = r[0];
    if (r[0] instanceof Date) {
      dStr = Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd");
    } else {
      dStr = String(r[0]).substring(0, 10);
    }
    // r[2]=顧客名, r[4]=予定及び活動結果
    txt += `- ${dStr} ${r[2]} / ${r[4]}\n`;
  });
  return txt;
}

function _applyMarkdownStyles(body, rawAiText) {
  if (!rawAiText) return;
  // BOM除去と整形
  const lines = rawAiText.replace(/^\uFEFF/, "").split('\n');
  
  lines.forEach(line => {
    let plain = line.trim();
    if (plain === "") {
      body.appendParagraph(""); // 空行を維持
      return;
    }

    let head = null;
    if (plain.startsWith("# ")) { head = DocumentApp.ParagraphHeading.TITLE; plain = plain.substring(2); }
    else if (plain.startsWith("## ")) { head = DocumentApp.ParagraphHeading.HEADING1; plain = plain.substring(3); }
    else if (plain.startsWith("### ")) { head = DocumentApp.ParagraphHeading.HEADING2; plain = plain.substring(4); }
    else if (plain.startsWith("#### ")) { head = DocumentApp.ParagraphHeading.HEADING3; plain = plain.substring(5); }

    if (line.match(/^(\s*)- /) || line.match(/^(\s*)\* /)) {
      // リストアイテム
      const listItem = body.appendListItem(plain.replace(/^[-*]\s+/, ""));
      listItem.setGlyphType(DocumentApp.GlyphType.BULLET);
    } else {
      // 通常段落または見出し
      const p = body.appendParagraph(plain);
      if (head) {
        p.setHeading(head);
      }
    }
  });
}