/**
 * @fileoverview 小学校週案作成支援スクリプト。単元計画PDFの読込、Gemini APIによる指導内容の自動生成機能を含む。
 * @version 2.2.0 - 2025-08-06: 既存の全機能（データ入出力、DBクリア等）を復元・統合した最終完成版。
 */

// --- 定数定義 ---

// --- シート名 ---
var SHEET_NAME_SETTINGS = "初期設定";
var SHEET_NAME_DATABASE = "データベース";
var SHEET_NAME_WEEKLY_PLAN = "週案";
var SHEET_NAME_INPUT = "週案入力用";
var SHEET_NAME_NEWSLETTER = "学級通信";
var SHEET_NAME_UNIT_MASTER = "単元マスタ";
var SHEET_NAME_LOG = "ログ";

// --- スクリプトプロパティ & トリガー管理用 ---
// 指導計画PDF用
var SCRIPT_PROP_PDF_QUEUE = 'pdfProcessingQueue';
var SCRIPT_PROP_PDF_TOTAL = 'pdfTotalCount';
var TRIGGER_FUNCTION_NAME = 'createUnitMasterFromPdfs';

// 行事予定PDF用
var SCRIPT_PROP_EVENT_PDF_QUEUE = 'eventPdfProcessingQueue';
var SCRIPT_PROP_EVENT_PDF_TOTAL = 'eventPdfTotalCount';
var SCRIPT_PROP_EVENT_PDF_YEAR = 'eventPdfProcessingYear';
var TRIGGER_FUNCTION_NAME_EVENT = 'processNextEventPdf'; 

// --- 固定時間割転記関連 定数 ---
var SETTINGS_RANGE_TIMETABLE = "B22:I26";

// --- 「初期設定」シートのセル定義 ---
var SETTINGS_CELL_COURSE_NAME = "B7";
var SETTINGS_CELL_EVENT_PDF_FOLDER_ID = "B8";
var SETTINGS_CELL_PDF_FOLDER_ID = "B9";
var SETTINGS_CELL_GEMINI_API_KEY = "B10";
var SETTINGS_RANGE_COURSE_LIST_OUTPUT = "B29";

// --- データベースシート列定義 (1始まり) ---
var DB_COL_WEEK_NUM = 1; var DB_COL_DATE = 2; var DB_COL_DAY_OF_WEEK = 3; var DB_COL_TIME = 4; var DB_COL_EVENT = 5; var DB_COL_MORNING = 6; var DB_COL_PERIOD1 = 7; var DB_COL_UNIT1 = 8; var DB_COL_CONTENT1 = 9; var DB_COL_PERIOD2 = 10; var DB_COL_UNIT2 = 11; var DB_COL_CONTENT2 = 12; var DB_COL_PERIOD3 = 13; var DB_COL_UNIT3 = 14; var DB_COL_CONTENT3 = 15; var DB_COL_PERIOD4 = 16; var DB_COL_UNIT4 = 17; var DB_COL_CONTENT4 = 18; var DB_COL_PERIOD5 = 19; var DB_COL_UNIT5 = 20; var DB_COL_CONTENT5 = 21; var DB_COL_PERIOD6 = 22; var DB_COL_UNIT6 = 23; var DB_COL_CONTENT6 = 24; var DB_COL_HOMEWORK = 25; var DB_COL_ITEMS = 26; var DB_COL_RECESS1 = 27; var DB_COL_RECESS2 = 28; var DB_COL_AFTERSCHOOL = 29;

// --- 週案入力用シート関連 定数 ---
var INPUT_ROW_DATE = 2; var INPUT_COL_MONDAY = 2; var INPUT_COL_SUNDAY = 8; var INPUT_ROW_DATA_START = 4; var INPUT_ROW_DATA_END = 29;

// --- 週案入力用シート行番号 -> データベースシート列番号 対応表 ---
var MAPPING = { 2: DB_COL_DATE, 3: DB_COL_DAY_OF_WEEK, 4: DB_COL_TIME, 5: DB_COL_EVENT, 6: DB_COL_MORNING, 7: DB_COL_PERIOD1, 8: DB_COL_UNIT1, 9: DB_COL_CONTENT1, 10: DB_COL_PERIOD2, 11: DB_COL_UNIT2, 12: DB_COL_CONTENT2, 13: DB_COL_RECESS1, 14: DB_COL_PERIOD3, 15: DB_COL_UNIT3, 16: DB_COL_CONTENT3, 17: DB_COL_PERIOD4, 18: DB_COL_UNIT4, 19: DB_COL_CONTENT4, 20: DB_COL_RECESS2, 21: DB_COL_PERIOD5, 22: DB_COL_UNIT5, 23: DB_COL_CONTENT5, 24: DB_COL_PERIOD6, 25: DB_COL_UNIT6, 26: DB_COL_CONTENT6, 27: DB_COL_AFTERSCHOOL, 28: DB_COL_HOMEWORK, 29: DB_COL_ITEMS };

// --- 単元マスタシート列定義 (1始まり) ---
var MASTER_COL_SUBJECT = 1; var MASTER_COL_UNIT_NAME = 2; var MASTER_COL_TOTAL_HOURS = 3; var MASTER_COL_HOUR_NUM = 4; var MASTER_COL_ACTIVITY = 5;

// --- スクリプトプロパティ & トリガー管理用 ---
var SCRIPT_PROP_PDF_QUEUE = 'pdfProcessingQueue';
var SCRIPT_PROP_PDF_TOTAL = 'pdfTotalCount';
var TRIGGER_FUNCTION_NAME = 'createUnitMasterFromPdfs';

// ============================================================
// ===== 基本機能・メニュー関連 =====
// ============================================================

/** スプレッドシートを開いた時にカスタムメニューを追加します。*/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('週案ツール');
  
  menu.addItem('データベース：今日の行を表示', 'TodaysRow')
    .addItem('週案：今週の週案を表示', 'updateValueToWeeklyPlan')
    .addItem('固定時間割を一括転記', 'showBulkTransferSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('クラスルーム連携')
      .addItem('明日の予定を投稿', 'postScheduleToClassroom')
      .addItem('学級通信を投稿', 'autoPostToClassroom'))
    .addSeparator()
    .addSubMenu(ui.createMenu('初期設定・その他')
      .addItem('指導計画PDFの読み込み', 'createUnitMasterFromPdfs_UI')
      .addItem('行事予定PDFをフォルダから読込', 'importEventsFromFolder_UI')
      .addSeparator()
      .addItem('（PDF読込処理を強制停止）', 'resetAllPdfProcessing_UI'));

  menu.addToUi();
}

/** データベースシートの今日の行を選択します。*/
function TodaysRow() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_DATABASE);
    if (!sheet) throw new Error(`シート「${SHEET_NAME_DATABASE}」が見つかりません。`);
    const today = new Date();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const cellValue = data[i][DB_COL_DATE - 1];
      if (cellValue instanceof Date && isSameDate(new Date(cellValue), today)) {
        sheet.setActiveRange(sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()));
        return;
      }
    }
    SpreadsheetApp.getUi().alert("データベースシートに今日の日付が見つかりませんでした。");
  } catch (e) {
    Logger.log(`TodaysRow エラー: ${e.message}`);
    SpreadsheetApp.getUi().alert(`エラー: ${e.message}`);
  }
}

/** 今日の週番号を週案F1に転記 */
function updateValueToWeeklyPlan() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const databaseSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
    const weeklyPlanSheet = ss.getSheetByName(SHEET_NAME_WEEKLY_PLAN);
    if (!databaseSheet || !weeklyPlanSheet) throw new Error("必要なシートが見つかりません");
    const today = new Date(); today.setHours(0, 0, 0, 0); const todayTime = today.getTime();
    const dbData = databaseSheet.getDataRange().getValues(); const dbValues = dbData.slice(1);
    let foundRowData = null;
    for (const row of dbValues) {
      const cellValue = row[DB_COL_DATE - 1];
      if (cellValue instanceof Date) {
        const cellDate = new Date(cellValue); cellDate.setHours(0, 0, 0, 0);
        if (cellDate.getTime() === todayTime) { foundRowData = row; break; }
      }
    }
    if (foundRowData) {
      const weekNumber = foundRowData[DB_COL_WEEK_NUM - 1];
      weeklyPlanSheet.getRange("F1").setValue(weekNumber); Logger.log(`週案F1に週番号 ${weekNumber} を転記`);
    } else Logger.log("DBに今日の日付見つからず、週案F1更新せず");
  } catch (e) { Logger.log(`updateValueToWeeklyPlan エラー: ${e.message}`); SpreadsheetApp.getUi().alert(`エラー: ${e.message}`); }
}

/** 週案F1に+1 */
function addOneToWeeklyPlan() {
  try {
    const weeklyPlanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_WEEKLY_PLAN);
    if (!weeklyPlanSheet) throw new Error(`シート「${SHEET_NAME_WEEKLY_PLAN}」が見つかりません`);
    const cell = weeklyPlanSheet.getRange("F1"); const currentValue = cell.getValue();
    if (typeof currentValue === 'number') cell.setValue(currentValue + 1);
    else SpreadsheetApp.getUi().alert("F1セルが数値ではありません");
  } catch (e) { Logger.log(`addOne エラー: ${e.message}`); SpreadsheetApp.getUi().alert(`エラー: ${e.message}`); }
}

/** 週案F1から-1 */
function subtractOneFromWeeklyPlan() {
  try {
    const weeklyPlanSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_WEEKLY_PLAN);
    if (!weeklyPlanSheet) throw new Error(`シート「${SHEET_NAME_WEEKLY_PLAN}」が見つかりません`);
    const cell = weeklyPlanSheet.getRange("F1"); const currentValue = cell.getValue();
    if (typeof currentValue === 'number') cell.setValue(currentValue - 1);
    else SpreadsheetApp.getUi().alert("F1セルが数値ではありません");
  } catch (e) { Logger.log(`subtractOne エラー: ${e.message}`); SpreadsheetApp.getUi().alert(`エラー: ${e.message}`); }
}

// ============================================================
// ===== 週案入力支援 関連関数 (読込/保存) =====
// ============================================================

/** データベースから週案入力用へデータを読み込みます。*/
function loadDataFromDatabase() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inputSheet = ss.getSheetByName(SHEET_NAME_INPUT);
    const dbSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
    if (!inputSheet || !dbSheet) throw new Error(`必要なシートが見つかりません`);

    const mondayCell = inputSheet.getRange(INPUT_ROW_DATE, INPUT_COL_MONDAY);
    const mondayDateValue = mondayCell.getValue();
    if (!(mondayDateValue instanceof Date)) throw new Error(`「${SHEET_NAME_INPUT}」${mondayCell.getA1Notation()}に有効な日付を入力してください。`);

    const mondayDate = new Date(mondayDateValue);
    mondayDate.setHours(0, 0, 0, 0);

    const dbData = dbSheet.getDataRange().getValues();
    const numRows = INPUT_ROW_DATA_END - INPUT_ROW_DATE + 1;
    const numCols = INPUT_COL_SUNDAY - INPUT_COL_MONDAY + 1;
    const outputData = Array.from({ length: numRows }, () => Array(numCols).fill(""));
    let foundCount = 0;

    for (let i = 0; i < 7; i++) {
      const targetDate = new Date(mondayDate);
      targetDate.setDate(mondayDate.getDate() + i);
      const foundRowData = dbData.find(row => row[DB_COL_DATE - 1] instanceof Date && isSameDate(row[DB_COL_DATE - 1], targetDate));

      if (foundRowData) {
        foundCount++;
        for (const inputRow in MAPPING) {
          const dbColIndex = MAPPING[inputRow] - 1;
          outputData[parseInt(inputRow) - INPUT_ROW_DATE][i] = foundRowData[dbColIndex];
        }
      }
    }

    inputSheet.getRange(INPUT_ROW_DATE, INPUT_COL_MONDAY, numRows, numCols).setValues(outputData);

    if (foundCount === 0) ui.alert("データベースに該当週のデータが見つかりませんでした。");
    else if (foundCount < 7) ui.alert(`${foundCount}日分のデータを読み込みました (一部の日付が見つかりませんでした)。`);
    else ui.alert("1週間分のデータの読み込みが完了しました。");

  } catch (e) {
    logError("loadDataFromDatabase", e);
    ui.alert(`読み込みエラー: ${e.message}`);
  }
}

/** 週案入力用からデータベースへデータを保存します。*/
function saveDataToDatabase() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inputSheet = ss.getSheetByName(SHEET_NAME_INPUT);
    const dbSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
    if (!inputSheet || !dbSheet) throw new Error(`必要なシートが見つかりません`);

    const dates = inputSheet.getRange(INPUT_ROW_DATE, INPUT_COL_MONDAY, 1, 7).getValues()[0];
    const inputValues = inputSheet.getRange(INPUT_ROW_DATA_START, INPUT_COL_MONDAY, INPUT_ROW_DATA_END - INPUT_ROW_DATA_START + 1, 7).getValues();
    const dbData = dbSheet.getDataRange().getValues();
    
    const dbDateRowIndexMap = new Map(dbData.map((row, index) => {
      if (row[DB_COL_DATE - 1] instanceof Date) {
        return [formatDate(row[DB_COL_DATE - 1]), index + 1];
      }
      return [null, null];
    }).filter(item => item[0]));

    let updatedCount = 0;
    let notFoundDates = [];

    for (let i = 0; i < 7; i++) {
      const targetDate = dates[i];
      if (!(targetDate instanceof Date)) continue;

      const targetDateStr = formatDate(targetDate);
      if (dbDateRowIndexMap.has(targetDateStr)) {
        const dbSheetRowIndex = dbDateRowIndexMap.get(targetDateStr);
        for (let inputRow = INPUT_ROW_DATA_START; inputRow <= INPUT_ROW_DATA_END; inputRow++) {
          if (MAPPING[inputRow]) {
            const targetCol = MAPPING[inputRow];
            // D列以降の場合のみ書き込みを実行
            if (targetCol > 3) {
              const value = inputValues[inputRow - INPUT_ROW_DATA_START][i];
              dbSheet.getRange(dbSheetRowIndex, targetCol).setValue(value);
            }
          }
        }
        updatedCount++;
      } else {
        notFoundDates.push(targetDateStr);
      }
    }

    let message = "";
    if (updatedCount > 0) message += `${updatedCount}日分のデータを保存しました。`;
    if (notFoundDates.length > 0) message += `\n以下の日付はDBに見つからず、保存されませんでした: ${notFoundDates.join(', ')}`;
    if (message === "") message = "保存対象のデータがありませんでした。";

    ui.alert(message);
  } catch (e) {
    logError("saveDataToDatabase", e);
    ui.alert(`保存エラー: ${e.message}`);
  }
}

// ============================================================
// ===== 固定時間割転記 関連関数 =====
// ============================================================

/** 指定週の月～金に固定時間割をデータベースに転記します（上書き）。*/
function transferWeeklyTimetable(targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSettings = ss.getSheetByName(SHEET_NAME_SETTINGS);
  const shData = ss.getSheetByName(SHEET_NAME_DATABASE);
  if (!shSettings || !shData) throw new Error("必要なシートが見つかりません");

  const timetableData = shSettings.getRange(SETTINGS_RANGE_TIMETABLE).getValues();
  const firstDayOfWeek = getMondayOfWeek(targetDate);
  const targetRowIndex = findRowIndexByDate(shData, firstDayOfWeek);

  if (targetRowIndex === -1) {
    Logger.log(`転記週 月曜(${formatDate(firstDayOfWeek)}) DBに見つからず`);
    return;
  }

  const numRowsToProcess = 5;
  // 読み書きの範囲をD列からに限定
  const startCol = DB_COL_TIME; // D列
  const numCols = shData.getLastColumn() - startCol + 1;
  const targetRange = shData.getRange(targetRowIndex, startCol, numRowsToProcess, numCols);
  const targetValues = targetRange.getValues();
  let dataUpdated = false;

  for (let i = 0; i < timetableData.length && i < numRowsToProcess; i++) {
    const dayTimetable = timetableData[i];
    const targetRow = targetValues[i]; // この配列はD列から始まります

    // 配列のインデックスをD列基準に調整
    const time_idx = DB_COL_TIME - startCol;          // 4-4=0
    const morning_idx = DB_COL_MORNING - startCol;    // 6-4=2
    const p1_idx = DB_COL_PERIOD1 - startCol;         // 7-4=3
    const p2_idx = DB_COL_PERIOD2 - startCol;         // 10-4=6
    const p3_idx = DB_COL_PERIOD3 - startCol;         // 13-4=9
    const p4_idx = DB_COL_PERIOD4 - startCol;         // 16-4=12
    const p5_idx = DB_COL_PERIOD5 - startCol;         // 19-4=15
    const p6_idx = DB_COL_PERIOD6 - startCol;         // 22-4=18

    if (targetRow[time_idx] !== dayTimetable[0]) { targetRow[time_idx] = dayTimetable[0]; dataUpdated = true; }
    if (targetRow[morning_idx] !== dayTimetable[1]) { targetRow[morning_idx] = dayTimetable[1]; dataUpdated = true; }
    if (targetRow[p1_idx] !== dayTimetable[2]) { targetRow[p1_idx] = dayTimetable[2]; dataUpdated = true; }
    if (targetRow[p2_idx] !== dayTimetable[3]) { targetRow[p2_idx] = dayTimetable[3]; dataUpdated = true; }
    if (targetRow[p3_idx] !== dayTimetable[4]) { targetRow[p3_idx] = dayTimetable[4]; dataUpdated = true; }
    if (targetRow[p4_idx] !== dayTimetable[5]) { targetRow[p4_idx] = dayTimetable[5]; dataUpdated = true; }
    if (targetRow[p5_idx] !== dayTimetable[6]) { targetRow[p5_idx] = dayTimetable[6]; dataUpdated = true; }
    if (targetRow[p6_idx] !== dayTimetable[7]) { targetRow[p6_idx] = dayTimetable[7]; dataUpdated = true; }
  }

  if (dataUpdated) {
    targetRange.setValues(targetValues);
    Logger.log(`${formatDate(firstDayOfWeek)} 週 固定時間割転記(上書き)完了`);
  } else {
    Logger.log(`${formatDate(firstDayOfWeek)} 週 更新不要`);
  }
}

/** 設定画面（サイドバー）を表示する。 */
function showBulkTransferSidebar() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('長期休業期間設定Sidebar').setTitle('長期休業設定と一括転記');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (e) {
    logError("showBulkTransferSidebar", e);
    SpreadsheetApp.getUi().alert(`サイドバーの表示中にエラーが発生しました: ${e.message}`);
  }
}

/** 長期休業期間を除外して、年間の固定時間割をデータベースに一括転記します。 */
function processBulkTransferWithExclusion(dates) {
  try {
    const exclusionPeriodsInput = [
      { name: "夏休み", startStr: dates.summerStart, endStr: dates.summerEnd },
      { name: "冬休み", startStr: dates.winterStart, endStr: dates.winterEnd },
      { name: "春休み", startStr: dates.springStart, endStr: dates.springEnd }
    ];
    const validExclusionPeriods = exclusionPeriodsInput
      .filter(p => p.startStr && p.endStr)
      .map(p => {
        const start = new Date(p.startStr.replace(/-/g, '/'));
        const end = new Date(p.endStr.replace(/-/g, '/'));
        if (start) start.setHours(0,0,0,0);
        if (end) end.setHours(0,0,0,0);
        return { name: p.name, start: start, end: end };
      }).filter(p =>
        p.start instanceof Date && !isNaN(p.start.getTime()) &&
        p.end instanceof Date && !isNaN(p.end.getTime()) &&
        p.start.getTime() <= p.end.getTime()
      );

    validExclusionPeriods.forEach(p => Logger.log(`有効な除外期間: ${p.name} ${formatDate(p.start)} ～ ${formatDate(p.end)}`));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shData = ss.getSheetByName(SHEET_NAME_DATABASE);
    if (!shData) throw new Error(`シート「${SHEET_NAME_DATABASE}」が見つかりません`);
    
    // ★★★ 修正: 日付が入っている最終行を正しく取得 ★★★
    const dateColumnValues = shData.getRange(1, DB_COL_DATE, shData.getLastRow(), 1).getValues();
    let lastRowWithDate = 0;
    for (let i = dateColumnValues.length - 1; i >= 0; i--) {
      if (dateColumnValues[i][0] instanceof Date) {
        lastRowWithDate = i + 1;
        break;
      }
    }
    if (lastRowWithDate < 2) return "DBに有効な日付データがありません";
    
    const lastDbDate = new Date(shData.getRange(lastRowWithDate, DB_COL_DATE).getValue());
    lastDbDate.setHours(0,0,0,0);

    const today = new Date();
    const dayOfWeek = today.getDay();
    const daysUntilNextMonday = (dayOfWeek === 0) ? 1 : (8 - dayOfWeek);
    let currentDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + daysUntilNextMonday);
    currentDate.setHours(0,0,0,0);

    Logger.log(`一括転記(除外あり) 開始: ${formatDate(currentDate)} ～ ${formatDate(lastDbDate)}`);
    let skippedDayCount = 0;

    while (currentDate <= lastDbDate) {
      const currentDayOfWeek = currentDate.getDay();
      if (currentDayOfWeek >= 1 && currentDayOfWeek <= 5) {
        const isExcluded = validExclusionPeriods.some(p => isDateInRange(currentDate, p.start, p.end));
        if (!isExcluded && currentDayOfWeek === 1) {
          try {
            transferWeeklyTimetable(currentDate);
          } catch (e) {
            logError(`一括転記中のエラー (${formatDate(currentDate)})`, e);
          }
        } else if (isExcluded) {
          skippedDayCount++;
        }
      }
      currentDate.setDate(currentDate.getDate() + 1);
    }

    Logger.log("一括転記(除外あり) 完了");
    const skipMessage = skippedDayCount > 0 ? ` (${skippedDayCount}日分スキップ)` : "";
    return `一括転記が完了しました${skipMessage}`;
  } catch (e) {
    logError("processBulkTransferWithExclusion", e);
    throw new Error(`一括転記処理中にエラーが発生しました: ${e.message}`);
  }
}

/** 長期休業期間のデフォルト日付を取得します (HTML側から呼び出される)。*/
function getDefaultExclusionDates() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
    const databaseSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
    if (!settingSheet || !databaseSheet) throw new Error("必要なシートが見つかりません");
    
    const yearValueB6 = settingSheet.getRange("B6").getValue();
    let summerYear;
    if (yearValueB6 instanceof Date) {
      summerYear = yearValueB6.getFullYear();
    } else if (typeof yearValueB6 === 'number' && yearValueB6 > 1900 && yearValueB6 < 2200) {
      summerYear = yearValueB6;
    } else {
      summerYear = new Date().getFullYear();
      logInfo("初期設定B6の年度が無効なため、現在の年を使用します。");
    }
    
    const currentYear = new Date().getFullYear();
    const summerStart = new Date(summerYear, 6, 21); // 7月21日
    const summerEnd = new Date(summerYear, 7, 31);   // 8月31日
    const winterStart = new Date(currentYear, 11, 26); // 12月26日
    const winterEnd = new Date(currentYear + 1, 0, 7);   // 翌年1月7日
    const springStart = new Date(currentYear + 1, 2, 26); // 翌年3月26日
    
    let springEnd = new Date(springStart);
    const dateColumnValues = databaseSheet.getRange(1, DB_COL_DATE, databaseSheet.getLastRow(), 1).getValues();
    let lastRowWithDate = 0;
    for (let i = dateColumnValues.length - 1; i >= 0; i--) {
      if (dateColumnValues[i][0] instanceof Date) {
        lastRowWithDate = i + 1;
        break;
      }
    }

    if (lastRowWithDate >= 2) {
      const lastDateValue = databaseSheet.getRange(lastRowWithDate, DB_COL_DATE).getValue();
      if (lastDateValue instanceof Date) {
        springEnd = new Date(lastDateValue);
      }
    } else {
      logInfo("データベースに有効な日付データがないため、春休みの終了日を仮設定します。");
    }

    // HTMLの <input type="date"> 用に YYYY-MM-DD 形式で返すためのヘルパー関数
    const formatDateForInput = (date) => {
        if (!(date instanceof Date) || isNaN(date.getTime())) return "";
        return Utilities.formatDate(date, "JST", "yyyy-MM-dd");
    };

    return {
      summerStart: formatDateForInput(summerStart), summerEnd: formatDateForInput(summerEnd),
      winterStart: formatDateForInput(winterStart), winterEnd: formatDateForInput(winterEnd),
      springStart: formatDateForInput(springStart), springEnd: formatDateForInput(springEnd)
    };
  } catch (e) {
     logError("getDefaultExclusionDates", e);
     return { summerStart: '', summerEnd: '', winterStart: '', winterEnd: '', springStart: '', springEnd: '' };
  }
}

/** 固定時間割を「初期設定」シートから「週案入力用」シートへ転記(時程と朝学習は転記対象外)します。*/
function transferFixedTimetableToInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
  const inputSheet = ss.getSheetByName(SHEET_NAME_INPUT);

  if (!settingsSheet || !inputSheet) {
    throw new Error("必要なシート（初期設定または週案入力用）が見つかりません。");
  }

  // 固定時間割データを読み込む
  const timetableData = settingsSheet.getRange(SETTINGS_RANGE_TIMETABLE).getValues();

  // 転記先の行番号を定義 (1〜6校時のみ)
  const destRows = {
    p1: 7,
    p2: 10,
    p3: 14,
    p4: 17,
    p5: 21,
    p6: 24
  };

  // 月曜日から金曜日までループ (5日間)
  for (let dayIndex = 0; dayIndex < 5; dayIndex++) {
    const destCol = INPUT_COL_MONDAY + dayIndex; // B列からF列
    const dayData = timetableData[dayIndex];   // 月曜日のデータ、火曜日のデータ...

    // 各項目を転記 (1〜6校時のみ)
    inputSheet.getRange(destRows.p1, destCol).setValue(dayData[2]); // 1校時
    inputSheet.getRange(destRows.p2, destCol).setValue(dayData[3]); // 2校時
    inputSheet.getRange(destRows.p3, destCol).setValue(dayData[4]); // 3校時
    inputSheet.getRange(destRows.p4, destCol).setValue(dayData[5]); // 4校時
    inputSheet.getRange(destRows.p5, destCol).setValue(dayData[6]); // 5校時
    inputSheet.getRange(destRows.p6, destCol).setValue(dayData[7]); // 6校時
  }
  logInfo("「週案入力用」シートへ固定時間割（1〜6校時）を転記しました。");
}

// ============================================================
// ===== 行事予定PDF転記 関連関数 =====
// ============================================================

/** 指定されたフォルダ内の行事予定PDFをすべて読み込み、データベースに転記します。*/
function importEventsFromFolder_UI() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
    if (!settingsSheet) throw new Error(`シート「${SHEET_NAME_SETTINGS}」が見つかりません。`);

    const folderId = settingsSheet.getRange(SETTINGS_CELL_EVENT_PDF_FOLDER_ID).getValue();
    if (!folderId) throw new Error(`「${SHEET_NAME_SETTINGS}」シートのセル「${SETTINGS_CELL_EVENT_PDF_FOLDER_ID}」に行事予定PDFのフォルダIDを入力してください。`);

    const yearResponse = ui.prompt('年度の入力', '処理対象の年度（4月始まり）を西暦で入力してください。\n例: 2025', ui.ButtonSet.OK_CANCEL);
    if (yearResponse.getSelectedButton() !== ui.Button.OK || !yearResponse.getResponseText()) {
      ui.alert('処理をキャンセルしました。');
      return;
    }
    const fiscalYear = parseInt(yearResponse.getResponseText().trim(), 10);

    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const fileIds = [];
    while (files.hasNext()) {
      fileIds.push(files.next().getId());
    }
    if (fileIds.length === 0) {
      throw new Error(`指定されたフォルダ「${folder.getName()}」にPDFファイルが見つかりませんでした。`);
    }

    const today = new Date();
    const firstDayOfCurrentMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    
    const allMonths = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]; // 年度内の全ての月
    const monthsToProcess = allMonths.filter(month => {
      const yearForMonth = (month >= 4) ? fiscalYear : fiscalYear + 1;
      const firstDayOfMonth = new Date(yearForMonth, month - 1, 1);
      // その月の1日が、今月の1日以降であれば処理対象とする
      return firstDayOfMonth >= firstDayOfCurrentMonth;
    });

    if (monthsToProcess.length === 0) {
      ui.alert('処理対象となる月（今月以降）がありませんでした。');
      return;
    }
    // 最適化された「やることリスト」を作成
    const processingQueue = [];
    fileIds.forEach(fileId => {
      monthsToProcess.forEach(month => {
        processingQueue.push({ fileId: fileId, month: month });
      });
    });

    const confirmResponse = ui.alert('処理の開始', 
      `${fileIds.length} 個のPDFファイルから、**${monthsToProcess.join('月, ')}月**の予定を読み込みます。（合計 ${processingQueue.length} タスク）\n` +
      `処理はバックグラウンドで自動的に中断・再開されます。\n\n` +
      `実行しますか？`, 
      ui.ButtonSet.YES_NO);
    if (confirmResponse !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    resetEventPdfProcessing(); 

    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(SCRIPT_PROP_EVENT_PDF_QUEUE, JSON.stringify(processingQueue));
    properties.setProperty(SCRIPT_PROP_EVENT_PDF_TOTAL, processingQueue.length);
    properties.setProperty(SCRIPT_PROP_EVENT_PDF_YEAR, fiscalYear.toString());

    ss.toast(`行事予定PDFの読み込みを開始しました。(0/${processingQueue.length})`, '処理開始', -1);
    
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME_EVENT).timeBased().after(1000).create();

  } catch (e) {
    logError("importEventsFromFolder_UI", e);
    ui.alert(`エラーが発生しました。\n\n詳細: ${e.message}\n\n「ログ」シートもご確認ください。`);
  }
}

/** トリガーによって呼び出される関数。ファイルを1つずつ処理し、次のトリガーをセットします。*/
function processNextEventPdf() {
  const startTime = new Date();
  const properties = PropertiesService.getScriptProperties();
  
  const queueJson = properties.getProperty(SCRIPT_PROP_EVENT_PDF_QUEUE);
  const year = properties.getProperty(SCRIPT_PROP_EVENT_PDF_YEAR);

  if (!queueJson || !year) {
    SpreadsheetApp.getActiveSpreadsheet().toast("行事予定PDFの読み込みがすべて完了しました。", "処理完了", 10);
    logInfo("すべての行事予定PDFの処理が完了しました。");
    resetEventPdfProcessing();
    return;
  }

  const queue = JSON.parse(queueJson);
  if (queue.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("行事予定PDFの読み込みがすべて完了しました。", "処理完了", 10);
    logInfo("キューが空です。すべての行事予定PDFの処理が完了しました。");
    resetEventPdfProcessing();
    return;
  }
  
  // やることリストの一番上から、仕事（ファイルIDと月のセット）を1つ取り出す
  const task = queue.shift(); 
  const file = DriveApp.getFileById(task.fileId);
  
  const totalTasks = parseInt(properties.getProperty(SCRIPT_PROP_EVENT_PDF_TOTAL), 10);
  const processedCount = totalTasks - queue.length;
  SpreadsheetApp.getActiveSpreadsheet().toast(`行事予定PDF 処理中... (${processedCount}/${totalTasks})\nファイル名: ${file.getName()} (${task.month}月)`, `処理中`, -1);

  try {
    // 実際にAIで分析する関数に、月の情報も渡す
    processEventPdf(task.fileId, year, task.month);
  } catch (e) {
    logError(`行事予定PDFの処理中にエラーが発生しました: ${file.getName()} (${task.month}月)`, e);
  }

  properties.setProperty(SCRIPT_PROP_EVENT_PDF_QUEUE, JSON.stringify(queue));

  const executionTime = (new Date() - startTime) / 1000;
  deleteTriggers_(TRIGGER_FUNCTION_NAME_EVENT);

  if (queue.length > 0) {
    if (executionTime < 300) { 
      ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME_EVENT).timeBased().after(1000).create();
    } else {
      logInfo(`時間切れのため行事予定PDFの処理を中断・再開します。残り: ${queue.length} タスク`);
      ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME_EVENT).timeBased().everyMinutes(5).create();
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("行事予定PDFの読み込みがすべて完了しました。", "処理完了", 10);
    logInfo("すべての行事予定PDFの処理が完了しました。");
    resetEventPdfProcessing();
  }
}

/** AIが抽出した予定をデータベースに書き込む関数。過去の日付は上書きせず、さらに「既に同じ予定が入力済み」の場合も追記しない。
 * @param {string} fileId 選択されたPDFのファイルID
 * @param {string} year PDFが対象とする年度 (4月始まり)
 * @param {number} month 処理対象の月 (1-12)
 * @returns {string} 処理結果のメッセージ
 */
function processEventPdf(fileId, year, month) {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const apiKey = getApiKey_();
    const schoolYear = parseInt(year, 10);

    // (AIへの指示プロンプトは変更ありません)
    const prompt = `
あなたは日本の学校の行事予定を整理する専門家です。
添付されたPDFファイルは、ある学校の ${schoolYear} 年度（ ${schoolYear} 年4月～ ${schoolYear + 1} 年3月）の行事予定表です。
このPDFの中から、「${month}月」に関する予定だけを抽出し、以下のルールに従ってJSON形式の配列で出力してください。
# 抽出ルール
1.  日付の特定:
    - ${month}月の日付と予定のみを抽出してください。他の月の情報は無視してください。
    - ${schoolYear} 年度なので、4月～12月は ${schoolYear} 年、1月～3月は ${schoolYear + 1} 年として日付を生成してください。
    - 最終的な日付は必ず "YYYY-MM-DD" 形式にしてください。
2.  内容の分類:
    - 児童生徒が関わる学校行事（例：始業式, 遠足, 運動会, 委員会, クラブ）は、typeを "event" としてください。
    - 教職員のみが関わる予定（例：会議, 研修, 出張, 初任研, 三部会）は、typeを "meeting" としてください。
3.  複数予定の分割: 1つの日付に複数の予定がある場合は、それぞれ別のオブジェクトとしてください。
# 出力形式 (JSON配列)
[
  { "date": "YYYY-MM-DD", "content": "（${month}月の予定の内容）", "type": "event" },
  { "date": "YYYY-MM-DD", "content": "（${month}月の予定の内容）", "type": "meeting" }
]
`;
    const extractedEvents = callGeminiApi_(prompt, apiKey, [blob]);
    if (!extractedEvents || !Array.isArray(extractedEvents)) {
      logInfo(`PDF「${file.getName()}」の${month}月からは、有効な予定が見つかりませんでした。`);
      return "0 件の予定を転記しました。";
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
    const dbData = dbSheet.getRange(2, DB_COL_DATE, dbSheet.getLastRow() - 1, 1).getValues();
    const dateMap = new Map(dbData.map((row, i) => [formatDate(row[0]), i + 2]));
    
    let updatedCount = 0;
    let pastDateSkippedCount = 0;
    let duplicateSkippedCount = 0; // ★ 重複してスキップした件数を数える変数を追加

    extractedEvents.forEach(item => {
      if (!item.date || !/^\d{4}-\d{2}-\d{2}$/.test(item.date) || !item.content) return;
      
      const targetDate = new Date(item.date.replace(/-/g, '/'));
      targetDate.setHours(0, 0, 0, 0);
      const targetDateStr = formatDate(targetDate);

      if (dateMap.has(targetDateStr)) {
        if (targetDate >= today) {
          const rowNum = dateMap.get(targetDateStr);
          let targetCol;
          if (item.type === 'event') {
            targetCol = DB_COL_EVENT;
          } else if (item.type === 'meeting') {
            targetCol = DB_COL_AFTERSCHOOL;
          }

          if (targetCol) {
            const cell = dbSheet.getRange(rowNum, targetCol);
            const currentValue = cell.getValue().toString();
            const newContent = item.content.toString().trim();

            // --- ★★★ ここが新しいチェック処理 ★★★ ---
            // もし、現在のセルの内容に、新しい予定が「含まれていない」場合だけ、追記処理を行います。
            if (!currentValue.includes(newContent)) {
              const newValue = currentValue ? `${currentValue}\n${newContent}` : newContent;
              cell.setValue(newValue);
              updatedCount++;
            } else {
              // 既に含まれている場合は、重複スキップとしてカウントします。
              duplicateSkippedCount++;
            }
          }
        } else {
          pastDateSkippedCount++; 
        }
      }
    });

    // --- 結果の報告をさらに詳しく ---
    let logMessage = `${updatedCount} 件の予定をPDF「${file.getName()}」(${month}月分)から転記。`;
    if (pastDateSkippedCount > 0) logMessage += ` ${pastDateSkippedCount} 件(過去),`;
    if (duplicateSkippedCount > 0) logMessage += ` ${duplicateSkippedCount} 件(重複)はスキップ。`;
    logInfo(logMessage);
    
    let resultMessage = `${updatedCount} 件の予定を転記しました。`;
    const skippedMessages = [];
    if (pastDateSkippedCount > 0) skippedMessages.push(`${pastDateSkippedCount} 件は過去`);
    if (duplicateSkippedCount > 0) skippedMessages.push(`${duplicateSkippedCount} 件は重複`);
    
    if (skippedMessages.length > 0) {
      resultMessage += ` (${skippedMessages.join(', ')}のためスキップ)`;
    }
    return resultMessage;
    
  } catch (e) {
    logError(`processEventPdf (${month}月分)`, e);
    throw new Error(e.message);
  }
}

// ============================================================
// ===== 指導計画PDF転記 関連関数 =====
// ============================================================
/** ユーザーに実行を確認した後、指導計画PDFから「単元マスタ」を作成する分割処理を開始させます。 */
function createUnitMasterFromPdfs_UI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert( '指導計画PDFの読み込み', '「初期設定」シートで指定されたフォルダ内のPDFをAIが読み取り、「単元マスタ」シートを作成・更新します。\n' + '処理はバックグラウンドで自動的に中断・再開され、完了まで数分～数十分かかる場合があります。\n\n' + '実行しますか？', ui.ButtonSet.YES_NO );
  if (response == ui.Button.YES) {
    resetPdfProcessing();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
    if (!settingsSheet) throw new Error(`シート「${SHEET_NAME_SETTINGS}」が見つかりません。`);
    const folderId = settingsSheet.getRange(SETTINGS_CELL_PDF_FOLDER_ID).getValue();
    if (!folderId) {
      ui.alert(`「${SHEET_NAME_SETTINGS}」シートのセル「${SETTINGS_CELL_PDF_FOLDER_ID}」にフォルダIDを入力してください。`);
      return;
    }
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const fileIds = [];
    while (files.hasNext()) { fileIds.push(files.next().getId()); }
    if (fileIds.length === 0) {
      ui.alert("指定されたフォルダにPDFファイルが見つかりませんでした。");
      return;
    }
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(SCRIPT_PROP_PDF_QUEUE, JSON.stringify(fileIds));
    properties.setProperty(SCRIPT_PROP_PDF_TOTAL, fileIds.length);
    let masterSheet = ss.getSheetByName(SHEET_NAME_UNIT_MASTER);
    if (masterSheet) { masterSheet.clear(); }
    else { masterSheet = ss.insertSheet(SHEET_NAME_UNIT_MASTER); }
    masterSheet.getRange("A1:E1").setValues([["教科", "単元名", "総時間数", "何時間目", "時間ごとの学習活動"]]).setFontWeight("bold");
    ss.toast(`PDF読み込み処理を開始しました。(0/${fileIds.length})`, '処理開始', -1);
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME).timeBased().after(1000).create();
  }
}

/** タイムアウトを避けながら、指導計画PDFを一つずつAIで分析し「単元マスタ」に書き込みます。 */
function createUnitMasterFromPdfs() {
  const startTime = new Date();
  const properties = PropertiesService.getScriptProperties();
  const queueJson = properties.getProperty(SCRIPT_PROP_PDF_QUEUE);
  if (!queueJson) {
    logInfo("すべてのPDF処理が完了しました。");
    SpreadsheetApp.getActiveSpreadsheet().toast("PDFの読み込みがすべて完了しました。", "処理完了", 10);
    resetPdfProcessing();
    return;
  }
  const fileIds = JSON.parse(queueJson);
  if (fileIds.length === 0) {
    logInfo("キューが空です。すべてのPDF処理が完了しました。");
    SpreadsheetApp.getActiveSpreadsheet().toast("PDFの読み込みがすべて完了しました。", "処理完了", 10);
    resetPdfProcessing();
    return;
  }
  const totalFiles = parseInt(properties.getProperty(SCRIPT_PROP_PDF_TOTAL), 10);
  const fileId = fileIds.shift();
  const file = DriveApp.getFileById(fileId);
  const processedCount = totalFiles - fileIds.length;
  SpreadsheetApp.getActiveSpreadsheet().toast(`PDF処理中... (${processedCount}/${totalFiles}) \nファイル名: ${file.getName()}`, `処理中`, -1);
  try {
    processSinglePdf(file);
  } catch (e) {
    logError(`PDF処理中に致命的なエラーが発生しました: ${file.getName()}`, e);
  }
  properties.setProperty(SCRIPT_PROP_PDF_QUEUE, JSON.stringify(fileIds));
  const executionTime = (new Date() - startTime) / 1000 / 60;
  deleteTriggers_(TRIGGER_FUNCTION_NAME);
  if (fileIds.length > 0) {
    if (executionTime < 5) {
      ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME).timeBased().after(1000).create();
    } else {
      logInfo(`時間切れのため処理を中断・再開します。残り: ${fileIds.length}件`);
      ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME).timeBased().everyMinutes(5).create();
    }
  } else {
    logInfo("すべてのPDF処理が完了しました。");
    SpreadsheetApp.getActiveSpreadsheet().toast("PDFの読み込みがすべて完了しました。", "処理完了", 10);
    resetPdfProcessing();
  }
}

/** 1つのPDFファイルを処理してシートに書き込みます。（抽出後、各単元の最後に「まとめ」の時間を自動で追加） @param {GoogleAppsScript.Drive.File} file 処理対象のPDFファイル*/
function processSinglePdf(file) {
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_UNIT_MASTER);
  const apiKey = getApiKey_();
  const prompt = `あなたは日本の小学校教育の専門家です。添付された年間指導計画のPDFから、以下の情報を抽出し、指定されたJSON形式で出力してください。
このPDFは複数ページにわたる長いものである可能性があります。すべてのページを注意深く読み取り、すべての単元を抽出してください。

抽出項目:
1.  **subject**: 教科名（例：国語, 算数）
2.  **unitName**: 単元名や題材名
3.  **totalHours**: その単元に配当されている合計時間数（半角数字）
4.  **hourlyActivities**: その単元の、時間ごとの「主な学習活動」または「学習内容」のリスト。
    - **hour**: 何時間目かを示す半角数字。
    - **activity**: その時間に行う具体的な学習活動の内容。ただし、その時間の「ねらい」「中心的な学習活動」「指導の要点」「評価」などを総合的に判断し、週案に記載するのにふさわしい100文字程度の簡潔な要約にしてください。

出力は、必ず単一の有効なJSON配列としてください。途中で途切れたり、フォーマットが崩れたりしないようにしてください。
出力形式（JSON配列）:
[
  {
    "subject": "教科名",
    "unitName": "単元名1",
    "totalHours": 8,
    "hourlyActivities": [
      { "hour": 1, "activity": "（1時間目の活動の100文字程度の要約）" },
      { "hour": 2, "activity": "（2時間目の活動の100文字程度の要約）" }
    ]
  }
]`;
  logInfo(`PDF処理中: ${file.getName()}`);
  try {
    const extractedUnits = callGeminiApi_(prompt, apiKey, [file.getBlob()]);
    if (extractedUnits && Array.isArray(extractedUnits)) {
      const allRows = [];
      extractedUnits.forEach(unit => {
        // First, add the regular hourly activities for the current unit
        if (unit.hourlyActivities && Array.isArray(unit.hourlyActivities)) {
          unit.hourlyActivities.forEach(activity => {
            allRows.push([
              unit.subject || '',
              unit.unitName || '',
              unit.totalHours || '',
              activity.hour || '',
              activity.activity || ''
            ]);
          });
        }
        // Now, add the summary "unit" immediately after
        if (unit.unitName && unit.totalHours > 0) {
            allRows.push([
                unit.subject || '',
                `${unit.unitName} のまとめ`, // The unit name for the summary
                1, // Total hours for a summary is always 1
                1, // The hour number for a summary is always 1
                "単元の内容を振り返り、学習の定着を確認する。" // A standard text for the summary
            ]);
        }
      });
      if (allRows.length > 0) {
        masterSheet.getRange(masterSheet.getLastRow() + 1, 1, allRows.length, allRows[0].length).setValues(allRows);
      }
    }
  } catch (e) {
    logError(`PDF解析エラー: ${file.getName()}`, e);
  }
}

// ============================================================
// ===== PDF処理リセット関連関数 ==============================
// ============================================================

/** ユーザーがメニューから実行する、すべてのPDF処理を停止するための関数です。*/
function resetAllPdfProcessing_UI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '処理の強制停止',
    '実行中のすべてのPDF読み込み処理（指導計画・行事予定）を停止し、待機状態を解除します。\nよろしいですか？',
    ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    // 指導計画と行事予定、両方の停止処理を呼び出します。
    resetUnitMasterProcessing();
    resetEventPdfProcessing();

    // 処理状況を表示しているセルをクリアします。
    try {
      const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SETTINGS);
      if (settingsSheet) {
        // B11セルをクリア
        settingsSheet.getRange(SETTINGS_CELL_STATUS).clearContent();
      }
    } catch (e) {
      logError("ステータスセルのクリアに失敗", e);
    }

    ui.alert('すべてのPDF読み込み処理を停止しました。');
  }
}

/** 指導計画PDFの処理キューとトリガーをリセット（強制停止）します。*/
function resetUnitMasterProcessing() {
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_PDF_QUEUE);
  PropertiesService.getScriptProperties().deleteProperty(SCRIPT_PROP_PDF_TOTAL);
  deleteTriggers_(TRIGGER_FUNCTION_NAME);
  logInfo("指導計画PDF処理のキューとトリガーをリセットしました。");
}

/** 行事予定PDFの処理キューとトリガーをリセット（強制停止）します。*/
function resetEventPdfProcessing() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(SCRIPT_PROP_EVENT_PDF_QUEUE);
  properties.deleteProperty(SCRIPT_PROP_EVENT_PDF_TOTAL);
  properties.deleteProperty(SCRIPT_PROP_EVENT_PDF_YEAR);
  deleteTriggers_(TRIGGER_FUNCTION_NAME_EVENT);
  logInfo("行事予定PDF処理のキューとトリガーをリセットしました。");
}

/** 指定された名前のトリガー（タイマー）をすべて削除するヘルパー関数です。 @param {string} functionName 削除したいトリガーが呼び出す関数名*/
function deleteTriggers_(functionName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ============================================================
// ===== 週案自動入力機能群 (マスタ参照版) =====
// ============================================================

/** ユーザーに実行を確認した後、「単元マスタ」を基に週案を自動入力する処理を開始させます。 */
function populateWeeklyPlanFromMaster_UI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('単元内容の自動入力', '「週案入力用」シートに表示されている週の単元名と内容を、「単元マスタ」を基に自動で入力します。\n\n実行しますか？', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().toast('単元内容の自動入力を開始しました...', '処理開始', -1);
    try {
      populateWeeklyPlanFromMaster();
      SpreadsheetApp.getActiveSpreadsheet().toast('自動入力が完了しました。', '処理完了', 10);
    } catch (e) {
      logError('populateWeeklyPlanFromMaster', e);
      ui.alert(`処理中にエラーが発生しました。\n詳細は「${SHEET_NAME_LOG}」シートを確認してください。\n\nエラー: ${e.message}`);
    }
  }
}

/** 「単元マスタ」と過去の進捗を基に、教科ごとの次の単元名と学習内容を「週案入力用」シートに自動で入力します。 */
function populateWeeklyPlanFromMaster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName(SHEET_NAME_INPUT);
  const dbSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
  const masterSheet = ss.getSheetByName(SHEET_NAME_UNIT_MASTER);
  if (!inputSheet || !dbSheet || !masterSheet) throw new Error(`必要なシートが見つかりません。`);

  const masterData = masterSheet.getDataRange().getValues();
  const weekDates = inputSheet.getRange(INPUT_ROW_DATE, INPUT_COL_MONDAY, 1, 5).getValues()[0];
  const timetableValues = inputSheet.getRange(7, INPUT_COL_MONDAY, 18, 5).getValues();
  const periodRowIndexes = { 1: 0, 2: 3, 3: 7, 4: 10, 5: 14, 6: 17 };

  let weeklyProgress = {};

  // ★★★ 追加: 週案の開始日を取得 ★★★
  const weekStartDate = weekDates[0];
  if (!(weekStartDate instanceof Date)) {
      SpreadsheetApp.getUi().alert("「週案入力用」シートの月曜日に有効な日付がありません。処理を中断します。");
      return;
  }

  for (let col = 0; col < 5; col++) {
    const currentDate = weekDates[col];
    if (!(currentDate instanceof Date)) continue;
    logInfo(`${formatDate(currentDate)} の処理を開始...`);
    for (let period = 1; period <= 6; period++) {
      const rowIndex = periodRowIndexes[period];
      const subject = timetableValues[rowIndex][col];
      if (!subject || subject.includes("行事")) continue;
      try {
        // ★★★ 修正: findLastLesson_に週の開始日を渡す ★★★
        const lastLesson = weeklyProgress[subject] || findLastLesson_(dbSheet, subject, weekStartDate);
        
        const nextLesson = determineNextLesson_(lastLesson, masterData, subject);
        weeklyProgress[subject] = nextLesson;
        const activity = findActivityFromMaster_(masterData, subject, nextLesson.unitName, nextLesson.currentHour);
        logInfo(`  [${subject}] 今回: ${nextLesson.unitName} (${nextLesson.currentHour}/${nextLesson.totalHours})`);
        const unitCell = inputSheet.getRange(7 + rowIndex + 1, INPUT_COL_MONDAY + col);
        const contentCell = inputSheet.getRange(7 + rowIndex + 2, INPUT_COL_MONDAY + col);
        unitCell.setValue(`${nextLesson.unitName} ${nextLesson.currentHour}/${nextLesson.totalHours}`);
        contentCell.setValue(activity);
      } catch (e) {
        logError(`  [${subject}] ${period}校時の処理中にエラー`, e);
      }
    }
  }
}

/** 「単元マスタ」の中から、指定された教科・単元名・時間に対応する学習活動を探し出します。 */
function findActivityFromMaster_(masterData, subject, unitName, hourNum) {
  for (let i = 1; i < masterData.length; i++) {
    const row = masterData[i];
    if (row[MASTER_COL_SUBJECT - 1] === subject && row[MASTER_COL_UNIT_NAME - 1] === unitName && row[MASTER_COL_HOUR_NUM - 1] == hourNum) {
      return row[MASTER_COL_ACTIVITY - 1];
    }
  }
  if (unitName.includes("のまとめ")) return "単元の内容を振り返り、学習の定着を確認する。";
  return "（単元マスタに該当する活動が見つかりませんでした）";
}

/** データベースを検索し、指定された教科の最後の授業情報を返します。
 * ★★★ 修正: 指定された日付の前日までのデータから最新の授業を検索するように変更 ★★★
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dbSheet - データベースシート。
 * @param {string} subject - 検索する教科名。
 * @param {Date} weekStartDate - 週案入力用シートの開始日（月曜日）。この日付より前のデータを検索対象とします。
 * @returns {{unitName: string, currentHour: number, totalHours: number}} 最後の授業情報。
 * @private
 */
function findLastLesson_(dbSheet, subject, weekStartDate) {
  const data = dbSheet.getDataRange().getValues();
  
  // 検索対象の終了日（週案開始日の前日）を設定
  const searchEndDate = new Date(weekStartDate);
  searchEndDate.setDate(searchEndDate.getDate() - 1);

  // データベースの最終行から2行目に向かって逆順にループ
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowDate = row[DB_COL_DATE - 1];

    // 日付が有効で、かつ検索範囲内であるかを確認
    if (rowDate instanceof Date && rowDate <= searchEndDate) {
      // 6校時から1校時に向かって逆順に教科を検索
      for (let col = DB_COL_PERIOD6 - 1; col >= DB_COL_PERIOD1 - 1; col -= 3) {
        if (row[col] === subject) {
          const unitText = row[col + 1]; // 単元名のセル
          if (unitText && typeof unitText === 'string') {
            // "単元名 1/8" のような形式を正規表現で解析
            const match = unitText.match(/(.+?)\s*(\d+)\/(\d+)/);
            if (match) {
              // 見つかったら即座にその情報を返す
              return {
                unitName: match[1].trim(),
                currentHour: parseInt(match[2], 10),
                totalHours: parseInt(match[3], 10)
              };
            }
          }
        }
      }
    }
  }

  // 検索範囲内に該当する教科が見つからなかった場合
  return { unitName: null, currentHour: 0, totalHours: 0 };
}

/** 前回の授業情報と単元マスタを基に、次の授業情報を決定します。*/
function determineNextLesson_(lastLesson, masterData, subject) {
  // Case 1: The previous lesson's unit is still in progress.
  if (lastLesson.unitName && lastLesson.currentHour < lastLesson.totalHours) {
    return {
      unitName: lastLesson.unitName,
      currentHour: lastLesson.currentHour + 1,
      totalHours: lastLesson.totalHours
    };
  }

  // Case 2: The previous lesson's unit is finished, or there is no history.
  let nextLessonRow;

  if (lastLesson.unitName) {
    // Find the index of the last lesson in the master data.
    const lastLessonIndex = masterData.findIndex(row =>
      row[MASTER_COL_SUBJECT - 1] === subject &&
      row[MASTER_COL_UNIT_NAME - 1] === lastLesson.unitName &&
      row[MASTER_COL_HOUR_NUM - 1] == lastLesson.currentHour
    );

    if (lastLessonIndex > -1 && lastLessonIndex + 1 < masterData.length) {
      const potentialNextRow = masterData[lastLessonIndex + 1];
      // Check if the next row belongs to the same subject
      if (potentialNextRow[MASTER_COL_SUBJECT - 1] === subject) {
        nextLessonRow = potentialNextRow;
      }
    }
  }
  
  // If a next lesson wasn't found (or there was no history), find the first lesson for the subject.
  if (!nextLessonRow) {
    nextLessonRow = masterData.find(row => row[MASTER_COL_SUBJECT - 1] === subject);
  }

  if (!nextLessonRow) {
    throw new Error(`単元マスタに教科「${subject}」のデータが見つかりません。`);
  }

  return {
    unitName: nextLessonRow[MASTER_COL_UNIT_NAME - 1],
    currentHour: parseInt(nextLessonRow[MASTER_COL_HOUR_NUM - 1], 10),
    totalHours: parseInt(nextLessonRow[MASTER_COL_TOTAL_HOURS - 1], 10)
  };
}

// ============================================================
// ===== データベースクリア・設定 =====
// ============================================================

/** データベースシートのデータ範囲をクリアします（確認付き）。*/
function clearDatabaseDataWithConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_NAME_DATABASE);

  // 先にシートの存在を確認
  if (!dbSheet) {
    ui.alert(`エラー: シート「${SHEET_NAME_DATABASE}」が見つかりません。`);
    return;
  }

  const confirmationMessage = `「${SHEET_NAME_DATABASE}」シートの入力内容（D2以降）を全てクリアします。\n元に戻せません。よろしいですか？`;
  const response = ui.alert('データクリア確認', confirmationMessage, ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    try {
      const lastRow = dbSheet.getLastRow();
      if (lastRow < 2) {
        Browser.msgBox(`「${SHEET_NAME_DATABASE}」にクリア対象のデータがありません。`, Browser.Buttons.OK);
        return;
      }
      // クリア範囲をD列からに修正
      const rangeToClear = dbSheet.getRange(2, DB_COL_TIME, lastRow - 1, DB_COL_AFTERSCHOOL - DB_COL_TIME + 1);
      rangeToClear.clearContent();
      Browser.msgBox(`データベースの入力内容をクリアしました。`, Browser.Buttons.OK);
      logInfo(`データベースクリア完了: ${rangeToClear.getA1Notation()}`);
    } catch (e) {
      logError("clearDatabaseDataWithConfirmation", e);
      Browser.msgBox(`クリアエラー: ${e.message}`, Browser.Buttons.OK);
    }
  } else {
    Browser.msgBox("クリア処理をキャンセルしました。", Browser.Buttons.OK);
  }
}

/** クラスルーム投稿用トリガーを設定します（時間指定）。*/
function setTriggers() {
  const ui = SpreadsheetApp.getUi();
  const functionNameToTrigger = "postScheduleToClassroom";
  const response = ui.prompt('トリガー時間設定', `「${functionNameToTrigger}」を実行する時間を0～23時の整数で入力してください (例: 15):`, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    const hour = parseInt(response.getResponseText(), 10);
    if (!isNaN(hour) && hour >= 0 && hour <= 23) {
      try {
        deleteTriggers_(functionNameToTrigger);
        ScriptApp.newTrigger(functionNameToTrigger).timeBased().everyDays(1).atHour(hour).create();
        logInfo(`トリガー作成: ${functionNameToTrigger} 毎日${hour}時`);
        ui.alert(`トリガー設定を完了しました。\n毎日${hour}時に投稿が実行されます。`);
      } catch (e) {
        logError("setTriggers", e);
        ui.alert(`トリガー設定エラー: ${e.message}\n(権限が不足している可能性があります)`);
      }
    } else {
      ui.alert(`入力が無効です。「${response.getResponseText()}」。0から23の整数で入力してください。`);
    }
  } else {
    ui.alert('トリガー設定をキャンセルしました。');
  }
}

// ============================================================
// ===== クラスルーム連携 関連関数 =====
// ============================================================

/** アカウント連携クラス一覧を初期設定シートに取得します。*/
function listCoursesToSheet() {
  try {
    let courses = [];
    let pageToken = null;
    do {
      const response = Classroom.Courses.list({ pageSize: 100, courseStates: ['ACTIVE'], pageToken: pageToken });
      if (response.courses) {
        courses = courses.concat(response.courses);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SETTINGS);
    if (!sheet) throw new Error(`シート「${SHEET_NAME_SETTINGS}」が見つかりません`);

    const startCell = sheet.getRange(SETTINGS_RANGE_COURSE_LIST_OUTPUT);
    const startRow = startCell.getRow();
    const startCol = startCell.getColumn();
    const lastRow = sheet.getLastRow();
    if (lastRow >= startRow) {
      sheet.getRange(startRow, startCol, lastRow - startRow + 1, 1).clearContent();
    }

    if (courses.length === 0) {
      startCell.setValue("（有効なコースが見つかりませんでした）");
      SpreadsheetApp.getUi().alert("有効なクラスが見つかりませんでした。");
    } else {
      const courseNames = courses.map(c => [c.name]);
      sheet.getRange(startRow, startCol, courseNames.length, 1).setValues(courseNames);
      SpreadsheetApp.getUi().alert('クラス一覧の取得が完了しました。');
    }
  } catch (e) {
    logError("listCoursesToSheet", e);
    SpreadsheetApp.getUi().alert(`クラス一覧取得エラー: ${e.message}\n（APIの有効化や権限を確認してください）`);
  }
}

/** データベースから翌日の予定を読み取り、Google Classroomにお知らせとして投稿します。 */
function postScheduleToClassroom() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
    const databaseSheet = ss.getSheetByName(SHEET_NAME_DATABASE);
    if (!settingSheet || !databaseSheet) throw new Error("必要シートなし");

    const courseName = settingSheet.getRange(SETTINGS_CELL_COURSE_NAME).getValue();
    if (!courseName) throw new Error(`${SETTINGS_CELL_COURSE_NAME} クラス名未設定`);
    const courseId = getCourseIdByName(courseName);
    if (!courseId) throw new Error(`クラス「${courseName}」見つからず`);

    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"];
    const formattedDateString = `${Utilities.formatDate(tomorrow, "JST", "yyyy/MM/dd")}（${daysOfWeek[tomorrow.getDay()]}）`;

    const dbData = databaseSheet.getDataRange().getValues();
    const foundRowData = dbData.find(row => row[DB_COL_DATE - 1] instanceof Date && isSameDate(row[DB_COL_DATE - 1], tomorrow) && row[DB_COL_PERIOD1 - 1]);

    if (!foundRowData) {
      Logger.log(`明日の予定なし/1校時空欄 スキップ`);
      return;
    }

    const schedule = {
      "朝学習": foundRowData[DB_COL_MORNING - 1], "1校時": foundRowData[DB_COL_PERIOD1 - 1], "単元1": foundRowData[DB_COL_UNIT1 - 1],
      "2校時": foundRowData[DB_COL_PERIOD2 - 1], "単元2": foundRowData[DB_COL_UNIT2 - 1], "3校時": foundRowData[DB_COL_PERIOD3 - 1],
      "単元3": foundRowData[DB_COL_UNIT3 - 1], "4校時": foundRowData[DB_COL_PERIOD4 - 1], "単元4": foundRowData[DB_COL_UNIT4 - 1],
      "5校時": foundRowData[DB_COL_PERIOD5 - 1], "単元5": foundRowData[DB_COL_UNIT5 - 1], "6校時": foundRowData[DB_COL_PERIOD6 - 1],
      "単元6": foundRowData[DB_COL_UNIT6 - 1], "宿題": foundRowData[DB_COL_HOMEWORK - 1], "持ち物": foundRowData[DB_COL_ITEMS - 1]
    };

    let postText = `${formattedDateString} の予定\n\n`;
    if (schedule["朝学習"]) postText += `朝学習：${schedule["朝学習"]}\n`;
    if (schedule["1校時"]) postText += `１時間目：${schedule["1校時"]} 「${schedule["単元1"] || ''}」\n`;
    if (schedule["2校時"]) postText += `２時間目：${schedule["2校時"]} 「${schedule["単元2"] || ''}」\n`;
    if (schedule["3校時"]) postText += `３時間目：${schedule["3校時"]} 「${schedule["単元3"] || ''}」\n`;
    if (schedule["4校時"]) postText += `４時間目：${schedule["4校時"]} 「${schedule["単元4"] || ''}」\n`;
    if (schedule["5校時"]) postText += `５時間目：${schedule["5校時"]} 「${schedule["単元5"] || ''}」\n`;
    if (schedule["6校時"]) postText += `６時間目：${schedule["6校時"]} 「${schedule["単元6"] || ''}」\n`;
    if (schedule["宿題"]) postText += `\n課題：\n${schedule["宿題"]}\n`;
    if (schedule["持ち物"]) postText += `\n持ち物：\n${schedule["持ち物"]}\n`;

    Classroom.Courses.Announcements.create({ text: postText.trim() }, courseId);
    logInfo(`クラス「${courseName}」へ予定投稿完了`);
  } catch (error) {
    logError("postScheduleToClassroom", error);
  }
}

/** 指定されたクラス名から、Google ClassroomのコースIDを探し出します。 */
function getCourseIdByName(courseName) {
  try {
    let pageToken = null;
    do {
      const response = Classroom.Courses.list({ pageSize: 100, courseStates: ['ACTIVE'], pageToken: pageToken });
      if (response.courses) {
        const course = response.courses.find(c => c.name === courseName);
        if (course) return course.id;
      }
      pageToken = response.nextPageToken;
    } while (pageToken);
    throw new Error(`クラス「${courseName}」が見つかりません`);
  } catch (e) {
    logError("getCourseIdByName", e);
    throw e;
  }
}

/** 「学級通信」シートをPDF化し、Google Classroomに投稿する一連の処理を実行します。 */
function autoPostToClassroom() {
  try {
    const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SETTINGS);
    if (!settingSheet) throw new Error(`シート「${SHEET_NAME_SETTINGS}」見つからず`);
    const classroomName = settingSheet.getRange(SETTINGS_CELL_COURSE_NAME).getValue();
    if (!classroomName) throw new Error(`${SETTINGS_CELL_COURSE_NAME} クラス名未設定`);
    const pdfFile = createAndSavePDF(SHEET_NAME_NEWSLETTER);
    if (!pdfFile) throw new Error("PDF作成/保存失敗");
    postToClassroomStream(classroomName, pdfFile);
    logInfo(`「${SHEET_NAME_NEWSLETTER}」PDFをクラス「${classroomName}」に投稿完了`);
  } catch (error) {
    logError("autoPostToClassroom", error);
  }
}

/** 指定されたシートをPDFとしてGoogleドライブに保存します。 */
function createAndSavePDF(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート「${sheetName}」見つからず`);
    const formattedDate = Utilities.formatDate(new Date(), "JST", "yyyyMMdd");
    const pdfFileName = `${sheetName}_${formattedDate}.pdf`;
    const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?` +
      `exportFormat=pdf&format=pdf&size=A4&portrait=true&fitToPage=true&gridlines=false&gid=${sheet.getSheetId()}`;
    const blob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(pdfFileName);
    const folder = DriveApp.getRootFolder();
    const existingFiles = folder.getFilesByName(pdfFileName);
    while (existingFiles.hasNext()) { existingFiles.next().setTrashed(true); }
    const file = folder.createFile(blob);
    logInfo(`PDF「${pdfFileName}」保存完了 (ID: ${file.getId()})`);
    return file;
  } catch (e) {
    logError(`createAndSavePDF (${sheetName})`, e);
    return null;
  }
}

/** 指定されたPDFファイルを、Google Classroomのストリームに投稿します。 */
function postToClassroomStream(classroomName, pdfFile) {
  try {
    const courseId = getCourseIdByName(classroomName);
    const announcement = { text: '学級通信', materials: [{ driveFile: { driveFile: { id: pdfFile.getId() } } }] };
    Classroom.Courses.Announcements.create(announcement, courseId);
    logInfo(`PDF(${pdfFile.getName()})をクラス「${classroomName}」に投稿`);
  } catch (e) {
    logError("postToClassroomStream", e);
    throw e;
  }
}

// ============================================================
// ===== API連携 & ログ機能 =====
// ============================================================

/** 「初期設定」シートから、Gemini APIを使用するためのAPIキーを取得します。 */
function getApiKey_() {
  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_SETTINGS);
  if (!settingsSheet) throw new Error(`シート「${SHEET_NAME_SETTINGS}」が見つかりません。`);
  const apiKey = settingsSheet.getRange(SETTINGS_CELL_GEMINI_API_KEY).getValue();
  if (!apiKey) throw new Error(`「${SHEET_NAME_SETTINGS}」シートのセル「${SETTINGS_CELL_GEMINI_API_KEY}」にGemini APIキーを入力してください。`);
  return apiKey;
}

/** 指定された指示（プロンプト）とPDFファイルを、Gemini APIに送信してAIによる分析を依頼します。 */
function callGeminiApi_(prompt, apiKey, blobs = []) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const parts = [{ "text": prompt }];
  blobs.forEach(blob => {
    parts.push({ "inline_data": { "mime_type": blob.getContentType(), "data": Utilities.base64Encode(blob.getBytes()) } });
  });
  const payload = { "contents": [{ "parts": parts }], "generationConfig": { "response_mime_type": "application/json" } };
  const options = { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true };
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseBody = response.getContentText();
  if (responseCode === 200) {
    const jsonResponse = JSON.parse(responseBody);
    if (jsonResponse.candidates && jsonResponse.candidates[0] && jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts) {
      const text = jsonResponse.candidates[0].content.parts[0].text;
      try {
        return JSON.parse(text);
      } catch (e) {
        logError("Gemini APIからのJSONレスポンスのパースに失敗しました。", e);
        logInfo(`パースに失敗したテキスト: ${text}`); // 失敗したテキストをログに記録
        return null; // 失敗した場合はnullを返す
      }
    } else {
      logError("Gemini APIからのレスポンス形式が不正です。", new Error(responseBody));
      return null;
    }
  } else {
    logError(`Gemini API Error (Code: ${responseCode})`, new Error(responseBody));
    throw new Error(`Gemini APIとの通信に失敗しました。レスポンスコード: ${responseCode}`);
  }
}

/** 「ログ」シートに、処理の成功や進捗を示す情報（INFO）を記録します。 */
function logInfo(message) { writeToLog_("INFO", message); }

/** 「ログ」シートに、発生したエラーの詳細を記録します。 */
function logError(message, error) {
  const errorMessage = `${message}\nエラー詳細: ${error.message}\nスタックトレース: ${error.stack}`;
  writeToLog_("ERROR", errorMessage);
}

/** 指定されたレベル（INFOやERROR）とメッセージを、「ログ」シートの最終行に書き込みます。 */
function writeToLog_(level, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(SHEET_NAME_LOG);
    if (!logSheet) {
      logSheet = ss.insertSheet(SHEET_NAME_LOG, ss.getSheets().length);
      logSheet.getRange("A1:C1").setValues([["日時", "レベル", "メッセージ"]]).setFontWeight("bold");
    }
    logSheet.appendRow([new Date(), level, message]);
  } catch (e) {
    console.error(`ログシートへの書き込みに失敗: ${e.message}`);
    console.error(`元のログ: [${level}] ${message}`);
  }
}

// ============================================================
// ===== ヘルパー関数群 =====
// ============================================================
/** 二つの日付が、年月日すべて同じ日であるかを判定します。 */
function isSameDate(date1, date2) { return date1.getFullYear() === date2.getFullYear() && date1.getMonth() === date2.getMonth() && date1.getDate() === date2.getDate(); }

/** ある日付が、指定された開始日と終了日の範囲内に含まれているかを判定します。 */
function isDateInRange(date, startDate, endDate) { const d = new Date(date); d.setHours(0, 0, 0, 0); return d.getTime() >= startDate.getTime() && d.getTime() <= endDate.getTime(); }

/** 日付を「yyyy/MM/dd」形式の文字列に変換します。 */
function formatDate(date) { if (!(date instanceof Date)) return ""; return Utilities.formatDate(date, "JST", "yyyy/MM/dd"); }

/** 指定された日付が含まれる週の、月曜日の日付を算出します。 */
function getMondayOfWeek(date) { const d = new Date(date); d.setHours(0, 0, 0, 0); const day = d.getDay(); const diff = d.getDate() - day + (day === 0 ? -6 : 1); return new Date(d.setDate(diff)); }

/** データベースシートの中から、指定された日付が入力されている行の番号を探し出します。 */
function findRowIndexByDate(sheet, dateToSearch) {
  const searchTime = getMondayOfWeek(dateToSearch).getTime();
  const dateColumnValues = sheet.getRange(2, DB_COL_DATE, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < dateColumnValues.length; i++) {
    if (dateColumnValues[i][0] instanceof Date) {
      const cellTime = new Date(dateColumnValues[i][0]).getTime();
      if (cellTime === searchTime) return i + 2;
    }
  }
  return -1;
}
