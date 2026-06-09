/**
 * 社内共有管理システム v9.0 整理版
 * Google Sheets + Apps Script + AppSheet
 *
 * 基本運用は以下の3つです。
 * 1. createCompanySheets() : 初回作成・列構成再作成
 * 2. refreshAll()          : 一覧・要確認・集計・表示整備
 * 3. onEdit(e)             : 入力時の通知更新・チェックボックス整備・予約重複チェック
 *
 * v9.0 方針
 * - fixCarNumberColumnAsText / fixDailyReportCheckboxes / cleanDailyReportEmptyRows は不要
 * - 車番の文字列固定は setupSheet() 内で自動適用
 * - 日報チェックボックスは setupSheet() / refreshInputSheets() / onEdit() で自動整備
 * - 日報空行整理は refreshInputSheets() に統合
 * - createFilter エラー対策として safeCreateFilter() を使用
 * - refreshAll は重くなりすぎないよう、入力整備と集計生成を分離して実行
 */

/***********************
 * 1. 基本設定
 ***********************/

const STAFF_MEMBERS = [
  "阿部遼太郎",
  "須藤智子",
  "佐藤愛子",
  "三船和美",
  "今野貴史",
  "豊嶋光弘",
  "三浦康正",
  "斎藤直樹",
  "渡邊充",
  "加藤明",
  "伊豆秀次",
  "佐藤純一",
  "捧麻美",
  "菊池孝幸",
  "須田隆吉"
];

// 既読確認列。増やしたい場合はここだけ変更します。
const PERSONAL_MEMBERS = [
  "阿部遼太郎",
  "三船和美",
  "須藤智子"
];

const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";
const CALENDAR_ID_HEADER = "カレンダーID";
const DATA_ROWS = 500;

const DATE_HEADERS = [
  "日付",
  "日時",
  "開始日",
  "終了日",
  "車検期限",
  "保険期限",
  "投稿日",
  "登録日",
  "期限",
  "移動日"
];

const TIME_HEADERS = ["開始時刻", "終了時刻", "開始時間", "終了時間"];
const IMAGE_HEADERS = ["写真", "画像"];
const TEXT_FORCE_HEADERS = ["車番"];

const STATUS_LIST = [CLEAR_LABEL, "予定", "移動中", "進行中", "着工前", "施工中", "完了", "延期", "中止", "未対応", "対応中", "要確認"];
const PHONE_LIST = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
const IMPORTANCE_LIST = [CLEAR_LABEL, "高", "中", "低"];
const STAFF_LIST = [CLEAR_LABEL].concat(STAFF_MEMBERS);
const CAR_LIST = [
  CLEAR_LABEL,
  "4t",
  "軽トラ②",
  "8tユニック",
  "軽ワゴン③",
  "ローダー",
  "P.BOX①",
  "P.BOX②",
  "3t",
  "キャブ③",
  "軽トラ③",
  "2t①",
  "軽ワゴン⑥",
  "軽バン⑤",
  "軽ダンプ②"
];

const SHEET_HEADERS = {
  "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, CALENDAR_ID_HEADER],
  "工事予定": ["工事名", "現場", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "備考", CALENDAR_ID_HEADER],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "メモ"],
  "車検管理": ["車両名", "車番", "車検期限", "通知", "保険期限", "状態", "写真", READ_HEADER, ...PERSONAL_MEMBERS],
  "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考", CALENDAR_ID_HEADER],
  "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...PERSONAL_MEMBERS, CALENDAR_ID_HEADER],
  "行事予定": ["日付", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, CALENDAR_ID_HEADER],
  "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考"],
  "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "日報": ["日付", "担当", "現場", "作業内容", "進捗", "問題点", "明日の予定", "写真", "日報文章", "状態", "備考", READ_HEADER, ...PERSONAL_MEMBERS],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"],
  "担当別未読": ["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"],
  "既読率集計": ["担当", "対象件数", "既読件数", "未読件数", "既読率"],
  "過去一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート", "備考", "移動日"],
  "ダッシュボード": ["項目", "件数"]
};

const INPUT_SHEETS = [
  "出先予定",
  "工事予定",
  "電話履歴",
  "車検管理",
  "社用車予約",
  "会議予定",
  "行事予定",
  "作業状況",
  "お知らせ",
  "個人ToDo",
  "日報"
];

const GROUP_TARGET_SHEETS = INPUT_SHEETS.concat(["一覧スケジュール", "過去一覧"]);

/***********************
 * 2. 基本運用関数
 ***********************/

function createCompanySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEET_HEADERS).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    resetSheetStructure_(sheet, SHEET_HEADERS[sheetName]);
    setupSheet(sheet, SHEET_HEADERS[sheetName]);
  });

  refreshAll();
  ss.toast("v9.0 整理版のシート作成が完了しました");
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    refreshInputSheets();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    rebuildCheckGroups();
    groupArchiveByMonthWithHeader();
    SpreadsheetApp.getActiveSpreadsheet().toast("全体更新が完了しました");
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row === 1) return;

  if (e.range.getValue() === CLEAR_LABEL) {
    e.range.clearContent();
    return;
  }

  updateNoticeForEditedRow(sheet, row, col);
  updateCheckboxesForEditedRow(sheet, row);
  normalizeEditedRow_(sheet, row);

  if (sheetName === "社用車予約") {
    checkCarReservationConflictForRow_(sheet, row, true);
  }

  // 入力シートを編集した時だけ一覧系を更新します。
  if (INPUT_SHEETS.includes(sheetName)) {
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
  }
}

/***********************
 * 3. 初期化・シート整備
 ***********************/

function resetSheetStructure_(sheet, headers) {
  removeFilter_(sheet);
  resetColumnGroupsForSheet_(sheet);
  resetRowGroups(sheet);

  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
  .clearDataValidations();
  sheet.setConditionalFormatRules([]);

  const targetCols = headers.length;
  const currentCols = sheet.getMaxColumns();

  if (currentCols > targetCols) {
    sheet.deleteColumns(targetCols + 1, currentCols - targetCols);
  } else if (currentCols < targetCols) {
    sheet.insertColumnsAfter(currentCols, targetCols - currentCols);
  }
}

function setupSheet(sheet, headers) {
  const lastCol = headers.length;

  sheet.getRange(1, 1, 1, lastCol).setValues([headers]);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  headers.forEach((header, index) => {
    const col = index + 1;

    if (DATE_HEADERS.includes(header)) setDateColumn(sheet, col);
    if (TIME_HEADERS.includes(header)) setTimeColumn_(sheet, col);
    if (TEXT_FORCE_HEADERS.includes(header)) setTextColumn_(sheet, col);
    if (IMAGE_HEADERS.includes(header)) setImageColumn_(sheet, col);

    if (header === "状態") setDropdown(sheet, col, STATUS_LIST);
    if (header === "担当" || header === "利用者" || header === "投稿者") setDropdown(sheet, col, STAFF_LIST);
    if (header === "社用車") setDropdown(sheet, col, CAR_LIST);
    if (header === "電話対応") setDropdown(sheet, col, PHONE_LIST);
    if (header === "重要度") setDropdown(sheet, col, IMPORTANCE_LIST);
  });

  setCheckboxesForDataRows(sheet);
  applyColorRules(sheet);
  safeCreateFilter(sheet, lastCol);
  hideSystemColumns_(sheet, headers);
  sheet.autoResizeColumns(1, lastCol);
}

function refreshInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  INPUT_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = getHeaders_(sheet);
    if (headers.length === 0) return;

    normalizeSheetRows_(sheet);
    setCheckboxesForDataRows(sheet);
    applyColumnFormats_(sheet, headers);
    applyColorRules(sheet);
    safeCreateFilter(sheet, headers.length);
    hideSystemColumns_(sheet, headers);
  });
}

function normalizeSheetRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = headers.length;
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowsToDelete = [];

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const hasData = row.some(value => value !== "" && value !== null && value !== undefined);

    if (!hasData && sheet.getName() === "日報") {
      rowsToDelete.push(rowNumber);
      return;
    }

    if (hasData) {
      const noticeCol = headers.indexOf("通知") + 1;
      const dateCol = findMainDateColumn(headers);
      if (noticeCol > 0 && dateCol > 0) {
        const dateValue = row[dateCol - 1];
        sheet.getRange(rowNumber, noticeCol).setValue(getNoticeText(dateValue));
      }
    }
  });

  rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
}

function normalizeEditedRow_(sheet, row) {
  const headers = getHeaders_(sheet);
  applyColumnFormats_(sheet, headers);

  const noticeCol = headers.indexOf("通知") + 1;
  const dateCol = findMainDateColumn(headers);
  if (noticeCol > 0 && dateCol > 0) {
    sheet.getRange(row, noticeCol).setValue(getNoticeText(sheet.getRange(row, dateCol).getValue()));
  }
}

function applyColumnFormats_(sheet, headers) {
  headers.forEach((header, index) => {
    const col = index + 1;
    if (DATE_HEADERS.includes(header)) setDateColumn(sheet, col);
    if (TIME_HEADERS.includes(header)) setTimeColumn_(sheet, col);
    if (TEXT_FORCE_HEADERS.includes(header)) setTextColumn_(sheet, col);
    if (IMAGE_HEADERS.includes(header)) setImageColumn_(sheet, col);
  });
}

function safeCreateFilter(sheet, lastCol) {
  if (!sheet || lastCol <= 0) return;

  try {
    const existing = sheet.getFilter();
    if (existing) return;

    const maxRows = Math.max(sheet.getMaxRows(), 2);
    sheet.getRange(1, 1, maxRows, lastCol).createFilter();
  } catch (error) {
    console.log("createFilter skipped: " + sheet.getName() + " / " + error);
  }
}

function removeFilter_(sheet) {
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (error) {
    console.log("removeFilter skipped: " + sheet.getName() + " / " + error);
  }
}

function setDateColumn(sheet, col) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, DATA_ROWS, 1)
    .setNumberFormat("yyyy/mm/dd")
    .setDataValidation(rule);
}

function setTimeColumn_(sheet, col) {
  sheet.getRange(2, col, DATA_ROWS, 1).setNumberFormat("hh:mm");
}

function setTextColumn_(sheet, col) {
  sheet.getRange(2, col, DATA_ROWS, 1).setNumberFormat("@");
}

function setImageColumn_(sheet, col) {
  // AppSheet 側で Type=Image にする列です。Sheets 側ではリンク・画像URLが入っても崩れにくいよう折返し表示にします。
  sheet.getRange(2, col, DATA_ROWS, 1).setWrap(true);
}

function setDropdown(sheet, col, list) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, DATA_ROWS, 1).setDataValidation(rule);
}

function setCheckboxesForDataRows(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = getHeaders_(sheet);
  const checkNames = [READ_HEADER].concat(PERSONAL_MEMBERS);
  const lastRow = sheet.getLastRow();

  checkNames.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;

    const dataWidth = Math.max(col - 1, 1);
    const dataValues = sheet.getRange(2, 1, lastRow - 1, dataWidth).getValues();
    const checkRange = sheet.getRange(2, col, lastRow - 1, 1);
    const currentValues = checkRange.getValues();

    const output = currentValues.map((row, index) => {
      const hasData = dataValues[index].some(value => value !== "" && value !== null && value !== undefined);
      if (!hasData) return [""];
      return [toBoolean_(row[0])];
    });

    checkRange.setValues(output);
    checkRange.insertCheckboxes();

    output.forEach((row, index) => {
      if (row[0] === "") {
        const cell = sheet.getRange(index + 2, col);
        cell.clearDataValidations();
        cell.clearContent();
      }
    });
  });
}

function updateCheckboxesForEditedRow(sheet, row) {
  const headers = getHeaders_(sheet);
  const readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol <= 0) return;

  const dataValues = sheet.getRange(row, 1, 1, Math.max(readCol - 1, 1)).getValues()[0];
  const hasData = dataValues.some(value => value !== "" && value !== null && value !== undefined);

  [READ_HEADER].concat(PERSONAL_MEMBERS).forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;

    const cell = sheet.getRange(row, col);
    if (hasData) {
      cell.insertCheckboxes();
      if (cell.getValue() === "") cell.setValue(false);
    } else {
      cell.clearDataValidations();
      cell.clearContent();
    }
  });
}

function hideSystemColumns_(sheet, headers) {
  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    try {
      sheet.hideColumns(calendarIdCol);
    } catch (error) {}
  }
}

/***********************
 * 4. 一覧・要確認・集計
 ***********************/

function createScheduleList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, "一覧スケジュール");
  const headers = SHEET_HEADERS["一覧スケジュール"];

  resetSheetStructure_(sheet, headers);
  setupSheet(sheet, headers);

  const rows = [];

  addScheduleRows_(ss, "出先予定", row => row["日付"], row => ({
    "日付": row["日付"],
    "種類": "出先予定",
    "内容": joinText(row["行き先"], row["用件"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": row["電話対応"],
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "出先予定"
  }), rows);

  addScheduleRows_(ss, "工事予定", row => row["開始日"] || row["工事名"], row => ({
    "日付": row["開始日"] || new Date(),
    "種類": "工事予定",
    "内容": joinText(row["工事名"], row["現場"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["開始日"]),
    "電話対応": row["電話対応"],
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "工事予定"
  }), rows);

  addScheduleRows_(ss, "電話履歴", row => row["日時"] || row["相手"] || row["内容"], row => ({
    "日付": row["日時"] || new Date(),
    "種類": "電話履歴",
    "内容": joinText(row["相手"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日時"]),
    "電話対応": row["電話対応"],
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "電話履歴"
  }), rows);

  addScheduleRows_(ss, "車検管理", row => row["車検期限"] || row["車両名"], row => ({
    "日付": row["車検期限"] || new Date(),
    "種類": "車検",
    "内容": joinText(row["車両名"], row["車番"]),
    "担当": "",
    "状態": row["状態"],
    "通知": getNoticeText(row["車検期限"]),
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "車検管理"
  }), rows);

  addScheduleRows_(ss, "社用車予約", row => row["日付"] || row["社用車"] || row["行き先"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "社用車予約",
    "内容": joinText(row["社用車"], joinText(row["行き先"], row["用途"])),
    "担当": row["利用者"],
    "状態": row["状態"],
    "通知": row["通知"] === "予約重複" ? "予約重複" : getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "社用車予約"
  }), rows);

  addScheduleRows_(ss, "会議予定", row => row["日付"] || row["会議名"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "会議",
    "内容": joinText(row["会議名"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "会議予定"
  }), rows);

  addScheduleRows_(ss, "行事予定", row => row["日付"] || row["行事名"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "行事",
    "内容": joinText(row["行事名"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "行事予定"
  }), rows);

  addScheduleRows_(ss, "作業状況", row => row["現場"] || row["作業内容"], row => ({
    "日付": new Date(),
    "種類": "作業状況",
    "内容": joinText(row["現場"], row["作業内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": row["通知"] || "",
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "作業状況"
  }), rows);

  addScheduleRows_(ss, "個人ToDo", row => row["期限"] || row["内容"], row => ({
    "日付": row["期限"] || row["登録日"] || new Date(),
    "種類": "ToDo",
    "内容": row["内容"],
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["期限"]),
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "個人ToDo"
  }), rows);

  addScheduleRows_(ss, "日報", row => row["日付"] || row["現場"] || row["作業内容"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "日報",
    "内容": joinText(row["現場"], row["作業内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": toBoolean_(row[READ_HEADER]),
    ...copyPersonalChecks(row),
    "元シート": "日報"
  }), rows);

  rows.sort((a, b) => new Date(a["日付"]) - new Date(b["日付"]));
  writeObjectsToSheet(sheet, rows);
}

function createAlertList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("一覧スケジュール");
  const sheet = getOrCreateSheet_(ss, "要確認一覧");
  const headers = SHEET_HEADERS["要確認一覧"];

  resetSheetStructure_(sheet, headers);
  setupSheet(sheet, headers);

  const priority = {"予約重複": 0, "重要": 0, "期限切れ": 1, "今日": 2, "3日以内": 3, "7日以内": 4};
  const rows = [];

  if (source && source.getLastRow() >= 2) {
    getSheetObjects_(source).forEach(row => {
      if (priority[row["通知"]] !== undefined || row["状態"] === "要確認" || row["電話対応"] === "未対応" || row["電話対応"] === "折返し") {
        rows.push(row);
      }
    });
  }

  const noticeSheet = ss.getSheetByName("お知らせ");
  if (noticeSheet && noticeSheet.getLastRow() >= 2) {
    getSheetObjects_(noticeSheet).forEach(row => {
      if (row["重要度"] === "高") {
        rows.unshift({
          "日付": row["投稿日"] || new Date(),
          "種類": "お知らせ",
          "内容": joinText(row["タイトル"], row["内容"]),
          "担当": row["投稿者"],
          "状態": "要確認",
          "通知": "重要",
          "電話対応": "",
          "元シート": "お知らせ"
        });
      }
    });
  }

  rows.sort((a, b) => {
    const pa = priority[a["通知"]] === undefined ? 99 : priority[a["通知"]];
    const pb = priority[b["通知"]] === undefined ? 99 : priority[b["通知"]];
    if (pa !== pb) return pa - pb;
    return new Date(a["日付"]) - new Date(b["日付"]);
  });

  writeObjectsToSheet(sheet, rows.map(row => ({
    "日付": row["日付"],
    "種類": row["種類"],
    "内容": row["内容"],
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": row["通知"],
    "電話対応": row["電話対応"],
    "元シート": row["元シート"]
  })));
}

function createAssigneeUnreadSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, "担当別未読");
  const headers = SHEET_HEADERS["担当別未読"];

  resetSheetStructure_(sheet, headers);
  setupSheet(sheet, headers);

  const schedule = ss.getSheetByName("一覧スケジュール");
  const todo = ss.getSheetByName("個人ToDo");
  const phone = ss.getSheetByName("電話履歴");

  const rows = PERSONAL_MEMBERS.map(member => ({
    "担当": member,
    "未読件数": countUnreadByMember(schedule, member),
    "未完了ToDo": countTodoByMember(todo, member),
    "未対応電話": countPhoneByMember(phone, member),
    "今日の予定": countTodayScheduleByMember(schedule, member)
  }));

  writeObjectsToSheet(sheet, rows);
}

function createReadRateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, "既読率集計");
  const headers = SHEET_HEADERS["既読率集計"];

  resetSheetStructure_(sheet, headers);
  setupSheet(sheet, headers);

  const rows = PERSONAL_MEMBERS.map(member => {
    let total = 0;
    let read = 0;

    GROUP_TARGET_SHEETS.forEach(sheetName => {
      const target = ss.getSheetByName(sheetName);
      if (!target || target.getLastRow() < 2) return;

      const targetHeaders = getHeaders_(target);
      const checkCol = targetHeaders.indexOf(member) + 1;
      if (checkCol <= 0) return;

      getSheetValues_(target).forEach(row => {
        if (row[1] === "月見出し") return;
        const hasData = row.some(value => value !== "" && value !== null && value !== undefined);
        if (!hasData) return;
        total++;
        if (row[checkCol - 1] === true) read++;
      });
    });

    const unread = total - read;
    const rate = total === 0 ? 0 : read / total;
    return {"担当": member, "対象件数": total, "既読件数": read, "未読件数": unread, "既読率": rate};
  });

  writeObjectsToSheet(sheet, rows);
  if (rows.length > 0) sheet.getRange(2, 5, rows.length, 1).setNumberFormat("0.0%");
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet_(ss, "ダッシュボード");

  removeFilter_(sheet);
  sheet.clear();
  sheet.clearFormats();

  sheet.getRange("A1").setValue("社内共有ダッシュボード").setFontSize(22).setFontWeight("bold");
  sheet.getRange("A3:B3").setValues([["項目", "件数"]]);
  sheet.getRange("A3:B3").setBackground("#d9ead3").setFontWeight("bold").setHorizontalAlignment("center");

  const data = [
    ["今日の予定", '=COUNTIF(一覧スケジュール!A2:A,TODAY())'],
    ["期限切れ", '=COUNTIF(一覧スケジュール!F2:F,"期限切れ")'],
    ["今日対応", '=COUNTIF(一覧スケジュール!F2:F,"今日")'],
    ["3日以内", '=COUNTIF(一覧スケジュール!F2:F,"3日以内")'],
    ["7日以内", '=COUNTIF(一覧スケジュール!F2:F,"7日以内")'],
    ["予約重複", '=COUNTIF(一覧スケジュール!F2:F,"予約重複")'],
    ["未読件数", '=COUNTIF(一覧スケジュール!H2:H,FALSE)'],
    ["未対応電話", '=COUNTIF(一覧スケジュール!G2:G,"未対応")'],
    ["折返し電話", '=COUNTIF(一覧スケジュール!G2:G,"折返し")'],
    ["完了件数", '=COUNTIF(一覧スケジュール!E2:E,"完了")'],
    ["日報件数", '=COUNTA(日報!A2:A)'],
    ["担当別未読あり", '=COUNTIF(担当別未読!B2:B,">0")']
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);
  sheet.getRange(3, 1, data.length + 1, 2).setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 2);
}

/***********************
 * 5. 日報 / GPT / PDF
 ***********************/

function createDailyReportText() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet || sheet.getName() !== "日報") {
    SpreadsheetApp.getUi().alert("日報シートで、文章を作成したい行を選択してください。");
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("見出し行ではなく、日報データの行を選択してください。");
    return;
  }

  const headers = getHeaders_(sheet);
  const data = objectFromRow(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
  const outputCol = headers.indexOf("日報文章") + 1;
  if (outputCol <= 0) return;

  const prompt = buildDailyReportPrompt_(data);
  const text = callOpenAI(prompt);
  if (!text) return;

  sheet.getRange(row, outputCol).setValue(text).setWrap(true);
  SpreadsheetApp.getActiveSpreadsheet().toast("日報文章を作成しました");
}

function buildDailyReportPrompt_(data) {
  return [
    "次の情報をもとに、社内向けの日報文章を作成してください。",
    "",
    "日付: " + formatDateForReport_(data["日付"]),
    "担当: " + (data["担当"] || ""),
    "現場: " + (data["現場"] || ""),
    "作業内容: " + (data["作業内容"] || ""),
    "進捗: " + (data["進捗"] || ""),
    "問題点: " + (data["問題点"] || ""),
    "明日の予定: " + (data["明日の予定"] || ""),
    "備考: " + (data["備考"] || ""),
    "",
    "条件:",
    "・丁寧な社内向け文章にする",
    "・事実を勝手に追加しない",
    "・長すぎず読みやすくする",
    "・空欄の項目は無理に書かない"
  ].join("\n");
}

function callOpenAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("OPENAI_API_KEY が未設定です。スクリプトプロパティにAPIキーを保存してください。");
    return "";
  }

  const payload = {model: "gpt-4.1-mini", input: prompt};
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    contentType: "application/json",
    headers: {Authorization: "Bearer " + apiKey},
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();
  if (statusCode < 200 || statusCode >= 300) {
    console.log(body);
    SpreadsheetApp.getUi().alert("OpenAI APIでエラーが発生しました。ステータス: " + statusCode);
    return "";
  }

  const json = JSON.parse(body);
  if (json.output_text) return json.output_text.trim();

  const parts = [];
  if (json.output) {
    json.output.forEach(item => {
      if (!item.content) return;
      item.content.forEach(content => {
        if (content.text) parts.push(content.text);
      });
    });
  }
  return parts.join("\n").trim();
}

function createSelectedDailyReportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== "日報") {
    SpreadsheetApp.getUi().alert("日報シートでPDF化したい行を選択してください。");
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("見出し行ではなく、日報データの行を選択してください。");
    return;
  }

  const headers = getHeaders_(sheet);
  const data = objectFromRow(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
  const dateText = formatDateForReport_(data["日付"]) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const fileDateText = dateText.replace(/\//g, "");
  const staffText = sanitizeFileName_(data["担当"] || "担当未設定");

  const doc = DocumentApp.create("日報_" + fileDateText + "_" + staffText);
  const body = doc.getBody();

  body.appendParagraph("作業日報").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  appendDailyReportParagraph_(body, "日付", dateText);
  appendDailyReportParagraph_(body, "担当", data["担当"]);
  appendDailyReportParagraph_(body, "現場", data["現場"]);
  appendDailyReportParagraph_(body, "作業内容", data["作業内容"]);
  appendDailyReportParagraph_(body, "進捗", data["進捗"]);
  appendDailyReportParagraph_(body, "問題点", data["問題点"]);
  appendDailyReportParagraph_(body, "明日の予定", data["明日の予定"]);
  appendDailyReportParagraph_(body, "日報文章", data["日報文章"]);
  appendDailyReportParagraph_(body, "備考", data["備考"]);

  doc.saveAndClose();

  const pdfFile = saveDocAsPdfNearSpreadsheet_(ss, doc.getId(), "日報_" + fileDateText + "_" + staffText + ".pdf");
  SpreadsheetApp.getActiveSpreadsheet().toast("選択行の日報PDFを作成しました");
  return pdfFile.getUrl();
}

function createDailyReportPdf() {
  refreshAll();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const dateText = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy年M月d日");
  const fileDateText = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
  const doc = DocumentApp.create("日報サマリー_" + fileDateText);
  const body = doc.getBody();

  body.appendParagraph("社内共有 日報サマリー").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(dateText);
  body.appendParagraph("");

  appendDailyReportSection(body, "本日の予定", getTodayScheduleRows_());
  appendDailyReportSection(body, "要確認", getAlertRows_());
  appendDailyReportSection(body, "未対応・折返し電話", getPhoneAlertRows_());
  appendDailyReportSection(body, "担当別未読", getAssigneeUnreadRows_());

  doc.saveAndClose();
  const pdfFile = saveDocAsPdfNearSpreadsheet_(ss, doc.getId(), "日報サマリー_" + fileDateText + ".pdf");
  SpreadsheetApp.getActiveSpreadsheet().toast("当日サマリーPDFを作成しました");
  return pdfFile.getUrl();
}

function appendDailyReportParagraph_(body, title, value) {
  if (value === "" || value === null || value === undefined) return;
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(String(value));
}

function appendDailyReportSection(body, title, rows) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (!rows || rows.length === 0) {
    body.appendParagraph("該当なし");
    return;
  }
  body.appendTable(rows);
  body.appendParagraph("");
}

function saveDocAsPdfNearSpreadsheet_(ss, docId, pdfName) {
  const docFile = DriveApp.getFileById(docId);
  const pdfBlob = docFile.getBlob().getAs(MimeType.PDF).setName(pdfName);
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parents = spreadsheetFile.getParents();
  let pdfFile;

  if (parents.hasNext()) {
    pdfFile = parents.next().createFile(pdfBlob);
  } else {
    pdfFile = DriveApp.createFile(pdfBlob);
  }

  docFile.setTrashed(true);
  return pdfFile;
}

function getTodayScheduleRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("一覧スケジュール");
  const result = [["日付", "種類", "内容", "担当", "状態", "通知"]];
  if (!sheet || sheet.getLastRow() < 2) return [];

  getSheetObjects_(sheet).forEach(row => {
    if (isSameDateOnly_(row["日付"], new Date())) {
      result.push([formatDateForReport_(row["日付"]), row["種類"] || "", row["内容"] || "", row["担当"] || "", row["状態"] || "", row["通知"] || ""]);
    }
  });

  return result.length === 1 ? [] : result;
}

function getAlertRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("要確認一覧");
  const result = [["日付", "種類", "内容", "担当", "通知"]];
  if (!sheet || sheet.getLastRow() < 2) return [];

  getSheetObjects_(sheet).forEach(row => {
    result.push([formatDateForReport_(row["日付"]), row["種類"] || "", row["内容"] || "", row["担当"] || "", row["通知"] || ""]);
  });

  return result.length === 1 ? [] : result;
}

function getPhoneAlertRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("電話履歴");
  const result = [["日時", "相手", "内容", "担当", "電話対応", "状態"]];
  if (!sheet || sheet.getLastRow() < 2) return [];

  getSheetObjects_(sheet).forEach(row => {
    if (row["電話対応"] === "未対応" || row["電話対応"] === "折返し") {
      result.push([formatDateForReport_(row["日時"]), row["相手"] || "", row["内容"] || "", row["担当"] || "", row["電話対応"] || "", row["状態"] || ""]);
    }
  });

  return result.length === 1 ? [] : result;
}

function getAssigneeUnreadRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("担当別未読");
  const result = [["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"]];
  if (!sheet || sheet.getLastRow() < 2) return [];

  sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues().forEach(row => {
    result.push([row[0] || "", String(row[1] || 0), String(row[2] || 0), String(row[3] || 0), String(row[4] || 0)]);
  });

  return result.length === 1 ? [] : result;
}

/***********************
 * 6. Googleカレンダー連携
 ***********************/

function syncCalendarEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendar = CalendarApp.getDefaultCalendar();

  const targets = [
    {
      sheetName: "出先予定",
      title: row => "出先予定: " + joinText(row["行き先"], row["用件"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || ""
    },
    {
      sheetName: "会議予定",
      title: row => "会議: " + joinText(row["会議名"], row["内容"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || ""
    },
    {
      sheetName: "行事予定",
      title: row => "行事: " + joinText(row["行事名"], row["内容"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || ""
    },
    {
      sheetName: "工事予定",
      title: row => "工事: " + joinText(row["工事名"], row["現場"]),
      startDate: row => row["開始日"],
      endDate: row => row["終了日"] || row["開始日"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || ""
    },
    {
      sheetName: "社用車予約",
      title: row => "社用車予約: " + joinText(row["社用車"], row["行き先"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"],
      endTime: row => row["終了時刻"]
    }
  ];

  targets.forEach(target => syncCalendarTarget_(ss, calendar, target));
  ss.toast("Googleカレンダーへ反映しました");
}

function syncCalendarTarget_(ss, calendar, target) {
  const sheet = ss.getSheetByName(target.sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = getHeaders_(sheet);
  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol <= 0) return;

  getSheetObjects_(sheet).forEach((row, index) => {
    const startDateValue = target.startDate(row);
    if (!startDateValue) return;

    const title = target.title(row);
    if (!title || title.endsWith(": ")) return;

    const oldEventId = row[CALENDAR_ID_HEADER];
    if (oldEventId) {
      try {
        const oldEvent = calendar.getEventById(oldEventId);
        if (oldEvent) oldEvent.deleteEvent();
      } catch (error) {}
    }

    const description = buildCalendarDescription(row, target.sheetName);
    const startTimeValue = target.startTime(row);
    const endTimeValue = target.endTime(row);
    let event = null;

    if (startTimeValue && endTimeValue) {
      const startDateTime = buildDateTime(startDateValue, startTimeValue);
      const endDateTime = buildDateTime(target.endDate(row) || startDateValue, endTimeValue);
      if (startDateTime && endDateTime && endDateTime.getTime() > startDateTime.getTime()) {
        event = calendar.createEvent(title, startDateTime, endDateTime, {description: description});
      }
    }

    if (!event) {
      const start = toDateOnly(startDateValue);
      const end = toDateOnly(target.endDate(row) || startDateValue);
      if (end.getTime() > start.getTime()) {
        const endPlusOne = new Date(end);
        endPlusOne.setDate(endPlusOne.getDate() + 1);
        event = calendar.createAllDayEvent(title, start, endPlusOne, {description: description});
      } else {
        event = calendar.createAllDayEvent(title, start, {description: description});
      }
    }

    sheet.getRange(index + 2, calendarIdCol).setValue(event.getId());
  });
}

function buildCalendarDescription(row, sourceName) {
  const lines = ["元シート: " + sourceName];
  Object.keys(row).forEach(key => {
    if (key === CALENDAR_ID_HEADER) return;
    if (row[key] === "" || row[key] === null || row[key] === undefined) return;
    lines.push(key + ": " + row[key]);
  });
  return lines.join("\n");
}

/***********************
 * 7. 社用車予約重複
 ***********************/

function checkCarReservationConflicts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("社用車予約");
  if (!sheet || sheet.getLastRow() < 2) return;

  let count = 0;
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    if (checkCarReservationConflictForRow_(sheet, row, false)) count++;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("社用車予約の重複チェック完了: " + count + "件");
}

function checkCarReservationConflictForRow_(sheet, targetRow, showAlert) {
  const message = getCarReservationConflictMessage(sheet, targetRow);
  const headers = getHeaders_(sheet);
  const noticeCol = headers.indexOf("通知") + 1;
  const dateCol = headers.indexOf("日付") + 1;

  if (noticeCol > 0) {
    if (message) {
      sheet.getRange(targetRow, noticeCol).setValue("予約重複");
    } else {
      const current = sheet.getRange(targetRow, noticeCol).getValue();
      if (current === "予約重複") {
        const dateValue = dateCol > 0 ? sheet.getRange(targetRow, dateCol).getValue() : "";
        sheet.getRange(targetRow, noticeCol).setValue(getNoticeText(dateValue));
      }
    }
  }

  if (message && showAlert) SpreadsheetApp.getUi().alert(message);
  return Boolean(message);
}

function getCarReservationConflictMessage(sheet, targetRow) {
  if (!sheet || sheet.getName() !== "社用車予約" || sheet.getLastRow() < 3) return "";

  const headers = getHeaders_(sheet);
  const target = objectFromRow(headers, sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0]);
  if (!target["日付"] || !target["開始時刻"] || !target["終了時刻"] || !target["社用車"]) return "";
  if (target["状態"] === "中止") return "";

  const targetStart = buildDateTime(target["日付"], target["開始時刻"]);
  const targetEnd = buildDateTime(target["日付"], target["終了時刻"]);
  if (!targetStart || !targetEnd || targetEnd.getTime() <= targetStart.getTime()) return "";

  const rows = getSheetObjects_(sheet);
  for (let i = 0; i < rows.length; i++) {
    const rowNumber = i + 2;
    if (rowNumber === targetRow) continue;

    const row = rows[i];
    if (row["状態"] === "中止") continue;
    if (row["社用車"] !== target["社用車"]) continue;

    const otherStart = buildDateTime(row["日付"], row["開始時刻"]);
    const otherEnd = buildDateTime(row["日付"], row["終了時刻"]);
    if (!otherStart || !otherEnd) continue;

    const overlap = targetStart.getTime() < otherEnd.getTime() && otherStart.getTime() < targetEnd.getTime();
    if (overlap) {
      return "社用車予約が重複しています。\n\n車両: " + target["社用車"] + "\n重複行: " + rowNumber + "行目";
    }
  }

  return "";
}

/***********************
 * 8. 過去一覧・月別グループ
 ***********************/

function movePastItemsToArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archiveSheet = getOrCreateSheet_(ss, "過去一覧");
  setupSheet(archiveSheet, SHEET_HEADERS["過去一覧"]);

  const today = toDateOnly(new Date());
  const archiveRows = [];

  ["出先予定", "会議予定", "行事予定", "工事予定", "電話履歴", "社用車予約", "個人ToDo", "日報"].forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const headers = getHeaders_(sheet);
    const dateCol = findMainDateColumn(headers);
    if (dateCol <= 0) return;

    const rowsToDelete = [];
    getSheetValues_(sheet).forEach((valuesRow, index) => {
      const rowNumber = index + 2;
      const row = objectFromRow(headers, valuesRow);
      const dateValue = valuesRow[dateCol - 1];
      if (!dateValue) return;

      const date = toDateOnly(dateValue);
      if (isNaN(date.getTime())) return;

      if (date.getTime() < today.getTime()) {
        archiveRows.push(buildArchiveRow(row, sheetName));
        rowsToDelete.push(rowNumber);
      }
    });

    rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
  });

  appendObjectsToSheet(archiveSheet, archiveRows);
  refreshAll();
  SpreadsheetApp.getActiveSpreadsheet().toast("過去一覧へ移動しました: " + archiveRows.length + "件");
}

function buildArchiveRow(row, sheetName) {
  const base = {
    "日付": row["日付"] || row["日時"] || row["開始日"] || row["期限"] || "",
    "種類": sheetName,
    "内容": row["内容"] || joinText(row["工事名"] || row["現場"] || row["相手"] || row["車両名"] || row["社用車"], row["用件"] || row["作業内容"] || row["行き先"] || row["用途"] || ""),
    "担当": row["担当"] || row["利用者"] || "",
    "状態": row["状態"] || "",
    "通知": row["通知"] || "",
    "電話対応": row["電話対応"] || "",
    "元シート": sheetName,
    "備考": row["日報文章"] || row["備考"] || row["メモ"] || "",
    "移動日": new Date()
  };
  return base;
}

function groupArchiveByMonthWithHeader() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("過去一覧");
  if (!sheet || sheet.getLastRow() < 2) return;

  removeArchiveMonthHeaderRows(sheet);
  resetRowGroups(sheet);

  const headers = getHeaders_(sheet);
  const lastCol = headers.length;
  const dataRows = getSheetValues_(sheet).filter(row => row[0] && row[1] !== "月見出し");

  if (sheet.getLastRow() >= 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).clearContent();
  if (dataRows.length === 0) return;

  dataRows.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  const outputRows = [];
  let currentKey = "";

  dataRows.forEach(row => {
    const key = getYearMonthKey(row[0]);
    if (!key) return;

    if (key !== currentKey) {
      const headerRow = new Array(lastCol).fill("");
      headerRow[0] = getYearMonthLabel(row[0]);
      headerRow[1] = "月見出し";
      outputRows.push(headerRow);
      currentKey = key;
    }
    outputRows.push(row);
  });

  sheet.getRange(2, 1, outputRows.length, lastCol).setValues(outputRows);

  let groupStart = 0;
  outputRows.forEach((row, index) => {
    const rowNumber = index + 2;
    if (row[1] === "月見出し") {
      if (groupStart > 0 && rowNumber > groupStart) {
        sheet.getRange(groupStart, 1, rowNumber - groupStart, 1).shiftRowGroupDepth(1);
      }
      sheet.getRange(rowNumber, 1, 1, lastCol).clearDataValidations().setBackground("#d9ead3").setFontWeight("bold");
      groupStart = rowNumber + 1;
    }
  });

  const finalLastRow = sheet.getLastRow();
  if (groupStart > 0 && finalLastRow >= groupStart) {
    sheet.getRange(groupStart, 1, finalLastRow - groupStart + 1, 1).shiftRowGroupDepth(1);
  }

  applyColorRules(sheet);
}

function removeArchiveMonthHeaderRows(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const rowsToDelete = [];
  getSheetValues_(sheet).forEach((row, index) => {
    if (row[1] === "月見出し") rowsToDelete.push(index + 2);
  });
  rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
}

function resetRowGroups(sheet) {
  if (!sheet) return;
  try { sheet.expandAllRowGroups(); } catch (error) {}
  for (let i = 0; i < 10; i++) {
    try { sheet.getRange(1, 1, sheet.getMaxRows(), 1).shiftRowGroupDepth(-1); } catch (error) {}
  }
}

function rebuildCheckGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  GROUP_TARGET_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    resetColumnGroupsForSheet_(sheet);
    createCheckGroupForSheet_(sheet);
  });
}

function resetColumnGroupsForSheet_(sheet) {
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (error) {}
  for (let i = 0; i < 10; i++) {
    try { sheet.getRange(1, 1, 1, sheet.getLastColumn()).shiftColumnGroupDepth(-1); } catch (error) {}
  }
}

function createCheckGroupForSheet_(sheet) {
  const headers = getHeaders_(sheet);
  const start = headers.indexOf(PERSONAL_MEMBERS[0]) + 1;
  const end = headers.indexOf(PERSONAL_MEMBERS[PERSONAL_MEMBERS.length - 1]) + 1;
  if (start <= 0 || end <= 0) return;

  try {
    sheet.getRange(1, start, sheet.getMaxRows(), end - start + 1).shiftColumnGroupDepth(1);
  } catch (error) {}
}

/***********************
 * 9. メニュー・トリガー・バックアップ
 ***********************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addItem("初回作成 / 列を再作成", "createCompanySheets")
    .addItem("全体更新", "refreshAll")
    .addSeparator()
    .addItem("Googleカレンダーへ反映", "syncCalendarEvents")
    .addItem("社用車予約の重複チェック", "checkCarReservationConflicts")
    .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
    .addSeparator()
    .addItem("選択行の日報文章をGPT作成", "createDailyReportText")
    .addItem("選択行の日報PDFを作成", "createSelectedDailyReportPdf")
    .addItem("当日サマリーPDFを作成", "createDailyReportPdf")
    .addSeparator()
    .addItem("スプレッドシートをバックアップ", "backupSpreadsheet")
    .addItem("毎日自動更新トリガーを設定", "installDailyRefreshTrigger")
    .addItem("自動化トリガーをまとめて設定", "installAllAutomationTriggers")
    .addToUi();
}

function installDailyRefreshTrigger() {
  deleteTriggersByFunctionName("dailyRefreshAutomation");
  ScriptApp.newTrigger("dailyRefreshAutomation").timeBased().everyDays(1).atHour(6).create();
  SpreadsheetApp.getActiveSpreadsheet().toast("毎日6時の自動更新トリガーを設定しました");
}

function installMonthlyArchiveTrigger() {
  deleteTriggersByFunctionName("movePastItemsToArchive");
  ScriptApp.newTrigger("movePastItemsToArchive").timeBased().onMonthDay(1).atHour(6).create();
  SpreadsheetApp.getActiveSpreadsheet().toast("毎月1日6時の過去一覧移動トリガーを設定しました");
}

function installAllAutomationTriggers() {
  installDailyRefreshTrigger();
  installMonthlyArchiveTrigger();
}

function dailyRefreshAutomation() {
  refreshAll();
  checkCarReservationConflicts();
}

function deleteTriggersByFunctionName(functionName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) ScriptApp.deleteTrigger(trigger);
  });
}

function backupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const backupName = "バックアップ_" + ss.getName() + "_" + timestamp;
  const parents = file.getParents();
  const backupFile = parents.hasNext() ? file.makeCopy(backupName, parents.next()) : file.makeCopy(backupName);
  ss.toast("バックアップを作成しました: " + backupName);
  return backupFile.getUrl();
}

function setOpenAIKeyForSetup() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt("OpenAI APIキー設定", "APIキーを入力してください", ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", result.getResponseText().trim());
  ui.alert("OPENAI_API_KEY を保存しました");
}

/***********************
 * 10. 共通ユーティリティ
 ***********************/

function addScheduleRows_(ss, sheetName, condition, mapper, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  getSheetObjects_(sheet).forEach(row => {
    if (condition(row)) rows.push(mapper(row));
  });
}

function writeObjectsToSheet(sheet, objects) {
  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  if (!objects || objects.length === 0) {
    setCheckboxesForDataRows(sheet);
    applyColorRules(sheet);
    return;
  }

  const values = objects.map(obj => headers.map(header => {
    if (header === READ_HEADER || PERSONAL_MEMBERS.includes(header)) return toBoolean_(obj[header]);
    return obj[header] !== undefined ? obj[header] : "";
  }));

  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  setCheckboxesForDataRows(sheet);
  applyColumnFormats_(sheet, headers);
  applyColorRules(sheet);
  safeCreateFilter(sheet, headers.length);
  hideSystemColumns_(sheet, headers);
  sheet.autoResizeColumns(1, headers.length);
}

function appendObjectsToSheet(sheet, objects) {
  if (!objects || objects.length === 0) return;
  const headers = getHeaders_(sheet);
  const values = objects.map(obj => headers.map(header => obj[header] !== undefined ? obj[header] : ""));
  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
}

function getOrCreateSheet_(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function getHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getSheetValues_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = getHeaders_(sheet);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
}

function getSheetObjects_(sheet) {
  const headers = getHeaders_(sheet);
  return getSheetValues_(sheet).map(row => objectFromRow(headers, row));
}

function objectFromRow(headers, valuesRow) {
  const obj = {};
  headers.forEach((header, index) => obj[header] = valuesRow[index]);
  return obj;
}

function copyPersonalChecks(row) {
  const result = {};
  PERSONAL_MEMBERS.forEach(member => result[member] = toBoolean_(row[member]));
  return result;
}

function countUnreadByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const headers = getHeaders_(sheet);
  const memberCol = headers.indexOf(member) + 1;
  if (memberCol <= 0) return 0;

  return getSheetValues_(sheet).filter(row => {
    const hasData = row.some(value => value !== "" && value !== null && value !== undefined);
    return hasData && row[memberCol - 1] !== true;
  }).length;
}

function countTodoByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return getSheetObjects_(sheet).filter(row => row["担当"] === member && row["内容"] && row["状態"] !== "完了").length;
}

function countPhoneByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return getSheetObjects_(sheet).filter(row => row["担当"] === member && (row["電話対応"] === "未対応" || row["電話対応"] === "折返し")).length;
}

function countTodayScheduleByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  return getSheetObjects_(sheet).filter(row => row["担当"] === member && isSameDateOnly_(row["日付"], new Date())).length;
}

function updateNoticeForEditedRow(sheet, row, col) {
  const headers = getHeaders_(sheet);
  const noticeCol = headers.indexOf("通知") + 1;
  if (noticeCol <= 0) return;

  const editedHeader = headers[col - 1];
  const dateCol = DATE_HEADERS.includes(editedHeader) ? col : findMainDateColumn(headers);
  if (dateCol <= 0) return;

  const dateValue = sheet.getRange(row, dateCol).getValue();
  sheet.getRange(row, noticeCol).setValue(getNoticeText(dateValue));
}

function findMainDateColumn(headers) {
  const priority = ["日付", "日時", "開始日", "期限", "投稿日", "登録日", "車検期限", "保険期限", "終了日", "移動日"];
  for (const name of priority) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) return col;
  }
  return 0;
}

function applyColorRules(sheet) {
  const headers = getHeaders_(sheet);
  const rules = [];

  const statusCol = headers.indexOf("状態") + 1;
  const phoneCol = headers.indexOf("電話対応") + 1;
  const readCol = headers.indexOf(READ_HEADER) + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const typeCol = headers.indexOf("種類") + 1;
  const importanceCol = headers.indexOf("重要度") + 1;

  addTextRule_(rules, sheet, statusCol, "完了", "#b6d7a8");
  addTextRule_(rules, sheet, statusCol, "進行中", "#cfe2f3");
  addTextRule_(rules, sheet, statusCol, "施工中", "#cfe2f3");
  addTextRule_(rules, sheet, statusCol, "予定", "#eeeeee");
  addTextRule_(rules, sheet, statusCol, "延期", "#ffe599");
  addTextRule_(rules, sheet, statusCol, "中止", "#ea9999");
  addTextRule_(rules, sheet, statusCol, "未対応", "#f4cccc");
  addTextRule_(rules, sheet, statusCol, "対応中", "#fce5cd");
  addTextRule_(rules, sheet, statusCol, "要確認", "#ea9999");

  addTextRule_(rules, sheet, phoneCol, "未対応", "#f4cccc");
  addTextRule_(rules, sheet, phoneCol, "折返し", "#fce5cd");
  addTextRule_(rules, sheet, phoneCol, "対応中", "#fff2cc");
  addTextRule_(rules, sheet, phoneCol, "完了", "#b6d7a8");

  addTextRule_(rules, sheet, noticeCol, "期限切れ", "#ea9999");
  addTextRule_(rules, sheet, noticeCol, "今日", "#fce5cd");
  addTextRule_(rules, sheet, noticeCol, "3日以内", "#fff2cc");
  addTextRule_(rules, sheet, noticeCol, "7日以内", "#d9ead3");
  addTextRule_(rules, sheet, noticeCol, "重要", "#ea9999");
  addTextRule_(rules, sheet, noticeCol, "予約重複", "#ea9999");

  addTextRule_(rules, sheet, typeCol, "車検", "#f4cccc");
  addTextRule_(rules, sheet, typeCol, "行事", "#fce5cd");
  addTextRule_(rules, sheet, typeCol, "会議", "#d9d2e9");
  addTextRule_(rules, sheet, typeCol, "工事予定", "#cfe2f3");
  addTextRule_(rules, sheet, typeCol, "出先予定", "#d9ead3");
  addTextRule_(rules, sheet, typeCol, "電話履歴", "#fff2cc");
  addTextRule_(rules, sheet, typeCol, "作業状況", "#eeeeee");
  addTextRule_(rules, sheet, typeCol, "ToDo", "#d9ead3");
  addTextRule_(rules, sheet, typeCol, "社用車予約", "#d9ead3");
  addTextRule_(rules, sheet, typeCol, "日報", "#d9ead3");

  addTextRule_(rules, sheet, importanceCol, "高", "#ea9999");
  addTextRule_(rules, sheet, importanceCol, "中", "#ffe599");
  addTextRule_(rules, sheet, importanceCol, "低", "#d9ead3");

  if (readCol > 0) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=$" + columnLetter(readCol) + "2=FALSE")
      .setBackground("#fff2cc")
      .setRanges([sheet.getRange(2, readCol, DATA_ROWS, 1)])
      .build());
  }

  sheet.setConditionalFormatRules(rules);
}

function addTextRule_(rules, sheet, col, text, bg) {
  if (col <= 0) return;
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(bg)
    .setRanges([sheet.getRange(2, col, DATA_ROWS, 1)])
    .build());
}

function getNoticeText(dateValue) {
  if (!dateValue) return "";

  const today = toDateOnly(new Date());
  const target = toDateOnly(dateValue);
  if (isNaN(target.getTime())) return "";

  const diff = Math.floor((target.getTime() - today.getTime()) / 86400000);
  if (diff < 0) return "期限切れ";
  if (diff === 0) return "今日";
  if (diff <= 3) return "3日以内";
  if (diff <= 7) return "7日以内";
  return "";
}

function buildDateTime(dateValue, timeValue) {
  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return null;

  const time = parseTimeValue(timeValue);
  if (!time) return null;

  date.setHours(time.hours, time.minutes, 0, 0);
  return date;
}

function parseTimeValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return {hours: value.getHours(), minutes: value.getMinutes()};

  if (typeof value === "number") {
    const totalMinutes = Math.round(value * 24 * 60);
    return {hours: Math.floor(totalMinutes / 60) % 24, minutes: totalMinutes % 60};
  }

  const text = String(value || "").trim();
  const match = text.match(/^(\d{1,2})[:：](\d{1,2})$/);
  if (!match) return null;

  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (hours < 0 || hours > 23 || minutes < 0 || minutes > 59) return null;
  return {hours: hours, minutes: minutes};
}

function toDateOnly(value) {
  const date = new Date(value);
  date.setHours(0, 0, 0, 0);
  return date;
}

function isSameDateOnly_(a, b) {
  if (!a || !b) return false;
  const da = toDateOnly(a);
  const db = toDateOnly(b);
  if (isNaN(da.getTime()) || isNaN(db.getTime())) return false;
  return da.getTime() === db.getTime();
}

function getYearMonthKey(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) return "";
  return date.getFullYear() + "-" + ("0" + (date.getMonth() + 1)).slice(-2);
}

function getYearMonthLabel(value) {
  const date = new Date(value);
  if (isNaN(date.getTime())) return "";
  return date.getFullYear() + "年" + (date.getMonth() + 1) + "月";
}

function formatDateForReport_(value) {
  if (!value) return "";
  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

function joinText(a, b) {
  if (a && b) return a + " / " + b;
  return a || b || "";
}

function toBoolean_(value) {
  return value === true || value === "TRUE" || value === "Y" || value === "はい";
}

function sanitizeFileName_(value) {
  return String(value || "").replace(/[\\/:*?"<>|]/g, "_");
}

function columnLetter(column) {
  let letter = "";
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function SHEET_ID(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  return sheet ? sheet.getSheetId() : "";
}


















