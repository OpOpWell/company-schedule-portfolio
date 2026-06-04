/**
 * 社内共有管理 v7.1 社用調整版
 *
 * - 担当者15人と既読確認者3人を分離
 * - 既読確認列はサンプル確認者3列のみ
 * - 担当プルダウンはサンプル担当者15人
 * - 工事予定の列順を「工事名、現場、開始日、終了日」へ変更
 * - Googleカレンダー時間連携を維持
 * - 社用車予約の重複チェックを追加
 * - 月ごとの行グループ化関数を追加
 * - 過去一覧シートを追加
 * - 期限切れ・過去予定を過去一覧へ移動する関数を追加
 * - GitHub公開用にサンプル名へ変更
 * - デモ用サンプルデータ投入関数を追加
 */

/**
 * 社用メモ
 * - STAFF_MEMBERS は担当プルダウン用です。
 * - PERSONAL_MEMBERS は既読確認列用です。
 * - GitHub公開用のため、担当者名はすべてサンプル名です。
 */

const STAFF_MEMBERS = [
  "田中太郎",
  "佐藤花子",
  "鈴木一郎",
  "高橋美咲",
  "伊藤健太",
  "渡辺優子",
  "山本大輔",
  "中村翔",
  "小林彩",
  "加藤誠",
  "吉田直人",
  "斎藤愛",
  "松本悠",
  "井上真由",
  "木村健"
];

const PERSONAL_MEMBERS = [
  "田中太郎",
  "佐藤花子",
  "鈴木一郎"
];

const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";
const CALENDAR_ID_HEADER = "カレンダーID";

const DATE_HEADERS = [
  "日付",
  "日時",
  "開始日",
  "終了日",
  "車検日",
  "保険期限",
  "投稿日",
  "登録日",
  "期限"
];

const SHEET_HEADERS = {
  "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, CALENDAR_ID_HEADER],
  "車検管理": ["車両名", "車番", "車検日", "通知", "保険期限", "状態", "写真", READ_HEADER, ...PERSONAL_MEMBERS],
  "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...PERSONAL_MEMBERS, CALENDAR_ID_HEADER],
  "行事予定": ["日付", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, CALENDAR_ID_HEADER],
  "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考"],
  "工事予定": ["工事名", "現場", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "備考", CALENDAR_ID_HEADER],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "メモ"],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"],
  "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考", CALENDAR_ID_HEADER],
  "担当別未読": ["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"],
  "過去一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート", "備考", "移動日"]
};

function createCompanySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEET_HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    resetSheet(sheet, SHEET_HEADERS[name].length);
    setupSheet(sheet, SHEET_HEADERS[name]);
  });


  refreshAll();

  console.log("作成完了");
}

function resetSheet(sheet, headerCount) {
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  try {
    sheet.showColumns(1, sheet.getMaxColumns());
  } catch (e) {}

  try {
    for (let i = 0; i < 10; i++) {
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).shiftColumnGroupDepth(-1);
    }
  } catch (e) {}

  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.setConditionalFormatRules([]);

  const currentCols = sheet.getMaxColumns();
  if (currentCols > headerCount) {
    sheet.deleteColumns(headerCount + 1, currentCols - headerCount);
  } else if (currentCols < headerCount) {
    sheet.insertColumnsAfter(currentCols, headerCount - currentCols);
  }
}

function setupSheet(sheet, headers) {
  const statusList = [CLEAR_LABEL, "予定", "移動中", "進行中", "着工前", "施工中", "完了", "延期", "中止", "未対応", "対応中", "要確認"];
  const staffList = [CLEAR_LABEL, ...STAFF_MEMBERS];

  const carList = [
    CLEAR_LABEL,
    "1号車",
    "2号車",
    "3号車",
    "軽トラック",
    "軽バン",
    "営業車A",
    "営業車B",
    "作業車A",
    "作業車B",
    "ダンプA",
    "ユニック車",
    "ローダー",
    "予備車A",
    "予備車B"
  ];
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
  const importanceList = [CLEAR_LABEL, "高", "中", "低"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).createFilter();

  headers.forEach((header, i) => {
    const col = i + 1;

    if (DATE_HEADERS.includes(header)) {
      setDateColumn(sheet, col);
    }

    if (header === "状態") setDropdown(sheet, col, statusList);
    if (header === "担当") setDropdown(sheet, col, staffList);
    if (header === "社用車") setDropdown(sheet, col, carList);
    if (header === "電話対応") setDropdown(sheet, col, phoneList);
    if (header === "重要度") setDropdown(sheet, col, importanceList);
  });

  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    sheet.hideColumns(calendarIdCol);
  }
}

function setDateColumn(sheet, col) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, 500)
    .setNumberFormat("yyyy/mm/dd")
    .setDataValidation(rule);
}

function setDropdown(sheet, col, list) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, 500).setDataValidation(rule);
}

function writeObjectsToSheet(sheet, objects) {
  if (!sheet || objects.length === 0) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const values = objects.map(obj => headers.map(header => {
    if (header === READ_HEADER) return obj[header] === true;
    if (PERSONAL_MEMBERS.includes(header)) return obj[header] === true;
    return obj[header] !== undefined ? obj[header] : "";
  }));

  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  setCheckboxesForDataRows(sheet);
  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    sheet.hideColumns(calendarIdCol);
  }
}

function createScheduleList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("一覧スケジュール");
  if (!sheet) sheet = ss.insertSheet("一覧スケジュール");

  const headers = SHEET_HEADERS["一覧スケジュール"];

  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  const rows = [];

  collectRowsAsObjects(ss, "出先予定", row => row["日付"], row => ({
    "日付": row["日付"],
    "種類": "出先予定",
    "内容": joinText(row["行き先"], row["用件"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": row["電話対応"],
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "出先予定"
  }), rows);

  collectRowsAsObjects(ss, "会議予定", row => row["日付"], row => ({
    "日付": row["日付"],
    "種類": "会議",
    "内容": joinText(row["会議名"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "会議予定"
  }), rows);

  collectRowsAsObjects(ss, "行事予定", row => row["日付"], row => ({
    "日付": row["日付"],
    "種類": "行事",
    "内容": joinText(row["行事名"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "行事予定"
  }), rows);

  collectRowsAsObjects(ss, "車検管理", row => row["車検日"], row => ({
    "日付": row["車検日"],
    "種類": "車検",
    "内容": joinText(row["車両名"], row["車番"]),
    "担当": "",
    "状態": row["状態"],
    "通知": getNoticeText(row["車検日"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "車検管理"
  }), rows);

  collectRowsAsObjects(ss, "作業状況", row => row["現場"] || row["作業内容"], row => ({
    "日付": new Date(),
    "種類": "作業状況",
    "内容": joinText(row["現場"], row["作業内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": row["通知"] || "",
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "作業状況"
  }), rows);

  collectRowsAsObjects(ss, "工事予定", row => row["開始日"], row => ({
    "日付": row["開始日"],
    "種類": "工事予定",
    "内容": joinText(row["工事名"], row["現場"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["開始日"]),
    "電話対応": row["電話対応"],
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "工事予定"
  }), rows);

  collectRowsAsObjects(ss, "電話履歴", row => row["日時"], row => ({
    "日付": row["日時"],
    "種類": "電話履歴",
    "内容": joinText(row["相手"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日時"]),
    "電話対応": row["電話対応"],
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "電話履歴"
  }), rows);


  collectRowsAsObjects(ss, "個人ToDo", row => row["期限"] || row["内容"], row => ({
    "日付": row["期限"] || row["登録日"] || new Date(),
    "種類": "ToDo",
    "内容": row["内容"],
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["期限"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "個人ToDo"
  }), rows);


  collectRowsAsObjects(ss, "社用車予約", row => row["日付"] || row["行き先"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "社用車予約",
    "内容": joinText(row["社用車"], joinText(row["行き先"], row["用途"])),
    "担当": row["利用者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "元シート": "社用車予約"
  }), rows);

  rows.sort((a, b) => new Date(a["日付"]) - new Date(b["日付"]));

  writeObjectsToSheet(sheet, rows);
  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    sheet.hideColumns(calendarIdCol);
  }
}

function collectRowsAsObjects(ss, sheetName, condition, mapper, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    if (condition(row)) rows.push(mapper(row));
  });
}

function objectFromRow(headers, valuesRow) {
  const obj = {};
  headers.forEach((header, i) => obj[header] = valuesRow[i]);
  return obj;
}

function copyPersonalChecks(row) {
  const result = {};

  PERSONAL_MEMBERS.forEach(name => {
    result[name] =
      row[name] === true ||
      row[name] === "TRUE" ||
      row[name] === "Y";
  });

  return result;
}


function createAlertList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("一覧スケジュール");

  let sheet = ss.getSheetByName("要確認一覧");
  if (!sheet) sheet = ss.insertSheet("要確認一覧");

  const headers = SHEET_HEADERS["要確認一覧"];

  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  if (!source || source.getLastRow() < 2) return;

  const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
  const values = source.getRange(2, 1, source.getLastRow() - 1, sourceHeaders.length).getValues();

  const priority = {
    "期限切れ": 1,
    "今日": 2,
    "3日以内": 3,
    "7日以内": 4
  };

  const rows = values
    .map(valuesRow => objectFromRow(sourceHeaders, valuesRow))
    .filter(row => priority[row["通知"]])
    .sort((a, b) => {
      const pa = priority[a["通知"]] || 99;
      const pb = priority[b["通知"]] || 99;
      if (pa !== pb) return pa - pb;
      return new Date(a["日付"]) - new Date(b["日付"]);
    });


  const noticeSheet = ss.getSheetByName("お知らせ");
  if (noticeSheet && noticeSheet.getLastRow() >= 2) {
    const noticeHeaders = noticeSheet.getRange(1, 1, 1, noticeSheet.getLastColumn()).getValues()[0];
    const noticeValues = noticeSheet.getRange(2, 1, noticeSheet.getLastRow() - 1, noticeHeaders.length).getValues();

    noticeValues.forEach(valuesRow => {
      const row = objectFromRow(noticeHeaders, valuesRow);
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

  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    sheet.hideColumns(calendarIdCol);
  }
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("ダッシュボード");
  if (!sheet) {
    sheet = ss.insertSheet("ダッシュボード");
  } else {
    try {
      sheet.showColumns(1, sheet.getMaxColumns());
    } catch (e) {}
    sheet.clear();
    sheet.clearFormats();
    sheet.setConditionalFormatRules([]);
  }

  sheet.getRange("A1").setValue("社内共有ダッシュボード");
  sheet.getRange("A1").setFontSize(22).setFontWeight("bold");

  sheet.getRange("A3:B3").setValues([["項目", "件数"]]);
  sheet.getRange("A3:B3")
    .setBackground("#d9ead3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  const data = [
    ["今日の予定", '=COUNTIF(一覧スケジュール!A2:A,TODAY())'],
    ["期限切れ", '=COUNTIF(一覧スケジュール!F2:F,"期限切れ")'],
    ["今日対応", '=COUNTIF(一覧スケジュール!F2:F,"今日")'],
    ["3日以内", '=COUNTIF(一覧スケジュール!F2:F,"3日以内")'],
    ["7日以内", '=COUNTIF(一覧スケジュール!F2:F,"7日以内")'],
    ["未読件数", '=COUNTIF(一覧スケジュール!H2:H,FALSE)'],
    ["未対応電話", '=COUNTIF(一覧スケジュール!G2:G,"未対応")'],
    ["折返し電話", '=COUNTIF(一覧スケジュール!G2:G,"折返し")'],
    ["完了件数", '=COUNTIF(一覧スケジュール!E2:E,"完了")'],
    ["高重要度お知らせ", '=COUNTIF(お知らせ!E2:E,"高")'],
    ["未完了ToDo", '=COUNTIFS(個人ToDo!C2:C,"<>",個人ToDo!E2:E,"<>完了")'],
    ["今日期限ToDo", '=COUNTIF(個人ToDo!F2:F,"今日")'],
    ["社用車予約", '=COUNTA(社用車予約!A2:A)'],
    ["担当別未読あり", '=COUNTIF(担当別未読!B2:B,">0")']
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);
  sheet.getRange(3, 1, data.length + 1, 2).setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 4);
}

function refreshAll() {
  refreshInputSheets();
  createScheduleList();
  createAlertList();
  createAssigneeUnreadSummary();
  createReadRateSummary();
  createDashboard();
  resetColumnGroups();
  createCheckGroup();
}

function setCheckboxesForDataRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1 || lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const checkNames = [READ_HEADER].concat(PERSONAL_MEMBERS);

  checkNames.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;

    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const values = range.getValues().map(row => {
      const v = row[0];
      if (v === true || v === "TRUE") return [true];
      if (v === false || v === "FALSE" || v === "") return [false];
      return [false];
    });

    range.setValues(values);
    range.insertCheckboxes();
  });
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row === 1) return;

  if (e.range.getValue() === CLEAR_LABEL) {
    e.range.clearContent();
    return;
  }

  if (sheet.getName() === "社用車予約") {
    const conflictMessage = getCarReservationConflictMessage(sheet, row);
    if (conflictMessage) {
      SpreadsheetApp.getUi().alert(conflictMessage);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const noticeCol = headers.indexOf("通知") + 1;
      if (noticeCol > 0) sheet.getRange(row, noticeCol).setValue("予約重複");
    }
  }

  updateNoticeForEditedRow(sheet, row, col);
updateCheckboxesForEditedRow(sheet, row);

refreshAll();
}

function updateNoticeForEditedRow(sheet, row, col) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const noticeCol = headers.indexOf("通知") + 1;
  if (noticeCol <= 0) return;

  const editedHeader = headers[col - 1];

  let dateCol = 0;

  if (DATE_HEADERS.includes(editedHeader)) {
    dateCol = col;
  } else {
    dateCol = findMainDateColumn(headers);
  }

  if (dateCol <= 0) return;

  const dateValue = sheet.getRange(row, dateCol).getValue();
  sheet.getRange(row, noticeCol).setValue(getNoticeText(dateValue));
}

function findMainDateColumn(headers) {
  const priority = ["日付", "日時", "開始日", "期限", "投稿日", "登録日", "車検日", "保険期限", "終了日"];

  for (const name of priority) {
    const col = headers.indexOf(name) + 1;
    if (col > 0) return col;
  }

  return 0;
}

function updateCheckboxesForEditedRow(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol <= 0) return;

  const rowData = sheet.getRange(row, 1, 1, readCol - 1).getValues()[0];
  const hasData = rowData.some(v => v !== "");

  const checkNames = [READ_HEADER].concat(PERSONAL_MEMBERS);

  checkNames.forEach(name => {
    const checkCol = headers.indexOf(name) + 1;
    if (checkCol <= 0) return;

    const cell = sheet.getRange(row, checkCol);

    if (hasData) {
      cell.insertCheckboxes();
    } else {
      cell.clearDataValidations();
      cell.clearContent();
    }
  });
}

function getTargetSheets() {
  return [
    "出先予定",
    "車検管理",
    "会議予定",
    "行事予定",
    "作業状況",
    "工事予定",
    "電話履歴",
    "一覧スケジュール",
    "お知らせ",
    "個人ToDo",
    "社用車予約",
    "過去一覧"
  ];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addItem("全体更新", "refreshAll")
    .addItem("デモ用サンプルデータ投入", "addSampleDataForGitHubDemo")
    .addItem("要確認一覧を更新", "createAlertList")
    .addItem("担当別未読を更新", "createAssigneeUnreadSummary")
    .addItem("既読率集計を更新", "createReadRateSummary")
    .addItem("Googleカレンダーへ反映", "syncCalendarEvents")
    .addItem("社用車予約の重複チェック", "checkCarReservationConflicts")
    .addItem("月ごとに折りたたみ", "groupRowsByMonth")
    .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
    .addSeparator()
    .addItem("PDF日報を作成", "createDailyReportPdf")
    .addItem("スプレッドシートをバックアップ", "backupSpreadsheet")
    .addItem("毎日自動更新トリガーを設定", "installDailyRefreshTrigger")
    .addItem("自動化トリガーをまとめて設定", "installAllAutomationTriggers")
    .addSeparator()
    .addItem("個人確認グループを作り直す", "rebuildCheckGroups")
    .addToUi();
}

function rebuildCheckGroups() {
  resetColumnGroups();
  createCheckGroup();
}

function resetColumnGroups() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getTargetSheets().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    try {
      sheet.showColumns(1, sheet.getMaxColumns());
    } catch (e) {}

    for (let i = 0; i < 10; i++) {
      try {
        sheet.getRange(1, 1, 1, sheet.getLastColumn()).shiftColumnGroupDepth(-1);
      } catch (e) {}
    }
  });
}

function createCheckGroup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getTargetSheets().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const start = headers.indexOf(PERSONAL_MEMBERS[0]) + 1;
    const end = headers.indexOf(PERSONAL_MEMBERS[PERSONAL_MEMBERS.length - 1]) + 1;

    if (start <= 0 || end <= 0) return;

    try {
    
      sheet.getRange(1, start, sheet.getMaxRows(), end - start + 1)
        .shiftColumnGroupDepth(1);
    } catch (e) {}
  });
}

function applyColorRules(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = headers.indexOf("状態") + 1;
  const phoneCol = headers.indexOf("電話対応") + 1;
  const readCol = headers.indexOf(READ_HEADER) + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const typeCol = headers.indexOf("種類") + 1;
  const importanceCol = headers.indexOf("重要度") + 1;

  const rules = [];

  function addTextRule(col, text, bg) {
    if (col <= 0) return;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(bg)
        .setRanges([sheet.getRange(2, col, 500)])
        .build()
    );
  }

  addTextRule(statusCol, "完了", "#b6d7a8");
  addTextRule(statusCol, "進行中", "#cfe2f3");
  addTextRule(statusCol, "施工中", "#cfe2f3");
  addTextRule(statusCol, "予定", "#eeeeee");
  addTextRule(statusCol, "延期", "#ffe599");
  addTextRule(statusCol, "中止", "#ea9999");
  addTextRule(statusCol, "未対応", "#f4cccc");
  addTextRule(statusCol, "対応中", "#fce5cd");

  addTextRule(phoneCol, "未対応", "#f4cccc");
  addTextRule(phoneCol, "折返し", "#fce5cd");
  addTextRule(phoneCol, "対応中", "#fff2cc");
  addTextRule(phoneCol, "完了", "#b6d7a8");

  if (readCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied("=$" + columnLetter(readCol) + "2=FALSE")
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(2, readCol, 500)])
        .build()
    );
  }

  if (noticeCol > 0) {
    addTextRule(noticeCol, "期限切れ", "#ea9999");
    addTextRule(noticeCol, "今日", "#fce5cd");
    addTextRule(noticeCol, "3日以内", "#fff2cc");
    addTextRule(noticeCol, "7日以内", "#d9ead3");
    addTextRule(noticeCol, "重要", "#ea9999");
    addTextRule(noticeCol, "予約重複", "#ea9999");
  }

  addTextRule(typeCol, "車検", "#f4cccc");
  addTextRule(typeCol, "行事", "#fce5cd");
  addTextRule(typeCol, "会議", "#d9d2e9");
  addTextRule(typeCol, "工事予定", "#cfe2f3");
  addTextRule(typeCol, "出先予定", "#d9ead3");
  addTextRule(typeCol, "電話履歴", "#fff2cc");
  addTextRule(typeCol, "作業状況", "#eeeeee");
  addTextRule(typeCol, "ToDo", "#d9ead3");
  addTextRule(typeCol, "社用車予約", "#d9ead3");

  addTextRule(importanceCol, "高", "#ea9999");
  addTextRule(importanceCol, "中", "#ffe599");
  addTextRule(importanceCol, "低", "#d9ead3");

  sheet.setConditionalFormatRules(rules);
}



function createAssigneeUnreadSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("担当別未読");
  if (!sheet) sheet = ss.insertSheet("担当別未読");

  const headers = SHEET_HEADERS["担当別未読"];
  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  const schedule = ss.getSheetByName("一覧スケジュール");
  const todo = ss.getSheetByName("個人ToDo");
  const phone = ss.getSheetByName("電話履歴");

  const rows = PERSONAL_MEMBERS.map(checkName => {
    const staffName = checkName.replace("確認", "");

    return {
      "担当": staffName,
      "未読件数": countUnreadByMember(schedule, checkName),
      "未完了ToDo": countTodoByMember(todo, staffName),
      "未対応電話": countPhoneByMember(phone, staffName),
      "今日の予定": countTodayScheduleByMember(schedule, staffName)
    };
  });

  writeObjectsToSheet(sheet, rows);
  sheet.autoResizeColumns(1, headers.length);
}

function countUnreadByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const memberCol = headers.indexOf(member) + 1;
  if (memberCol <= 0) return 0;

  const values = sheet.getRange(2, memberCol, sheet.getLastRow() - 1, 1).getValues();
  return values.filter(row => row[0] !== true).length;
}

function countTodoByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const staffCol = headers.indexOf("担当") + 1;
  const statusCol = headers.indexOf("状態") + 1;
  const contentCol = headers.indexOf("内容") + 1;
  if (staffCol <= 0 || statusCol <= 0 || contentCol <= 0) return 0;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.filter(row => {
    const content = row[contentCol - 1];
    const staff = row[staffCol - 1];
    const status = row[statusCol - 1];
    return content && staff === member && status !== "完了";
  }).length;
}

function countPhoneByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const staffCol = headers.indexOf("担当") + 1;
  const phoneCol = headers.indexOf("電話対応") + 1;
  if (staffCol <= 0 || phoneCol <= 0) return 0;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.filter(row => {
    const staff = row[staffCol - 1];
    const phone = row[phoneCol - 1];
    return staff === member && (phone === "未対応" || phone === "折返し");
  }).length;
}

function countTodayScheduleByMember(sheet, member) {
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateCol = headers.indexOf("日付") + 1;
  const staffCol = headers.indexOf("担当") + 1;
  if (dateCol <= 0 || staffCol <= 0) return 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.filter(row => {
    const date = new Date(row[dateCol - 1]);
    date.setHours(0, 0, 0, 0);
    return row[staffCol - 1] === member && date.getTime() === today.getTime();
  }).length;
}

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
      endTime: row => row["終了時刻"] || row["終了時間"] || "",
      description: row => buildCalendarDescription(row, "出先予定")
    },
    {
      sheetName: "会議予定",
      title: row => "会議: " + joinText(row["会議名"], row["内容"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || "",
      description: row => buildCalendarDescription(row, "会議予定")
    },
    {
      sheetName: "行事予定",
      title: row => "行事: " + joinText(row["行事名"], row["内容"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || "",
      description: row => buildCalendarDescription(row, "行事予定")
    },
    {
      sheetName: "工事予定",
      title: row => "工事: " + joinText(row["工事名"], row["現場"]),
      startDate: row => row["開始日"],
      endDate: row => row["終了日"] || row["開始日"],
      startTime: row => row["開始時刻"] || row["開始時間"] || "",
      endTime: row => row["終了時刻"] || row["終了時間"] || "",
      description: row => buildCalendarDescription(row, "工事予定")
    },
    {
      sheetName: "社用車予約",
      title: row => "社用車予約: " + joinText(row["社用車"], row["行き先"]),
      startDate: row => row["日付"],
      endDate: row => row["日付"],
      startTime: row => row["開始時刻"],
      endTime: row => row["終了時刻"],
      description: row => buildCalendarDescription(row, "社用車予約")
    }
  ];

  targets.forEach(target => {
    const sheet = ss.getSheetByName(target.sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
    if (calendarIdCol <= 0) return;

    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

    values.forEach((valuesRow, index) => {
      const row = objectFromRow(headers, valuesRow);
      const startDateValue = target.startDate(row);
      if (!startDateValue) return;

      const title = target.title(row);
      if (!title || title.endsWith(": ")) return;

      const oldEventId = row[CALENDAR_ID_HEADER];
      if (oldEventId) {
        try {
          const oldEvent = calendar.getEventById(oldEventId);
          if (oldEvent) oldEvent.deleteEvent();
        } catch (e) {}
      }

      const description = target.description(row);
      const startTimeValue = target.startTime(row);
      const endTimeValue = target.endTime(row);

      let event;

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
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("Googleカレンダーへ反映しました");
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

function buildDateTime(dateValue, timeValue) {
  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return null;

  const time = parseTimeValue(timeValue);
  if (!time) return null;

  date.setHours(time.hours, time.minutes, 0, 0);
  return date;
}

function parseTimeValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      hours: value.getHours(),
      minutes: value.getMinutes()
    };
  }

  if (typeof value === "number") {
    const totalMinutes = Math.round(value * 24 * 60);
    return {
      hours: Math.floor(totalMinutes / 60) % 24,
      minutes: totalMinutes % 60
    };
  }

  const text = String(value || "").trim();
  if (!text) return null;

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

function getNoticeText(dateValue) {
  if (!dateValue) return "";

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const target = new Date(dateValue);
  target.setHours(0, 0, 0, 0);

  if (isNaN(target.getTime())) return "";

  const diff = Math.floor((target - today) / 86400000);

  if (diff < 0) return "期限切れ";
  if (diff === 0) return "今日";
  if (diff <= 3) return "3日以内";
  if (diff <= 7) return "7日以内";

  return "";
}


function joinText(a, b) {
  if (a && b) return a + " / " + b;
  return a || b || "";
}

function columnLetter(column) {
  let temp = "";
  let letter = "";

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}


function checkCarReservationConflicts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("社用車予約");
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const noticeCol = headers.indexOf("通知") + 1;

  let conflictCount = 0;

  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const message = getCarReservationConflictMessage(sheet, row);
    if (noticeCol > 0) {
      if (message) {
        sheet.getRange(row, noticeCol).setValue("予約重複");
        conflictCount++;
      } else {
        const current = sheet.getRange(row, noticeCol).getValue();
        if (current === "予約重複") {
          const dateCol = headers.indexOf("日付") + 1;
          const dateValue = dateCol > 0 ? sheet.getRange(row, dateCol).getValue() : "";
          sheet.getRange(row, noticeCol).setValue(getNoticeText(dateValue));
        }
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("社用車予約の重複チェック完了: " + conflictCount + "件");
}

function getCarReservationConflictMessage(sheet, targetRow) {
  if (!sheet || sheet.getName() !== "社用車予約") return "";

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return "";

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const dateCol = headers.indexOf("日付") + 1;
  const startCol = headers.indexOf("開始時刻") + 1;
  const endCol = headers.indexOf("終了時刻") + 1;
  const carCol = headers.indexOf("社用車") + 1;
  const statusCol = headers.indexOf("状態") + 1;

  if (dateCol <= 0 || startCol <= 0 || endCol <= 0 || carCol <= 0) return "";

  const targetDate = sheet.getRange(targetRow, dateCol).getValue();
  const targetStartTime = sheet.getRange(targetRow, startCol).getValue();
  const targetEndTime = sheet.getRange(targetRow, endCol).getValue();
  const targetCar = sheet.getRange(targetRow, carCol).getValue();
  const targetStatus = statusCol > 0 ? sheet.getRange(targetRow, statusCol).getValue() : "";

  if (!targetDate || !targetStartTime || !targetEndTime || !targetCar) return "";
  if (targetStatus === "中止") return "";

  const targetStart = buildDateTime(targetDate, targetStartTime);
  const targetEnd = buildDateTime(targetDate, targetEndTime);
  if (!targetStart || !targetEnd || targetEnd.getTime() <= targetStart.getTime()) return "";

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    if (rowNumber === targetRow) continue;

    const row = objectFromRow(headers, values[i]);
    if (row["状態"] === "中止") continue;
    if (row["社用車"] !== targetCar) continue;

    const otherStart = buildDateTime(row["日付"], row["開始時刻"]);
    const otherEnd = buildDateTime(row["日付"], row["終了時刻"]);
    if (!otherStart || !otherEnd) continue;

    const overlap = targetStart.getTime() < otherEnd.getTime() && otherStart.getTime() < targetEnd.getTime();
    if (overlap) {
      return "社用車予約が重複しています。\n\n車両: " + targetCar + "\n重複行: " + rowNumber + "行目";
    }
  }

  return "";
}

function groupRowsByMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getLastRow() < 3) {
    SpreadsheetApp.getActiveSpreadsheet().toast("月ごとに折りたたむ対象データがありません");
    return;
  }

  resetRowGroups(sheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateCol = findMainDateColumn(headers);

  if (dateCol <= 0) {
    SpreadsheetApp.getUi().alert("このシートには日付列が見つかりません");
    return;
  }

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();

  let currentKey = "";
  let groupStart = 2;

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const dateValue = values[i][0];
    const key = getYearMonthKey(dateValue);

    if (!key) continue;

    if (!currentKey) {
      currentKey = key;
      groupStart = rowNumber;
      continue;
    }

    if (key !== currentKey) {
      const groupEnd = rowNumber - 1;
      if (groupEnd > groupStart) {
        sheet.getRange(groupStart, 1, groupEnd - groupStart + 1, 1).shiftRowGroupDepth(1);
      }
      currentKey = key;
      groupStart = rowNumber;
    }
  }

  if (currentKey && lastRow > groupStart) {
    sheet.getRange(groupStart, 1, lastRow - groupStart + 1, 1).shiftRowGroupDepth(1);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("月ごとの折りたたみを作成しました");
}

function resetRowGroups(sheet) {
  if (!sheet) return;

  try {
    sheet.expandAllRowGroups();
  } catch (e) {}

  for (let i = 0; i < 10; i++) {
    try {
      sheet.getRange(1, 1, sheet.getMaxRows(), 1).shiftRowGroupDepth(-1);
    } catch (e) {}
  }
}

function getYearMonthKey(value) {
  if (!value) return "";

  const date = new Date(value);
  if (isNaN(date.getTime())) return "";

  return date.getFullYear() + "-" + ("0" + (date.getMonth() + 1)).slice(-2);
}



function movePastItemsToArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let archiveSheet = ss.getSheetByName("過去一覧");
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet("過去一覧");
  }

  resetSheet(archiveSheet, SHEET_HEADERS["過去一覧"].length);
  setupSheet(archiveSheet, SHEET_HEADERS["過去一覧"]);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const targetSheetNames = [
    "出先予定",
    "会議予定",
    "行事予定",
    "工事予定",
    "電話履歴",
    "社用車予約",
    "個人ToDo"
  ];

  const archiveRows = [];

  targetSheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dateCol = findMainDateColumn(headers);
    if (dateCol <= 0) return;

    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    const rowsToDelete = [];

    values.forEach((valuesRow, index) => {
      const rowNumber = index + 2;
      const row = objectFromRow(headers, valuesRow);
      const dateValue = row[headers[dateCol - 1]];

      if (!dateValue) return;

      const date = new Date(dateValue);
      if (isNaN(date.getTime())) return;
      date.setHours(0, 0, 0, 0);

      if (date.getTime() < today.getTime()) {
        archiveRows.push(buildArchiveRow(row, sheetName));
        rowsToDelete.push(rowNumber);
      }
    });

    rowsToDelete.reverse().forEach(rowNumber => {
      sheet.deleteRow(rowNumber);
    });
  });

  if (archiveRows.length > 0) {
    appendObjectsToSheet(archiveSheet, archiveRows);
  }

  refreshAll();
  groupArchiveByMonthWithHeader();

  SpreadsheetApp.getActiveSpreadsheet().toast("過去一覧へ移動しました: " + archiveRows.length + "件");
}

function buildArchiveRow(row, sheetName) {
  const today = new Date();

  let dateValue = "";
  let content = "";
  let staff = "";
  let status = row["状態"] || "";
  let notice = row["通知"] || "";
  let phone = row["電話対応"] || "";
  let note = row["備考"] || row["メモ"] || "";

  if (sheetName === "出先予定") {
    dateValue = row["日付"];
    content = joinText(row["行き先"], row["用件"]);
    staff = row["担当"];
  } else if (sheetName === "会議予定") {
    dateValue = row["日付"];
    content = joinText(row["会議名"], row["内容"]);
    staff = row["担当"];
  } else if (sheetName === "行事予定") {
    dateValue = row["日付"];
    content = joinText(row["行事名"], row["内容"]);
    staff = row["担当"];
  } else if (sheetName === "工事予定") {
    dateValue = row["開始日"];
    content = joinText(row["工事名"], row["現場"]);
    staff = row["担当"];
  } else if (sheetName === "電話履歴") {
    dateValue = row["日時"];
    content = joinText(row["相手"], row["内容"]);
    staff = row["担当"];
  } else if (sheetName === "社用車予約") {
    dateValue = row["日付"];
    content = joinText(row["社用車"], joinText(row["行き先"], row["用途"]));
    staff = row["利用者"];
  } else if (sheetName === "個人ToDo") {
    dateValue = row["期限"] || row["登録日"];
    content = row["内容"];
    staff = row["担当"];
  } else {
    dateValue = row["日付"] || row["日時"] || row["開始日"] || row["期限"] || "";
    content = row["内容"] || "";
    staff = row["担当"] || "";
  }

  return {
    "日付": dateValue,
    "種類": sheetName,
    "内容": content,
    "担当": staff,
    "状態": status,
    "通知": notice,
    "電話対応": phone,
    "元シート": sheetName,
    "備考": note,
    "移動日": today
  };
}

function appendObjectsToSheet(sheet, objects) {
  if (!sheet || objects.length === 0) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const values = objects.map(obj => headers.map(header => {
    return obj[header] !== undefined ? obj[header] : "";
  }));

  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);
}

function installMonthlyArchiveTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "movePastItemsToArchive") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("movePastItemsToArchive")
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("毎月1日 6時の過去一覧移動トリガーを設定しました");
}


function SHEET_ID(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "";
  return sheet.getSheetId();
}



/**
 * GitHubデモ用サンプルデータ投入
 *
 * 使い方:
 * 1. createCompanySheets() を実行してシートを作成
 * 2. addSampleDataForGitHubDemo() を実行
 * 3. refreshAll() で一覧・集計を更新
 *
 * 注意:
 * - 既存データを消さず、末尾にサンプル行を追加します。
 * - 社内実運用版では実行不要です。
 */
function addSampleDataForGitHubDemo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();

  const d = days => {
    const date = new Date(today);
    date.setDate(date.getDate() + days);
    date.setHours(0, 0, 0, 0);
    return date;
  };

  appendObjectsToSheet(ss.getSheetByName("出先予定"), [
    {"日付": d(0), "行き先": "市役所", "用件": "申請書類提出", "担当": "田中太郎", "社用車": "1号車", "状態": "予定", "通知": getNoticeText(d(0)), "電話対応": "", "既読": false, "田中太郎": true, "佐藤花子": false, "鈴木一郎": false},
    {"日付": d(2), "行き先": "取引先A", "用件": "工程打合せ", "担当": "佐藤花子", "社用車": "営業車A", "状態": "予定", "通知": getNoticeText(d(2)), "電話対応": "", "既読": false, "田中太郎": false, "佐藤花子": true, "鈴木一郎": false}
  ]);

  appendObjectsToSheet(ss.getSheetByName("会議予定"), [
    {"日付": d(1), "会議名": "安全衛生会議", "内容": "今週の作業確認", "担当": "鈴木一郎", "状態": "予定", "通知": getNoticeText(d(1)), "資料": "", "既読": false, "田中太郎": true, "佐藤花子": true, "鈴木一郎": false}
  ]);

  appendObjectsToSheet(ss.getSheetByName("行事予定"), [
    {"日付": d(5), "行事名": "社内研修", "内容": "AppSheet操作説明", "担当": "高橋美咲", "状態": "予定", "通知": getNoticeText(d(5)), "既読": false, "田中太郎": false, "佐藤花子": false, "鈴木一郎": false}
  ]);

  appendObjectsToSheet(ss.getSheetByName("車検管理"), [
    {"車両名": "1号車", "車番": "秋田500 あ 1234", "車検日": d(7), "通知": getNoticeText(d(7)), "保険期限": d(30), "状態": "予定", "写真": "", "既読": false, "田中太郎": true, "佐藤花子": false, "鈴木一郎": false}
  ]);

  appendObjectsToSheet(ss.getSheetByName("作業状況"), [
    {"現場": "サンプル現場A", "作業内容": "側溝据付", "状態": "施工中", "担当": "伊藤健太", "写真": "", "通知": "", "既読": false, "田中太郎": false, "佐藤花子": true, "鈴木一郎": false, "備考": "午後から写真整理"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("工事予定"), [
    {"工事名": "道路改良工事", "現場": "サンプル地区", "開始日": d(3), "終了日": d(20), "状態": "着工前", "担当": "山本大輔", "通知": getNoticeText(d(3)), "電話対応": "", "既読": false, "田中太郎": false, "佐藤花子": false, "鈴木一郎": true, "備考": "着工前打合せ予定"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("電話履歴"), [
    {"日時": d(0), "相手": "取引先B", "内容": "見積書の確認依頼", "電話対応": "折返し", "担当": "渡辺優子", "状態": "未対応", "通知": getNoticeText(d(0)), "既読": false, "田中太郎": false, "佐藤花子": false, "鈴木一郎": false, "メモ": "午前中に折返し"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("お知らせ"), [
    {"投稿日": d(0), "タイトル": "月例点検のお知らせ", "内容": "社用車の点検予定を確認してください", "投稿者": "小林彩", "重要度": "高", "通知": "重要", "既読": false, "田中太郎": false, "佐藤花子": false, "鈴木一郎": false}
  ]);

  appendObjectsToSheet(ss.getSheetByName("個人ToDo"), [
    {"登録日": d(0), "担当": "田中太郎", "内容": "工程表を更新", "期限": d(2), "状態": "未対応", "通知": getNoticeText(d(2)), "既読": false, "田中太郎": false, "佐藤花子": true, "鈴木一郎": false}
  ]);

  appendObjectsToSheet(ss.getSheetByName("社用車予約"), [
    {"日付": d(1), "開始時刻": "09:00", "終了時刻": "11:00", "社用車": "1号車", "利用者": "佐藤花子", "行き先": "サンプル現場A", "用途": "現場確認", "状態": "予定", "通知": getNoticeText(d(1)), "既読": false, "田中太郎": false, "佐藤花子": true, "鈴木一郎": false, "備考": "午前使用"}
  ]);

  getTargetSheets().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) setCheckboxesForDataRows(sheet);
  });

  refreshAll();
  SpreadsheetApp.getActiveSpreadsheet().toast("GitHubデモ用サンプルデータを追加しました");
}


// GitHub公開用のサンプルデータ投入機能を追加しています。
// 社内実運用版では addSampleDataForGitHubDemo() は実行不要です。

function groupArchiveByMonthWithHeader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("過去一覧");

  if (!sheet || sheet.getLastRow() < 2) return;

  removeArchiveMonthHeaderRows(sheet);
  resetRowGroups(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  // 空行を除いて、過去データだけを取り出す。
  const allValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const dataRows = allValues.filter(row => row[0] && row[1] !== "月見出し");

  // 既存データ範囲を一度クリアして、空行を詰める。
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (dataRows.length === 0) {
    setArchiveDropdownsAndFormat(sheet);
    return;
  }

  dataRows.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  let outputRows = [];
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

  // 月見出し行を装飾し、見出し下のデータだけ折りたたむ。
  const finalLastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, finalLastRow - 1, lastCol).getValues();

  let groupStart = 0;

  values.forEach((row, index) => {
    const rowNumber = index + 2;
    const isMonthHeader = row[1] === "月見出し";

    if (isMonthHeader) {
      if (groupStart > 0 && rowNumber > groupStart) {
        sheet.getRange(groupStart, 1, rowNumber - groupStart, 1).shiftRowGroupDepth(1);
      }

      const headerRange = sheet.getRange(rowNumber, 1, 1, lastCol);
      headerRange.clearDataValidations();
      headerRange.setBackground("#d9ead3");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("left");

      groupStart = rowNumber + 1;
    }
  });

  if (groupStart > 0 && finalLastRow >= groupStart) {
    sheet.getRange(groupStart, 1, finalLastRow - groupStart + 1, 1).shiftRowGroupDepth(1);
  }

  setArchiveDropdownsAndFormat(sheet);
  sheet.autoResizeColumns(1, lastCol);
}

function groupArchiveByMonth() {
  groupArchiveByMonthWithHeader();
}

function removeArchiveMonthHeaderRows(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowsToDelete = [];

  values.forEach((row, index) => {
    if (row[1] === "月見出し") {
      rowsToDelete.push(index + 2);
    }
  });

  rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
}

function setArchiveDropdownsAndFormat(sheet) {
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("状態") + 1;
  const phoneCol = headers.indexOf("電話対応") + 1;

  if (statusCol > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, "予定", "移動中", "進行中", "着工前", "施工中", "完了", "延期", "中止", "未対応", "対応中", "要確認"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, statusCol, sheet.getMaxRows() - 1).setDataValidation(rule);
  }

  if (phoneCol > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, phoneCol, sheet.getMaxRows() - 1).setDataValidation(rule);
  }

  applyColorRules(sheet);
}

function getYearMonthLabel(value) {
  if (!value) return "";

  const date = new Date(value);
  if (isNaN(date.getTime())) return "";

  return date.getFullYear() + "年" + (date.getMonth() + 1) + "月";
}

/****************************************************************
 * 追加機能 v7.1
 * - 毎日自動更新トリガー
 * - スプレッドシート自動バックアップ
 * - PDF日報出力
 * - 既読率集計
 ****************************************************************/

function installDailyRefreshTrigger() {
  deleteTriggersByFunctionName("dailyRefreshAutomation");

  ScriptApp.newTrigger("dailyRefreshAutomation")
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("毎日6時の自動更新トリガーを設定しました");
}

function installAllAutomationTriggers() {
  installDailyRefreshTrigger();
  installMonthlyArchiveTrigger();
  SpreadsheetApp.getActiveSpreadsheet().toast("毎日更新・毎月過去移動のトリガーを設定しました");
}

function dailyRefreshAutomation() {
  refreshAll();
  checkCarReservationConflicts();
  groupArchiveByMonthWithHeader();
}

function deleteTriggersByFunctionName(functionName) {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function backupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const file = DriveApp.getFileById(ss.getId());
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const backupName = "バックアップ_" + ss.getName() + "_" + timestamp;

  let backupFile;

  const parents = file.getParents();
  if (parents.hasNext()) {
    const folder = parents.next();
    backupFile = file.makeCopy(backupName, folder);
  } else {
    backupFile = file.makeCopy(backupName);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("バックアップを作成しました: " + backupName);
  return backupFile.getUrl();
}

function createReadRateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("既読率集計");

  if (!sheet) {
    sheet = ss.insertSheet("既読率集計");
  } else {
    try {
      sheet.showColumns(1, sheet.getMaxColumns());
    } catch (e) {}
    sheet.clear();
    sheet.clearFormats();
    sheet.setConditionalFormatRules([]);
  }

  const headers = ["担当", "対象件数", "既読件数", "未読件数", "既読率"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  const targetSheets = [
    "出先予定",
    "車検管理",
    "会議予定",
    "行事予定",
    "作業状況",
    "工事予定",
    "電話履歴",
    "一覧スケジュール",
    "お知らせ",
    "個人ToDo",
    "社用車予約",
    "過去一覧"
  ];

  const rows = PERSONAL_MEMBERS.map(checkName => {
    const staffName = checkName.replace("確認", "");
    let total = 0;
    let read = 0;

    targetSheets.forEach(sheetName => {
      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet || targetSheet.getLastRow() < 2) return;

      const sheetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      const checkCol = sheetHeaders.indexOf(checkName) + 1;
      if (checkCol <= 0) return;

      const values = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).getValues();

      values.forEach(row => {
        if (row[1] === "月見出し") return;

        const hasData = row.some(value => value !== "" && value !== null && value !== undefined);
        if (!hasData) return;

        total++;
        if (row[checkCol - 1] === true) read++;
      });
    });

    const unread = total - read;
    const rate = total === 0 ? 0 : read / total;

    return [staffName, total, read, unread, rate];
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat("0.0%");
  }

  sheet.getRange(1, 1, rows.length + 1, headers.length).setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, headers.length);
}

function createDailyReportPdf() {
  refreshAll();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const dateText = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy年M月d日");
  const fileDateText = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
  const doc = DocumentApp.create("日報_" + fileDateText);
  const body = doc.getBody();

  body.appendParagraph("社内共有 日報").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(dateText);
  body.appendParagraph("");

  appendDailyReportSection(body, "本日の予定", getTodayScheduleRows_());
  appendDailyReportSection(body, "要確認", getAlertRows_());
  appendDailyReportSection(body, "未対応・折返し電話", getPhoneAlertRows_());
  appendDailyReportSection(body, "担当別未読", getAssigneeUnreadRows_());

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getBlob().getAs(MimeType.PDF).setName("日報_" + fileDateText + ".pdf");

  let pdfFile;
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parents = spreadsheetFile.getParents();

  if (parents.hasNext()) {
    pdfFile = parents.next().createFile(pdfBlob);
  } else {
    pdfFile = DriveApp.createFile(pdfBlob);
  }

  docFile.setTrashed(true);

  SpreadsheetApp.getActiveSpreadsheet().toast("PDF日報を作成しました");
  return pdfFile.getUrl();
}

function appendDailyReportSection(body, title, rows) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (!rows || rows.length === 0) {
    body.appendParagraph("該当なし");
    body.appendParagraph("");
    return;
  }

  const table = body.appendTable(rows);
  table.setBorderWidth(1);
  body.appendParagraph("");
}

function getTodayScheduleRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("一覧スケジュール");
  const result = [["日付", "種類", "内容", "担当", "状態", "通知"]];

  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    if (isSameDateOnly_(row["日付"], new Date())) {
      result.push([
        formatDateForReport_(row["日付"]),
        row["種類"] || "",
        row["内容"] || "",
        row["担当"] || "",
        row["状態"] || "",
        row["通知"] || ""
      ]);
    }
  });

  return result.length === 1 ? [] : result;
}

function getAlertRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("要確認一覧");
  const result = [["日付", "種類", "内容", "担当", "状態", "通知"]];

  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    result.push([
      formatDateForReport_(row["日付"]),
      row["種類"] || "",
      row["内容"] || "",
      row["担当"] || "",
      row["状態"] || "",
      row["通知"] || ""
    ]);
  });

  return result.length === 1 ? [] : result;
}

function getPhoneAlertRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("電話履歴");
  const result = [["日時", "相手", "内容", "担当", "電話対応", "状態"]];

  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    if (row["電話対応"] === "未対応" || row["電話対応"] === "折返し") {
      result.push([
        formatDateForReport_(row["日時"]),
        row["相手"] || "",
        row["内容"] || "",
        row["担当"] || "",
        row["電話対応"] || "",
        row["状態"] || ""
      ]);
    }
  });

  return result.length === 1 ? [] : result;
}

function getAssigneeUnreadRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("担当別未読");
  const result = [["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"]];

  if (!sheet || sheet.getLastRow() < 2) return result;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

  values.forEach(row => {
    result.push([
      row[0] || "",
      String(row[1] || 0),
      String(row[2] || 0),
      String(row[3] || 0),
      String(row[4] || 0)
    ]);
  });

  return result.length === 1 ? [] : result;
}

function isSameDateOnly_(a, b) {
  if (!a || !b) return false;

  const da = new Date(a);
  const db = new Date(b);

  if (isNaN(da.getTime()) || isNaN(db.getTime())) return false;

  da.setHours(0, 0, 0, 0);
  db.setHours(0, 0, 0, 0);

  return da.getTime() === db.getTime();
}

function formatDateForReport_(value) {
  if (!value) return "";

  const date = new Date(value);
  if (isNaN(date.getTime())) return String(value);

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}



function refreshInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheets = [
    "出先予定",
    "工事予定",
    "電話履歴",
    "会議予定",
    "行事予定",
    "車検管理",
    "社用車予約",
    "個人ToDo",
    "お知らせ"
  ];

  targetSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const noticeCol = headers.indexOf("通知") + 1;
    const dateCol = findMainDateColumn(headers);

    if (noticeCol > 0 && dateCol > 0) {
      for (let row = 2; row <= sheet.getLastRow(); row++) {
        const dateValue = sheet.getRange(row, dateCol).getValue();
        sheet.getRange(row, noticeCol).setValue(getNoticeText(dateValue));
      }
    }

    setCheckboxesForDataRows(sheet);
    applyColorRules(sheet);
  });
}


function copyPersonalChecks(row) {
  const result = {};
  PERSONAL_MEMBERS.forEach(name => {
    result[name] = row[name] === true || row[name] === "TRUE" || row[name] === "Y";
  });
  return result;
}































