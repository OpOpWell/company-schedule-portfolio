/**
 * 社内共有管理 v6.1 GitHub公開用サンプル版
 *
 * - 三角の列非表示方式は廃止
 * - 既読は常時表示
 * - 個人確認列だけを列グループ化
 * - 日付列はカレンダー入力
 * - プルダウンに「空欄に戻す」追加
 * - お知らせ掲示板を追加
 * - 個人ToDoを追加
 * - 実在の個人名をサンプル名（社員A〜社員O）へ置換
 * - Googleカレンダー連携を追加
 * - 社用車予約を追加
 * - お知らせを要確認一覧へ反映
 * - 担当別未読件数を追加
 * - 社用車予約の開始時刻・終了時刻をGoogleカレンダーへ反映
 */

const PERSONAL_MEMBERS = [
  "社員A",
  "社員B",
  "社員C",
  "社員D",
  "社員E",
  "社員F",
  "社員G",
  "社員H",
  "社員I",
  "社員J",
  "社員K",
  "社員L",
  "社員M",
  "社員N",
  "社員O",

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
  "工事予定": ["開始日", "終了日", "工事名", "現場", "状態", "担当", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "備考", CALENDAR_ID_HEADER],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "メモ"],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"],
  "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考", CALENDAR_ID_HEADER],
  "担当別未読": ["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"]
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
  const statusList = [CLEAR_LABEL, "予定", "移動中", "進行中", "着工前", "施工中", "完了", "延期", "中止", "未対応", "対応中"];
  const staffList = [CLEAR_LABEL, ...PERSONAL_MEMBERS];

  const carList = [CLEAR_LABEL, "1号車", "2号車", "軽トラ", "バン", "未定"];
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
    "既読": row["既読"] === true,
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
  createScheduleList();
  createAlertList();
  createAssigneeUnreadSummary();
  createDashboard();
  resetColumnGroups();
  createCheckGroup();
}

function setCheckboxesForDataRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const checkNames = [READ_HEADER].concat(PERSONAL_MEMBERS);

  checkNames.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;

    sheet.getRange(2, col, sheet.getMaxRows() - 1).clearDataValidations();
    sheet.getRange(2, col, sheet.getMaxRows() - 1).clearContent();

    if (lastRow >= 2) {
      sheet.getRange(2, col, lastRow - 1).insertCheckboxes();
    }
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

  updateNoticeForEditedRow(sheet, row, col);
  updateCheckboxesForEditedRow(sheet, row);

  createScheduleList();
  createAlertList();
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
    "社用車予約"
  ];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addItem("全体更新", "refreshAll")
    .addItem("要確認一覧を更新", "createAlertList")
    .addItem("担当別未読を更新", "createAssigneeUnreadSummary")
    .addItem("Googleカレンダーへ反映", "syncCalendarEvents")
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

    const start = headers.indexOf("社員A") + 1;
    const end = headers.indexOf("社員O") + 1;

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

  const rows = PERSONAL_MEMBERS.map(member => {
    return {
      "担当": member,
      "未読件数": countUnreadByMember(schedule, member),
      "未完了ToDo": countTodoByMember(todo, member),
      "未対応電話": countPhoneByMember(phone, member),
      "今日の予定": countTodayScheduleByMember(schedule, member)
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

function SHEET_ID(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "";
  return sheet.getSheetId();
}


// サンプルデータ投入用
function addSampleData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("社用車予約");

  const data = [
    ["2026/06/03","09:00","12:00","1号車","社員A","市役所","書類提出","予定","","",""],
    ["2026/06/03","13:00","17:00","軽トラ","社員B","現場A","現場確認","予定","","",""],
    ["2026/06/04","08:30","16:00","バン","社員C","客先","打合せ","予定","","",""]
  ];

  sheet.getRange(2,1,data.length,data[0].length).setValues(data);
}




















