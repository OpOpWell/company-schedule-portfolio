/**
 * 社内共有管理 v4.0 統一版
 *
 * - 三角の列非表示方式は廃止
 * - 既読は常時表示
 * - 佐藤確認〜伊藤確認だけを列グループ化
 * - 日付列はカレンダー入力
 * - プルダウンに「空欄に戻す」追加
 */

const PERSONAL_MEMBERS = ["佐藤確認", "鈴木確認", "高橋確認", "田中確認", "伊藤確認"];
const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";

const DATE_HEADERS = [
  "日付",
  "日時",
  "開始日",
  "終了日",
  "車検日",
  "保険期限"
];

const SHEET_HEADERS = {
  "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS],
  "車検管理": ["車両名", "車番", "車検日", "通知", "保険期限", "状態", "写真", READ_HEADER, ...PERSONAL_MEMBERS],
  "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...PERSONAL_MEMBERS],
  "行事予定": ["日付", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考"],
  "工事予定": ["開始日", "終了日", "工事名", "現場", "状態", "担当", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "備考"],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "メモ"],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"]
};

function createCompanySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SHEET_HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    resetSheet(sheet, SHEET_HEADERS[name].length);
    setupSheet(sheet, SHEET_HEADERS[name]);
  });

  addSampleData(ss);
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
  const staffList = [CLEAR_LABEL, "佐藤", "鈴木", "高橋", "田中", "伊藤"];
  const carList = [CLEAR_LABEL, "1号車", "2号車", "軽トラ", "バン", "未定"];
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];

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
  });

  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);
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

function addSampleData(ss) {
  const today = new Date();

  writeObjectsToSheet(ss.getSheetByName("出先予定"), [
    {"日付": today, "行き先": "市役所", "用件": "書類提出", "担当": "佐藤", "社用車": "1号車", "状態": "予定", "通知": getNoticeText(today), "電話対応": ""},
    {"日付": today, "行き先": "現場A", "用件": "進捗確認", "担当": "鈴木", "社用車": "軽トラ", "状態": "移動中", "通知": getNoticeText(today), "電話対応": ""},
    {"日付": addDays(1), "行き先": "取引先B", "用件": "打合せ", "担当": "高橋", "社用車": "バン", "状態": "予定", "通知": getNoticeText(addDays(1)), "電話対応": "未対応"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("車検管理"), [
    {"車両名": "1号車", "車番": "秋田500 あ 12-34", "車検日": addDays(20), "通知": getNoticeText(addDays(20)), "保険期限": addDays(40), "状態": "予定", "写真": ""},
    {"車両名": "2号車", "車番": "秋田500 い 56-78", "車検日": addDays(60), "通知": getNoticeText(addDays(60)), "保険期限": addDays(90), "状態": "予定", "写真": ""},
    {"車両名": "軽トラ", "車番": "秋田480 う 11-22", "車検日": addDays(-5), "通知": getNoticeText(addDays(-5)), "保険期限": addDays(10), "状態": "延期", "写真": ""}
  ]);

  writeObjectsToSheet(ss.getSheetByName("会議予定"), [
    {"日付": today, "会議名": "朝礼", "内容": "本日の予定確認", "担当": "佐藤", "状態": "完了", "通知": getNoticeText(today), "資料": "", "既読": true},
    {"日付": addDays(1), "会議名": "工程会議", "内容": "来週工程の確認", "担当": "田中", "状態": "予定", "通知": getNoticeText(addDays(1)), "資料": ""},
    {"日付": addDays(3), "会議名": "安全会議", "内容": "安全対策確認", "担当": "伊藤", "状態": "予定", "通知": getNoticeText(addDays(3)), "資料": ""}
  ]);

  writeObjectsToSheet(ss.getSheetByName("行事予定"), [
    {"日付": today, "行事名": "安全大会", "内容": "社内安全行事", "担当": "佐藤", "状態": "予定", "通知": getNoticeText(today)},
    {"日付": addDays(2), "行事名": "健康診断", "内容": "午前中実施", "担当": "鈴木", "状態": "予定", "通知": getNoticeText(addDays(2))},
    {"日付": addDays(5), "行事名": "社内清掃", "内容": "月例清掃", "担当": "田中", "状態": "予定", "通知": getNoticeText(addDays(5))}
  ]);

  writeObjectsToSheet(ss.getSheetByName("作業状況"), [
    {"現場": "現場A", "作業内容": "水路工", "状態": "進行中", "担当": "鈴木", "写真": "", "通知": "", "備考": "午前中作業"},
    {"現場": "現場B", "作業内容": "掘削", "状態": "完了", "担当": "高橋", "写真": "", "通知": "", "備考": "写真確認済み", "既読": true},
    {"現場": "現場C", "作業内容": "整地", "状態": "予定", "担当": "田中", "写真": "", "通知": "", "備考": ""}
  ]);

  writeObjectsToSheet(ss.getSheetByName("工事予定"), [
    {"開始日": today, "終了日": addDays(3), "工事名": "水路工", "現場": "現場A", "状態": "施工中", "担当": "鈴木", "通知": getNoticeText(today), "電話対応": "", "備考": "資材確認"},
    {"開始日": addDays(2), "終了日": addDays(5), "工事名": "掘削工", "現場": "現場B", "状態": "予定", "担当": "高橋", "通知": getNoticeText(addDays(2)), "電話対応": "未対応", "備考": ""},
    {"開始日": addDays(7), "終了日": addDays(10), "工事名": "整地工", "現場": "現場C", "状態": "着工前", "担当": "田中", "通知": getNoticeText(addDays(7)), "電話対応": "", "備考": ""}
  ]);

  writeObjectsToSheet(ss.getSheetByName("電話履歴"), [
    {"日時": today, "相手": "市役所", "内容": "書類確認の電話", "電話対応": "折返し", "担当": "佐藤", "状態": "未対応", "通知": getNoticeText(today), "メモ": "午後折返し"},
    {"日時": today, "相手": "取引先B", "内容": "打合せ日程確認", "電話対応": "", "担当": "高橋", "状態": "対応中", "通知": getNoticeText(today), "メモ": ""},
    {"日時": today, "相手": "協力会社", "内容": "資材搬入確認", "電話対応": "", "担当": "鈴木", "状態": "完了", "通知": getNoticeText(today), "メモ": "", "既読": true}
  ]);
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

  rows.sort((a, b) => new Date(a["日付"]) - new Date(b["日付"]));

  writeObjectsToSheet(sheet, rows);
  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);
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
    ["完了件数", '=COUNTIF(一覧スケジュール!E2:E,"完了")']
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);
  sheet.getRange(3, 1, data.length + 1, 2).setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 4);

  applyDashboardColorRules(sheet);
}

function refreshAll() {
  createScheduleList();
  createAlertList();
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
  const priority = ["日付", "日時", "開始日", "車検日", "保険期限", "終了日"];

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
    "一覧スケジュール"
  ];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addItem("全体更新", "refreshAll")
    .addItem("要確認一覧を更新", "createAlertList")
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

    const start = headers.indexOf("佐藤確認") + 1;
    const end = headers.indexOf("伊藤確認") + 1;

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
  }

  addTextRule(typeCol, "車検", "#f4cccc");
  addTextRule(typeCol, "行事", "#fce5cd");
  addTextRule(typeCol, "会議", "#d9d2e9");
  addTextRule(typeCol, "工事予定", "#cfe2f3");
  addTextRule(typeCol, "出先予定", "#d9ead3");
  addTextRule(typeCol, "電話履歴", "#fff2cc");
  addTextRule(typeCol, "作業状況", "#eeeeee");

  sheet.setConditionalFormatRules(rules);
}

function applyDashboardColorRules(sheet) {
  sheet.setConditionalFormatRules([]);
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

function addDays(days) {
  return new Date(Date.now() + days * 86400000);
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

