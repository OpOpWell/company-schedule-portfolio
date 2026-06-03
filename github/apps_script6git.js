/**
 * 社内共有管理 v5.1 GitHub公開用サンプル版
 *
 * - 三角の列非表示方式は廃止
 * - 既読は常時表示
 * - 個人確認列だけを列グループ化
 * - 日付列はカレンダー入力
 * - プルダウンに「空欄に戻す」追加
 * - お知らせ掲示板を追加
 * - 個人ToDoを追加
 * - 実在の個人名をサンプル名（社員A〜社員O）へ置換
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
  "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS],
  "車検管理": ["車両名", "車番", "車検日", "通知", "保険期限", "状態", "写真", READ_HEADER, ...PERSONAL_MEMBERS],
  "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...PERSONAL_MEMBERS],
  "行事予定": ["日付", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "備考"],
  "工事予定": ["開始日", "終了日", "工事名", "現場", "状態", "担当", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "備考"],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS, "メモ"],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...PERSONAL_MEMBERS, "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"],
  "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...PERSONAL_MEMBERS],
  "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...PERSONAL_MEMBERS]
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
    ["完了件数", '=COUNTIF(一覧スケジュール!E2:E,"完了")'],
    ["高重要度お知らせ", '=COUNTIF(お知らせ!E2:E,"高")'],
    ["未完了ToDo", '=COUNTIFS(個人ToDo!C2:C,"<>",個人ToDo!E2:E,"<>完了")'],
    ["今日期限ToDo", '=COUNTIF(個人ToDo!F2:F,"今日")']
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);
  sheet.getRange(3, 1, data.length + 1, 2).setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 4);
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
    "一覧スケジュール",
    "お知らせ",
    "個人ToDo"
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
  }

  addTextRule(typeCol, "車検", "#f4cccc");
  addTextRule(typeCol, "行事", "#fce5cd");
  addTextRule(typeCol, "会議", "#d9d2e9");
  addTextRule(typeCol, "工事予定", "#cfe2f3");
  addTextRule(typeCol, "出先予定", "#d9ead3");
  addTextRule(typeCol, "電話履歴", "#fff2cc");
  addTextRule(typeCol, "作業状況", "#eeeeee");
  addTextRule(typeCol, "ToDo", "#d9ead3");

  addTextRule(importanceCol, "高", "#ea9999");
  addTextRule(importanceCol, "中", "#ffe599");
  addTextRule(importanceCol, "低", "#d9ead3");

  sheet.setConditionalFormatRules(rules);
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


















