function createCompanySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    { name: "出先予定", headers: ["日付", "行き先", "用件", "担当", "社用車", "状態", "既読", "電話対応"] },
    { name: "車検管理", headers: ["車両名", "車番", "車検日", "保険期限", "状態", "写真", "既読", "電話対応"] },
    { name: "会議予定", headers: ["会議名", "日付", "内容", "担当", "状態", "資料", "既読", "電話対応"] },
    { name: "作業状況", headers: ["現場", "作業内容", "担当", "状態", "写真", "備考", "既読", "電話対応"] },
    { name: "工事予定", headers: ["開始日", "終了日", "工事名", "現場", "担当", "状態", "既読", "電話対応", "備考"] },
    { name: "電話履歴", headers: ["日時", "相手", "内容", "担当", "状態", "既読", "電話対応", "メモ"] },
    { name: "一覧スケジュール", headers: ["日付", "種類", "内容", "担当", "状態", "既読", "電話対応", "元シート"] }
  ];

  sheets.forEach(item => {
    let sheet = ss.getSheetByName(item.name);
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
    } else {
      const filter = sheet.getFilter();
      if (filter) filter.remove();
      sheet.clear();
    }

    setupSheet(sheet, item.headers);
  });

  addSampleData(ss);
  refreshAll();
  addPersonalReadColumns();
  console.log("作成完了");
}

function setupSheet(sheet, headers) {
  const statusList = ["予定", "移動中", "進行中", "着工前", "施工中", "完了", "延期", "中止", "未対応", "対応中"];
  const staffList = ["佐藤", "鈴木", "高橋", "田中", "伊藤"];
  const carList = ["1号車", "2号車", "軽トラ", "バン", "未定"];
  const phoneList = ["", "未対応", "対応中", "折返し", "完了"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).createFilter();

  headers.forEach((header, i) => {
    const col = i + 1;

    if (header.includes("日") || header.includes("日時") || header.includes("期限")) {
      sheet.getRange(2, col, 500).setNumberFormat("yyyy/mm/dd");
    }

    if (header === "状態") {
      setDropdown(sheet, col, statusList);
    }

    if (header === "担当") {
      setDropdown(sheet, col, staffList);
    }

    if (header === "社用車") {
      setDropdown(sheet, col, carList);
    }

    if (header === "電話対応") {
      setDropdown(sheet, col, phoneList);
    }

    if (header === "既読") {
      sheet.getRange(2, col, 500).insertCheckboxes();
    }
  });

  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);
}

function setDropdown(sheet, col, list) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, 500).setDataValidation(rule);
}

function addSampleData(ss) {
  ss.getSheetByName("出先予定").getRange(2, 1, 3, 8).setValues([
    [new Date(), "市役所", "書類提出", "佐藤", "1号車", "予定", false, ""],
    [new Date(), "現場A", "進捗確認", "鈴木", "軽トラ", "移動中", false, ""],
    [new Date(new Date().getTime() + 86400000), "取引先B", "打合せ", "高橋", "バン", "予定", false, "未対応"]
  ]);

  ss.getSheetByName("車検管理").getRange(2, 1, 3, 8).setValues([
    ["1号車", "秋田500 あ 12-34", new Date(new Date().getTime() + 20 * 86400000), new Date(new Date().getTime() + 40 * 86400000), "予定", "", false, ""],
    ["2号車", "秋田500 い 56-78", new Date(new Date().getTime() + 60 * 86400000), new Date(new Date().getTime() + 90 * 86400000), "予定", "", false, ""],
    ["軽トラ", "秋田480 う 11-22", new Date(new Date().getTime() - 5 * 86400000), new Date(new Date().getTime() + 10 * 86400000), "延期", "", false, "折返し"]
  ]);

  ss.getSheetByName("会議予定").getRange(2, 1, 3, 8).setValues([
    ["朝礼", new Date(), "本日の予定確認", "佐藤", "完了", "", true, ""],
    ["工程会議", new Date(new Date().getTime() + 86400000), "来週工程の確認", "田中", "予定", "", false, ""],
    ["安全会議", new Date(new Date().getTime() + 3 * 86400000), "安全対策確認", "伊藤", "予定", "", false, ""]
  ]);

  ss.getSheetByName("作業状況").getRange(2, 1, 3, 8).setValues([
    ["現場A", "水路工", "鈴木", "進行中", "", "午前中作業", false, ""],
    ["現場B", "掘削", "高橋", "完了", "", "写真確認済み", true, ""],
    ["現場C", "整地", "田中", "予定", "", "", false, ""]
  ]);

  ss.getSheetByName("工事予定").getRange(2, 1, 3, 9).setValues([
    [new Date(), new Date(new Date().getTime() + 3 * 86400000), "水路工", "現場A", "鈴木", "施工中", false, "", "資材確認"],
    [new Date(new Date().getTime() + 2 * 86400000), new Date(new Date().getTime() + 5 * 86400000), "掘削工", "現場B", "高橋", "予定", false, "未対応", ""],
    [new Date(new Date().getTime() + 7 * 86400000), new Date(new Date().getTime() + 10 * 86400000), "整地工", "現場C", "田中", "着工前", false, "", ""]
  ]);

  ss.getSheetByName("電話履歴").getRange(2, 1, 3, 8).setValues([
    [new Date(), "市役所", "書類確認の電話", "佐藤", "未対応", false, "折返し", "午後折返し"],
    [new Date(), "取引先B", "打合せ日程確認", "高橋", "対応中", false, "", ""],
    [new Date(), "協力会社", "資材搬入確認", "鈴木", "完了", true, "", ""]
  ]);
}

function createScheduleList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("一覧スケジュール");
  if (!sheet) return;

  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.clear();

  const headers = ["日付", "種類", "内容", "担当", "状態", "既読", "電話対応", "元シート"];
  setupSheet(sheet, headers);

  const rows = [];

  collectRows(ss, "出先予定", 8, row => row[0], row => [
    row[0], "出先予定", row[1] + " / " + row[2], row[3], row[5], row[6], row[7], "出先予定"
  ], rows);

  collectRows(ss, "会議予定", 8, row => row[1], row => [
    row[1], "会議", row[0] + " / " + row[2], row[3], row[4], row[6], row[7], "会議予定"
  ], rows);

  collectRows(ss, "車検管理", 8, row => row[2], row => [
    row[2], "車検", row[0] + " / " + row[1], "", row[4], row[6], row[7], "車検管理"
  ], rows);

  collectRows(ss, "作業状況", 8, row => row[0], row => [
    new Date(), "作業状況", row[0] + " / " + row[1], row[2], row[3], row[6], row[7], "作業状況"
  ], rows);

  collectRows(ss, "工事予定", 9, row => row[0], row => [
    row[0], "工事予定", row[2] + " / " + row[3], row[4], row[5], row[6], row[7], "工事予定"
  ], rows);

  collectRows(ss, "電話履歴", 8, row => row[0], row => [
    row[0], "電話履歴", row[1] + " / " + row[2], row[3], row[4], row[5], row[6], "電話履歴"
  ], rows);

  rows.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);
  applyColorRules(sheet);
}

function collectRows(ss, sheetName, colCount, condition, mapper, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();

  values.forEach(row => {
    if (condition(row)) rows.push(mapper(row));
  });
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("ダッシュボード");
  if (!sheet) {
    sheet = ss.insertSheet("ダッシュボード");
  } else {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.clear();
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
    ["出先予定", '=COUNTIF(一覧スケジュール!B2:B,"出先予定")'],
    ["会議予定", '=COUNTIF(一覧スケジュール!B2:B,"会議")'],
    ["工事予定", '=COUNTIF(一覧スケジュール!B2:B,"工事予定")'],
    ["作業状況", '=COUNTIF(一覧スケジュール!B2:B,"作業状況")'],
    ["電話履歴", '=COUNTIF(一覧スケジュール!B2:B,"電話履歴")'],
    ["30日以内の車検", '=COUNTIFS(車検管理!C2:C,">="&TODAY(),車検管理!C2:C,"<="&TODAY()+30)'],
    ["未読件数", '=COUNTIF(一覧スケジュール!F2:F,FALSE)'],
    ["未対応電話", '=COUNTIF(一覧スケジュール!G2:G,"未対応")'],
    ["折返し電話", '=COUNTIF(一覧スケジュール!G2:G,"折返し")'],
    ["進行中/施工中", '=COUNTIF(一覧スケジュール!E2:E,"進行中")+COUNTIF(一覧スケジュール!E2:E,"施工中")'],
    ["完了件数", '=COUNTIF(一覧スケジュール!E2:E,"完了")']
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);
  sheet.getRange(3, 1, data.length + 1, 2).setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(1, 2);
}

function refreshAll() {
  createScheduleList();
  createDashboard();
  console.log("全体更新完了");
}

function addPersonalReadColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ["電話履歴", "工事予定", "会議予定", "作業状況", "出先予定", "車検管理", "一覧スケジュール"];
  const members = ["佐藤確認", "鈴木確認", "高橋確認", "田中確認", "伊藤確認"];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    members.forEach(member => {
      let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      if (!headers.includes(member)) {
        sheet.insertColumnAfter(sheet.getLastColumn());
        const col = sheet.getLastColumn();

        sheet.getRange(1, col).setValue(member)
          .setFontWeight("bold")
          .setBackground("#d9ead3")
          .setHorizontalAlignment("center");

        sheet.getRange(2, col, 500).insertCheckboxes();
      }
    });

    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  console.log("個人確認列を追加しました");
}

function togglePersonalCheckColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ["電話履歴", "工事予定", "会議予定", "作業状況", "出先予定", "車検管理", "一覧スケジュール"];
  const members = ["佐藤確認", "鈴木確認", "高橋確認", "田中確認", "伊藤確認"];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const cols = members
      .map(name => headers.indexOf(name) + 1)
      .filter(col => col > 0);

    if (cols.length === 0) return;

    if (sheet.isColumnHiddenByUser(cols[0])) {
      cols.forEach(col => sheet.showColumns(col));
    } else {
      cols.forEach(col => sheet.hideColumns(col));
    }
  });

  console.log("個人確認列を開閉しました");
}

function hidePersonalCheckColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ["電話履歴", "工事予定", "会議予定", "作業状況", "出先予定", "車検管理", "一覧スケジュール"];
  const members = ["佐藤確認", "鈴木確認", "高橋確認", "田中確認", "伊藤確認"];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    members.forEach(name => {
      const col = headers.indexOf(name) + 1;
      if (col > 0) sheet.hideColumns(col);
    });
  });

  console.log("個人確認列を非表示にしました");
}

function showPersonalCheckColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = ["電話履歴", "工事予定", "会議予定", "作業状況", "出先予定", "車検管理", "一覧スケジュール"];
  const members = ["佐藤確認", "鈴木確認", "高橋確認", "田中確認", "伊藤確認"];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    members.forEach(name => {
      const col = headers.indexOf(name) + 1;
      if (col > 0) sheet.showColumns(col);
    });
  });

  console.log("個人確認列を表示しました");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addItem("全体更新", "refreshAll")
    .addItem("個人確認列を追加", "addPersonalReadColumns")
    .addItem("個人確認列を開く/閉じる", "togglePersonalCheckColumns")
    .addItem("個人確認列を隠す", "hidePersonalCheckColumns")
    .addItem("個人確認列を表示", "showPersonalCheckColumns")
    .addToUi();
}

function onEdit(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  if (row === 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const readCol = headers.indexOf("既読") + 1;

  if (readCol <= 0) return;

  const rowData = sheet.getRange(row, 1, 1, readCol - 1).getValues()[0];
  const hasData = rowData.some(v => v !== "");

  const cell = sheet.getRange(row, readCol);

  if (hasData) {
    cell.insertCheckboxes();
  } else {
    cell.clearDataValidations();
    cell.clearContent();
  }
}

function applyColorRules(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = headers.indexOf("状態") + 1;
  const phoneCol = headers.indexOf("電話対応") + 1;
  const readCol = headers.indexOf("既読") + 1;

  const rules = [];

  if (statusCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("完了")
        .setBackground("#b6d7a8")
        .setRanges([sheet.getRange(2, statusCol, 500)])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("進行中")
        .setBackground("#cfe2f3")
        .setRanges([sheet.getRange(2, statusCol, 500)])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("施工中")
        .setBackground("#cfe2f3")
        .setRanges([sheet.getRange(2, statusCol, 500)])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("延期")
        .setBackground("#ffe599")
        .setRanges([sheet.getRange(2, statusCol, 500)])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("中止")
        .setBackground("#ea9999")
        .setRanges([sheet.getRange(2, statusCol, 500)])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("未対応")
        .setBackground("#f4cccc")
        .setRanges([sheet.getRange(2, statusCol, 500)])
        .build()
    );
  }

  if (phoneCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("未対応")
        .setBackground("#f4cccc")
        .setRanges([sheet.getRange(2, phoneCol, 500)])
        .build()
    );

    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("折返し")
        .setBackground("#fce5cd")
        .setRanges([sheet.getRange(2, phoneCol, 500)])
        .build()
    );
  }

  if (readCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied("=$" + columnLetter(readCol) + "2=FALSE")
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(2, readCol, 500)])
        .build()
    );
  }

  sheet.setConditionalFormatRules(rules);
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

function normalizePersonalColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    "電話履歴",
    "工事予定",
    "会議予定",
    "作業状況",
    "出先予定",
    "車検管理",
    "一覧スケジュール"
  ];

  const members = [
    "佐藤確認",
    "鈴木確認",
    "高橋確認",
    "田中確認",
    "伊藤確認"
  ];

  sheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // まず不足している個人確認列を追加
    members.forEach(name => {
      let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers.includes(name)) {
        sheet.insertColumnAfter(sheet.getLastColumn());
        const col = sheet.getLastColumn();
        sheet.getRange(1, col).setValue(name);
        sheet.getRange(2, col, 500).insertCheckboxes();
      }
    });

    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const phoneCol = headers.indexOf("電話対応") + 1;
    const readCol = headers.indexOf("既読") + 1;

    const baseCol = phoneCol > 0 ? phoneCol : readCol;
    if (baseCol <= 0) return;

    // 個人確認列を右端へ一度退避
    members.forEach(name => {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const col = headers.indexOf(name) + 1;
      if (col > 0) {
        sheet.moveColumns(
          sheet.getRange(1, col, sheet.getMaxRows(), 1),
          sheet.getLastColumn() + 1
        );
      }
    });

    SpreadsheetApp.flush();

    // 電話対応または既読の右へ、順番通りに戻す
    members.slice().reverse().forEach(name => {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      const currentCol = headers.indexOf(name) + 1;
      const newHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      const targetPhoneCol = newHeaders.indexOf("電話対応") + 1;
      const targetReadCol = newHeaders.indexOf("既読") + 1;
      const targetBaseCol = targetPhoneCol > 0 ? targetPhoneCol : targetReadCol;

      if (currentCol > 0 && targetBaseCol > 0) {
        sheet.moveColumns(
          sheet.getRange(1, currentCol, sheet.getMaxRows(), 1),
          targetBaseCol + 1
        );
      }
    });

    sheet.autoResizeColumns(1, sheet.getLastColumn());
  });

  console.log("個人確認列を統一しました");
}
function hidePersonalCheckColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    "電話履歴",
    "工事予定",
    "会議予定",
    "作業状況",
    "出先予定",
    "車検管理",
    "一覧スケジュール"
  ];

  const members = [
    "佐藤確認",
    "鈴木確認",
    "高橋確認",
    "田中確認",
    "伊藤確認"
  ];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    members.forEach(name => {
      const col = headers.indexOf(name) + 1;
      if (col > 0) sheet.hideColumns(col);
    });
  });
}

function togglePersonalCheckColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    "電話履歴",
    "工事予定",
    "会議予定",
    "作業状況",
    "出先予定",
    "車検管理",
    "一覧スケジュール"
  ];

  const members = [
    "佐藤確認",
    "鈴木確認",
    "高橋確認",
    "田中確認",
    "伊藤確認"
  ];

  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const cols = members
      .map(name => headers.indexOf(name) + 1)
      .filter(col => col > 0);

    if (cols.length === 0) return;

    if (sheet.isColumnHiddenByUser(cols[0])) {
      cols.forEach(col => sheet.showColumns(col));
    } else {
      cols.forEach(col => sheet.hideColumns(col));
    }
  });
}
function hidePersonalOnlyActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();

  const members = [
    "佐藤確認",
    "鈴木確認",
    "高橋確認",
    "田中確認",
    "伊藤確認"
  ];

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  members.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.hideColumns(col);
    }
  });
}

function showPersonalOnlyActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();

  const members = [
    "佐藤確認",
    "鈴木確認",
    "高橋確認",
    "田中確認",
    "伊藤確認"
  ];

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  members.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col > 0) {
      sheet.showColumns(col);
    }
  });
}

function testHideColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("工事予定");

  sheet.hideColumns(10, 5); // J～N
}




function toggleKoujiPersonalColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("工事予定");

  // J～N = 佐藤確認～伊藤確認
  if (sheet.isColumnHiddenByUser(10)) {
    sheet.showColumns(10, 5);
  } else {
    sheet.hideColumns(10, 5);
  }
}
function groupKoujiPersonalColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("工事予定");

  // J～N をグループ化
  sheet.getRange("J:N").shiftColumnGroupDepth(1);
}



















