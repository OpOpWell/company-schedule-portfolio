/**
 * 社内共有管理 v8.2 社用調整版（設定シート・免許資格・備品修理・帳簿PDF追加）
 *
 * - 担当者15人と既読確認者3人を分離
 * - 既読確認列は安倍・三好・佐藤の3列のみ
 * - 担当プルダウンは安倍を先頭にした15人
 * - 工事予定の列順を「工事名、現場、開始日、終了日」へ変更
 * - Googleカレンダー時間連携を維持
 * - 社用車予約の重複チェックを追加
 * - 月ごとの行グループ化関数を追加
 * - 過去一覧シートを追加
 * - 期限切れ・過去予定を過去一覧へ移動する関数を追加
 * - 社用運用向けにサンプルデータ投入関数を削除
 * - 日報シートを追加
 * - 選択行の日報文章をGPTで作成
 * - 選択行の日報PDFを作成
 * - 免許資格管理シートを追加
 * - 備品修理管理シートを追加
 * - 他部門共有用の一覧帳簿PDF出力を追加
 * - 担当者名・個人既読者名を「設定」シートから読み込む方式に変更
 * - 各入力シートに備考欄を追加
 * - 行事予定を開始日・終了日で管理できるように変更
 * - 状態プルダウンをシート別に分離
 * - 要確認一覧は毎日/編集時更新、過去一覧移動は毎月実行に分離
 * - 未返却備品・未完了ToDo・未対応電話などは過去一覧へ移動しない安全条件を追加
 */

/**
 * 社用メモ
 * - 担当者・既読確認者は「設定」シートから読み込みます。
 * - 設定シートが未作成の場合は DEFAULT_STAFF_MEMBERS / DEFAULT_PERSONAL_MEMBERS を使います。
 * - GitHub公開用は「設定」シートの名前だけサンプル名に差し替えます。
 */

const DEFAULT_STAFF_MEMBERS = [
  "安倍",
  "佐藤",
  "後藤",
  "三好",
  "小田",
  "豊田",
  "三浦",
  "斎藤",
  "渡辺",
  "加藤",
  "稲田",
  "阿藤",
  "小野",
  "菊田",
  "須田",
];

const DEFAULT_PERSONAL_MEMBERS = [
  "安倍",
  "三好",
  "佐藤"
];

const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";
const CALENDAR_ID_HEADER = "カレンダーID";

const DATE_HEADERS = [
  "日付",
  "日時",
  "開始日",
  "終了日",
  "車検期限",
  "次回車検期限",
  "保険期限",
  "投稿日",
  "登録日",
  "期限",
  "取得日",
  "更新期限",
  "修理依頼日",
  "返却予定日"
];

function getSheetHeaders_() {
  const personalMembers = getPersonalMembers_();

  return {
  "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER],
  "車検管理": ["車両名", "車番", "車検期限", "次回車検期限", "通知", "保険期限", "状態", "写真", READ_HEADER, ...personalMembers, "備考"],
  "車検履歴": ["更新日", "車両名", "車番", "旧車検期限", "新車検期限", "保険期限", "担当", "備考"],
  "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER],
  "行事予定": ["開始日", "終了日", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER],
  "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...personalMembers, "備考"],
  "工事予定": ["工事名", "現場", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "状態", "通知", READ_HEADER, ...personalMembers, "備考"],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "備考", "元シート"],
  "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...personalMembers, "備考"],
  "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...personalMembers, "備考"],
  "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER],
  "日報": ["日付", "担当", "現場", "作業内容", "進捗", "問題点", "明日の予定", "写真", "日報文章", "状態", "備考", READ_HEADER, ...personalMembers],
  "免許資格管理": ["担当", "資格名", "区分", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personalMembers, "備考"],
  "備品修理管理": ["登録日", "備品名", "場所", "内容", "担当", "修理依頼日", "返却予定日", "返却済", "状態", "通知", READ_HEADER, ...personalMembers, "備考"],
  "帳簿PDF履歴": ["作成日時", "対象", "期間", "PDFリンク", "備考"],
  "担当別未読": ["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"],
  "過去一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート", "備考", "移動日"]
};
}



function ensureSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("設定");

  if (!sheet) {
    sheet = ss.insertSheet("設定");
  }

  const headers = ["設定種別", "値", "備考"];
  const currentHeader = sheet.getRange(1, 1, 1, 3).getValues()[0];
  const hasHeader = currentHeader[0] === "設定種別" && currentHeader[1] === "値";

  if (!hasHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#d9ead3")
      .setHorizontalAlignment("center");

    const rows = [];

    DEFAULT_STAFF_MEMBERS.forEach(name => {
      rows.push(["担当者", name, "担当プルダウン用。社用版/GitHub版はここを差し替え"]);
    });

    DEFAULT_PERSONAL_MEMBERS.forEach(name => {
      rows.push(["既読確認者", name, "個人既読列用。列名として使われます"]);
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}

let SETTINGS_CACHE_ = null;

function getSettingsCache_() {
  if (SETTINGS_CACHE_) return SETTINGS_CACHE_;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("設定");

  const cache = {
    "担当者": DEFAULT_STAFF_MEMBERS.slice(),
    "既読確認者": DEFAULT_PERSONAL_MEMBERS.slice()
  };

  if (!sheet || sheet.getLastRow() < 2) {
    SETTINGS_CACHE_ = cache;
    return SETTINGS_CACHE_;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const staff = [];
  const personal = [];

  values.forEach(row => {
    const type = String(row[0] || "").trim();
    const value = String(row[1] || "").trim();

    if (!value) return;
    if (type === "担当者") staff.push(value);
    if (type === "既読確認者") personal.push(value);
  });

  if (staff.length > 0) cache["担当者"] = staff;
  if (personal.length > 0) cache["既読確認者"] = personal;

  SETTINGS_CACHE_ = cache;
  return SETTINGS_CACHE_;
}

function getSettingValues_(settingType, defaultValues) {
  const cache = getSettingsCache_();
  return cache[settingType] && cache[settingType].length > 0
    ? cache[settingType].slice()
    : defaultValues.slice();
}

function clearSettingsCache_() {
  SETTINGS_CACHE_ = null;
}

function getStaffMembers_() {
  return getSettingValues_("担当者", DEFAULT_STAFF_MEMBERS);
}

function getPersonalMembers_() {
  return getSettingValues_("既読確認者", DEFAULT_PERSONAL_MEMBERS);
}


/**
 * シートごとの状態プルダウン
 *
 * 全シート共通の状態候補にすると、入力ミスや運用ミスが増えるため、
 * 入力シートごとに必要な状態だけ表示します。
 */
function getStatusListForSheet_(sheetName) {
  const map = {
    "出先予定": [CLEAR_LABEL, "予定", "移動中", "完了", "延期", "中止"],
    "工事予定": [CLEAR_LABEL, "着工前", "施工中", "完了", "延期", "中止"],
    "会議予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "行事予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "作業状況": [CLEAR_LABEL, "着工前", "施工中", "完了", "中止", "要確認"],
    "電話履歴": [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"],
    "社用車予約": [CLEAR_LABEL, "予定", "使用中", "返却済", "中止", "予約重複"],
    "日報": [CLEAR_LABEL, "下書き", "提出済", "確認済", "差戻し", "完了"],
    "免許資格管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "備品修理管理": [CLEAR_LABEL, "未依頼", "依頼済", "修理中", "返却待ち", "返却済", "完了", "中止"],
    "車検管理": [CLEAR_LABEL, "未対応", "予約済", "実施済", "完了", "更新済"],
    "個人ToDo": [CLEAR_LABEL, "未着手", "対応中", "完了", "中止"],
    "お知らせ": [CLEAR_LABEL, "要確認", "確認済", "完了"],
    "一覧スケジュール": [CLEAR_LABEL, "予定", "移動中", "着工前", "施工中", "未対応", "対応中", "要確認", "完了", "返却済", "更新済", "延期", "中止"],
    "要確認一覧": [CLEAR_LABEL, "要確認", "未対応", "対応中", "期限切れ", "完了", "返却済", "更新済"],
    "過去一覧": [CLEAR_LABEL, "完了", "返却済", "更新済", "確認済", "中止"]
  };

  return map[sheetName] || [CLEAR_LABEL, "予定", "進行中", "完了", "延期", "中止", "未対応", "対応中", "要確認"];
}

function getCopyStatusList_() {
  return [CLEAR_LABEL, "有", "無", "未確認"];
}

function getReturnedStatusList_() {
  return [CLEAR_LABEL, "未", "済"];
}

function setupSettingsSheet() {
  ensureSettingsSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast("設定シートを作成/確認しました");
}

function applySettingsToSheets() {
  clearSettingsCache_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_();

  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    // 既存データを消さずに、ヘッダー・入力規則・既読列だけ再調整します。
    setupSheetWithoutClearing_(sheet, headersMap[name]);
  });

  // 設定反映後は軽量更新だけ行います。
  refreshAll();
  SpreadsheetApp.getActiveSpreadsheet().toast("設定シートの担当者名を反映しました");
}

function setupSheetWithoutClearing_(sheet, headers) {
  ensureVehicleInspectionNextColumn_(sheet);

  const currentMaxCols = sheet.getMaxColumns();

  if (currentMaxCols < headers.length) {
    sheet.insertColumnsAfter(currentMaxCols, headers.length - currentMaxCols);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);

  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).createFilter();
  } catch (e) {}

  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).clearDataValidations();

  const sheetName = sheet.getName();
  const statusList = getStatusListForSheet_(sheetName);
  const staffList = [CLEAR_LABEL, ...getStaffMembers_()];
  const carList = [
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
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
  const importanceList = [CLEAR_LABEL, "高", "中", "低"];
  const copyList = getCopyStatusList_();
  const returnedList = getReturnedStatusList_();

  headers.forEach((header, i) => {
    const col = i + 1;

    if (DATE_HEADERS.includes(header)) setDateColumn(sheet, col);
    if (header === "状態") setDropdown(sheet, col, statusList);
    if (header === "担当") setDropdown(sheet, col, staffList);
    if (header === "利用者") setDropdown(sheet, col, staffList);
    if (header === "社用車") setDropdown(sheet, col, carList);
    if (header === "電話対応") setDropdown(sheet, col, phoneList);
    if (header === "重要度") setDropdown(sheet, col, importanceList);
    if (header === "コピー有無") setDropdown(sheet, col, copyList);
    if (header === "返却済") setDropdown(sheet, col, returnedList);
  });

  setCheckboxesForDataRows(sheet);
  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    try {
      sheet.hideColumns(calendarIdCol);
    } catch (e) {}
  }
}

function createCompanySheets() {
  clearSettingsCache_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_();

  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    let sheet = ss.getSheetByName(name);
    const headers = headersMap[name];

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    // 既存の車検管理だけは、ヘッダー上書き前に「次回車検期限」列を安全に追加する。
    if (name === "車検管理") {
      ensureVehicleInspectionNextColumn_(sheet);
    }

    // createCompanySheets は再現用の軽量版。
    // 入力規則・色付け・チェックボックス・フィルタは applySettingsToSheets で別実行する。
    setupHeaderOnly_(sheet, headers);
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "シート構成を作成しました。次に applySettingsToSheets を実行してください。"
  );

  console.log("シート構成作成完了（軽量版）");
}

/**
 * ヘッダーだけを軽く作成する。
 * createCompanySheets から呼び出す再現用処理。
 */
function setupHeaderOnly_(sheet, headers) {
  if (!sheet || !headers || headers.length === 0) return;

  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < headers.length) {
    sheet.insertColumnsAfter(currentMaxCols, headers.length - currentMaxCols);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    try {
      sheet.hideColumns(calendarIdCol);
    } catch (e) {}
  }
}

function rebuildCompanySheetsDangerously_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_();

  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    resetSheet(sheet, headersMap[name].length);
    setupSheet(sheet, headersMap[name]);
  });

  refreshAllHeavy();

  console.log("全シート再作成完了");
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
  const sheetName = sheet.getName();
  const statusList = getStatusListForSheet_(sheetName);
  const staffList = [CLEAR_LABEL, ...getStaffMembers_()];

  const carList = [
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
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
  const importanceList = [CLEAR_LABEL, "高", "中", "低"];
  const copyList = getCopyStatusList_();
  const returnedList = getReturnedStatusList_();

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
    if (header === "利用者") setDropdown(sheet, col, staffList);
    if (header === "社用車") setDropdown(sheet, col, carList);
    if (header === "電話対応") setDropdown(sheet, col, phoneList);
    if (header === "重要度") setDropdown(sheet, col, importanceList);
    if (header === "コピー有無") setDropdown(sheet, col, copyList);
    if (header === "返却済") setDropdown(sheet, col, returnedList);
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
    if (getPersonalMembers_().includes(header)) return obj[header] === true;
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

  const headers = getSheetHeaders_()["一覧スケジュール"];

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

    "備考": row["備考"] || "",
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

    "備考": row["備考"] || "",
    "元シート": "会議予定"
  }), rows);

  collectRowsAsObjects(ss, "行事予定", row => row["開始日"] || row["日付"], row => ({
    "日付": row["開始日"] || row["日付"],
    "種類": "行事",
    "内容": joinText(row["行事名"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["開始日"] || row["日付"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "備考": row["備考"] || "",
    "元シート": "行事予定"
  }), rows);

  collectRowsAsObjects(ss, "車検管理", row => row["車検期限"], row => ({
  "日付": row["車検期限"],
  "種類": "車検",
  "内容": joinText(row["車両名"], row["車番"]),
  "担当": "",
  "状態": row["状態"],
  "通知": getNoticeText(row["車検期限"]),
  "電話対応": "",
  "既読":
    row["既読"] === true ||
    row["既読"] === "TRUE" ||
    row["既読"] === "Y",
  ...copyPersonalChecks(row),
  "備考": row["備考"] || "",
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

    "備考": row["備考"] || "",
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

    "備考": row["備考"] || "",
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

    "備考": row["備考"] || row["メモ"] || "",
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

    "備考": row["備考"] || "",
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

    "備考": row["備考"] || "",
    "元シート": "社用車予約"
  }), rows);



  collectRowsAsObjects(ss, "日報", row => row["日付"] || row["現場"] || row["作業内容"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "日報",
    "内容": joinText(row["現場"], row["作業内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "備考": row["備考"] || "",
    "元シート": "日報"
  }), rows);


  collectRowsAsObjects(ss, "免許資格管理", row => row["更新期限"] || row["資格名"], row => ({
    "日付": row["更新期限"] || new Date(),
    "種類": "免許資格",
    "内容": joinText(row["担当"], joinText(row["資格名"], row["区分"])),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["更新期限"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "備考": row["備考"] || "",
    "元シート": "免許資格管理"
  }), rows);

  collectRowsAsObjects(ss, "備品修理管理", row => row["返却予定日"] || row["備品名"], row => ({
    "日付": row["返却予定日"] || row["登録日"] || new Date(),
    "種類": "備品修理",
    "内容": joinText(row["備品名"], joinText(row["場所"], row["内容"])),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["返却予定日"] || row["修理依頼日"] || row["登録日"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "備考": row["備考"] || "",
    "元シート": "備品修理管理"
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

  getPersonalMembers_().forEach(name => {
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

  const headers = getSheetHeaders_()["要確認一覧"];

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
    .filter(row => shouldIncludeInAlertList_(row, priority))
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
      if (row["重要度"] === "高" && !isNoticeReadByAll_(row)) {
        rows.unshift({
          "日付": row["投稿日"] || new Date(),
          "種類": "お知らせ",
          "内容": joinText(row["タイトル"], row["内容"]),
          "担当": row["投稿者"],
          "状態": "要確認",
          "通知": "重要",
          "電話対応": "",
          "備考": row["備考"] || "",
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
    "備考": row["備考"] || "",
    "元シート": row["元シート"]
  })));

  applyColorRules(sheet);
  sheet.autoResizeColumns(1, headers.length);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    sheet.hideColumns(calendarIdCol);
  }
}


/**
 * 要確認一覧に残すかどうかを判定する。
 *
 * 要確認一覧は「見たもの」ではなく「まだ対応が必要なもの」を残す場所です。
 * お知らせは既読で消してよいですが、電話・備品・車検・免許などは
 * 既読だけでは消さず、完了・返却済・更新済などの状態で消します。
 */
function shouldIncludeInAlertList_(row, priority) {
  const sourceName = row["元シート"] || row["種類"] || "";
  const status = String(row["状態"] || "").trim();
  const notice = String(row["通知"] || "").trim();
  const phone = String(row["電話対応"] || "").trim();

  if (sourceName === "お知らせ" || row["種類"] === "お知らせ") {
    return false;
  }

  if (!priority[notice]) return false;

  if (sourceName === "電話履歴" || row["種類"] === "電話履歴") {
    return phone !== "完了" && status !== "完了";
  }

  if (sourceName === "備品修理管理" || row["種類"] === "備品修理") {
    return status !== "返却済" && status !== "完了" && status !== "中止";
  }

  if (sourceName === "免許資格管理" || row["種類"] === "免許資格") {
    return status !== "有効" && status !== "更新済";
  }

  if (sourceName === "車検管理" || row["種類"] === "車検") {
    return status !== "完了" && status !== "更新済";
  }

  if (sourceName === "社用車予約" || row["種類"] === "社用車予約") {
    return status !== "返却済" && status !== "完了" && status !== "中止";
  }

  if (sourceName === "個人ToDo" || row["種類"] === "ToDo") {
    return status !== "完了" && status !== "中止";
  }

  return status !== "完了" && status !== "確認済" && status !== "更新済" && status !== "返却済" && status !== "中止";
}

function isNoticeReadByAll_(row) {
  const members = getPersonalMembers_();
  if (members.length === 0) return false;

  return members.every(name => {
    const value = row[name];
    return value === true || value === "TRUE" || value === "Y";
  });
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("ダッシュボード");
  if (!sheet) {
    sheet = ss.insertSheet("ダッシュボード");
  }

  // ダッシュボードはA〜D列を使うため、列数が足りない場合は追加する
  if (sheet.getMaxColumns() < 4) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 4 - sheet.getMaxColumns());
  }

  // 既存フィルタ・結合セルがあると再作成時にエラーになるため解除する
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {}

  try {
    sheet.showColumns(1, sheet.getMaxColumns());
  } catch (e) {}

  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (e) {}

  sheet.clear();
  sheet.clearFormats();

  // Sheetオブジェクトには clearDataValidations() がないため、Rangeに対して実行する
  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .clearDataValidations();

  sheet.setConditionalFormatRules([]);

  sheet.getRange("A1").setValue("社内共有ダッシュボード");
  sheet.getRange("A1:D1")
    .setBackground("#d9ead3")
    .setFontWeight("bold");

  sheet.getRange("A1")
    .setFontSize(22)
    .setFontWeight("bold");

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
    ["日報件数", '=COUNTA(日報!A2:A)'],
    ["免許資格件数", '=COUNTA(免許資格管理!A2:A)'],
    ["備品修理件数", '=COUNTA(備品修理管理!A2:A)'],
    ["担当別未読あり", '=COUNTIF(担当別未読!B2:B,">0")']
  ];

  sheet.getRange(4, 1, data.length, 2).setValues(data);

  sheet.getRange(3, 1, data.length + 1, 2)
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(4, 2, data.length, 1)
    .setHorizontalAlignment("center");

  const resizeCols = Math.min(4, sheet.getMaxColumns());
  sheet.autoResizeColumns(1, resizeCols);

  sheet.setFrozenRows(3);
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の更新処理中のため、全体更新をスキップしました");
    return;
  }

  try {
    // 通常更新は軽量版です。
    // 行グループ再作成や過去一覧の月別再編は重いので、毎回は実行しません。
    refreshInputSheets();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
  } finally {
    lock.releaseLock();
  }
}

function refreshAllHeavy() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の更新処理中のため、重い全体更新をスキップしました");
    return;
  }

  try {
    refreshInputSheets();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    resetColumnGroups();
    createCheckGroup();
    groupArchiveByMonthWithHeader();
  } finally {
    lock.releaseLock();
  }
}

function refreshSummaryOnly_() {
  createScheduleList();
  createAlertList();
  createAssigneeUnreadSummary();
  createReadRateSummary();
  createDashboard();
}

function getInputSheetNames_() {
  return [
    "出先予定",
    "車検管理",
    "車検履歴",
    "会議予定",
    "行事予定",
    "作業状況",
    "工事予定",
    "電話履歴",
    "お知らせ",
    "個人ToDo",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理"
  ];
}

function setCheckboxesForDataRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1 || lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const checkNames = [READ_HEADER].concat(getPersonalMembers_());

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

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return;

  try {
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

    if (sheetName === "社用車予約") {
      const conflictMessage = getCarReservationConflictMessage(sheet, row);
      if (conflictMessage) {
        SpreadsheetApp.getUi().alert(conflictMessage);
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const noticeCol = headers.indexOf("通知") + 1;
        if (noticeCol > 0) sheet.getRange(row, noticeCol).setValue("予約重複");
      }
    }

    // 入力用シートを編集した時だけ、一覧・要確認・集計を更新します。
    // 行グループ作成や過去一覧整理は重いので onEdit では実行しません。
    if (getInputSheetNames_().includes(sheetName)) {
      refreshSummaryOnly_();
    }
  } finally {
    lock.releaseLock();
  }
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
  const priority = ["日付", "日時", "開始日", "期限", "投稿日", "登録日", "車検期限", "保険期限", "終了日"];

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

  const checkNames = [READ_HEADER].concat(getPersonalMembers_());

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
    "車検履歴",
    "会議予定",
    "行事予定",
    "作業状況",
    "工事予定",
    "電話履歴",
    "一覧スケジュール",
    "お知らせ",
    "個人ToDo",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理",
    "過去一覧"
  ];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addItem("シート構成を作成（軽量）", "createCompanySheets")
    .addItem("設定シートを作成/確認", "setupSettingsSheet")
    .addItem("設定を各シートに反映", "applySettingsToSheets")
    .addSeparator()
    .addItem("全体更新（軽量）", "refreshAll")
    .addItem("全体更新（重い：グループ/過去一覧も再作成）", "refreshAllHeavy")
    .addItem("要確認一覧を更新", "createAlertList")
    .addItem("担当別未読を更新", "createAssigneeUnreadSummary")
    .addItem("既読率集計を更新", "createReadRateSummary")
    .addItem("Googleカレンダーへ反映", "syncCalendarEvents")
    .addItem("社用車予約の重複チェック", "checkCarReservationConflicts")
    .addItem("車検更新完了処理", "processCompletedVehicleInspections")
    .addItem("車検管理を並び替え", "sortVehicleInspectionSheet")
    .addItem("現在のシートを月ごとに折りたたみ", "groupRowsByMonth")
    .addItem("全シートを月ごとに折りたたみ", "groupAllRowsByMonth")
    .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
    .addItem("月次メンテナンス", "monthlyMaintenance")
    .addItem("裏方シートを非表示", "hideSupportSheets")
    .addSeparator()
    .addItem("選択行の日報文章をGPT作成", "createDailyReportText")
    .addItem("選択行の日報PDFを作成", "createSelectedDailyReportPdf")
    .addItem("当日サマリーPDFを作成", "createDailyReportPdf")
    .addItem("一覧帳簿PDFを作成", "createLedgerPdf")
    .addItem("スプレッドシートをバックアップ", "backupSpreadsheet")
    .addItem("毎日自動更新トリガーを設定", "installDailyRefreshTrigger")
    .addItem("毎月過去一覧移動トリガーを設定", "installMonthlyArchiveTrigger")
    .addItem("毎月メンテナンストリガーを設定", "installMonthlyMaintenanceTrigger")
    .addItem("自動化トリガーをまとめて設定", "installAllAutomationTriggers")
    .addSeparator()
    .addItem("SQL設計シートを作成", "createSqlDesignSheet")
    .addItem("移行対応表を作成", "createSqlMigrationMappingSheet")
    .addItem("SQLサンプル集を作成", "createSqlSampleQueries")
    .addItem("SQL関連シートをまとめて作成", "createSqlSupportSheets")
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

    const start = headers.indexOf(getPersonalMembers_()[0]) + 1;
    const end = headers.indexOf(getPersonalMembers_()[getPersonalMembers_().length - 1]) + 1;

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
  addTextRule(statusCol, "使用中", "#cfe2f3");
  addTextRule(statusCol, "返却待ち", "#fff2cc");
  addTextRule(statusCol, "返却済", "#b6d7a8");
  addTextRule(statusCol, "提出済", "#d9ead3");
  addTextRule(statusCol, "確認済", "#b6d7a8");
  addTextRule(statusCol, "差戻し", "#f4cccc");
  addTextRule(statusCol, "更新予定", "#fff2cc");
  addTextRule(statusCol, "更新済", "#b6d7a8");
  addTextRule(statusCol, "有効", "#b6d7a8");
  addTextRule(statusCol, "期限切れ", "#ea9999");
  addTextRule(statusCol, "失効", "#ea9999");
  addTextRule(statusCol, "予約済", "#d9ead3");
  addTextRule(statusCol, "実施済", "#cfe2f3");
  addTextRule(statusCol, "未依頼", "#f4cccc");
  addTextRule(statusCol, "依頼済", "#fff2cc");
  addTextRule(statusCol, "修理中", "#cfe2f3");

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

  const headers = getSheetHeaders_()["担当別未読"];
  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  const schedule = ss.getSheetByName("一覧スケジュール");
  const todo = ss.getSheetByName("個人ToDo");
  const phone = ss.getSheetByName("電話履歴");

  const rows = getPersonalMembers_().map(checkName => {
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
      startDate: row => row["開始日"] || row["日付"],
      endDate: row => row["終了日"] || row["開始日"] || row["日付"],
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


function groupAllRowsByMonth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targetSheetNames = [
    "出先予定",
    "車検管理",
    "車検履歴",
    "会議予定",
    "行事予定",
    "工事予定",
    "電話履歴",
    "一覧スケジュール",
    "個人ToDo",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理"
  ];

  targetSheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    groupRowsByMonthForSheet(sheet);
  });
}

function groupRowsByMonthForSheet(sheet) {
  if (!sheet || sheet.getLastRow() < 3) return;

  resetRowGroups(sheet);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dateCol = findMainDateColumn(headers);

  if (dateCol <= 0) return;

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();

  let currentKey = "";
  let groupStart = 2;

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const key = getYearMonthKey(values[i][0]);

    if (!key) continue;

    if (!currentKey) {
      currentKey = key;
      groupStart = rowNumber;
      continue;
    }

    if (key !== currentKey) {
      const groupEnd = rowNumber - 1;

      if (groupEnd > groupStart) {
        sheet.getRange(groupStart, 1, groupEnd - groupStart + 1, 1)
          .shiftRowGroupDepth(1);
      }

      currentKey = key;
      groupStart = rowNumber;
    }
  }

  if (currentKey && lastRow > groupStart) {
    sheet.getRange(groupStart, 1, lastRow - groupStart + 1, 1)
      .shiftRowGroupDepth(1);
  }
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

  resetSheet(archiveSheet, getSheetHeaders_()["過去一覧"].length);
  setupSheet(archiveSheet, getSheetHeaders_()["過去一覧"]);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 車検管理は過去一覧へ移動しません。車検履歴で管理します。
  // 免許資格管理も原則、現在有効な資格管理表として残します。
  const targetSheetNames = [
    "出先予定",
    "会議予定",
    "行事予定",
    "工事予定",
    "電話履歴",
    "社用車予約",
    "個人ToDo",
    "日報",
    "備品修理管理"
  ];

  const archiveRows = [];

  targetSheetNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    const rowsToDelete = [];

    values.forEach((valuesRow, index) => {
      const rowNumber = index + 2;
      const row = objectFromRow(headers, valuesRow);

      if (shouldMoveToArchive_(row, sheetName, today)) {
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

/**
 * 過去一覧へ移動してよいかを判定する。
 *
 * 日付が過ぎただけでは移動しません。
 * 未返却備品、未対応電話、未完了ToDo、未提出日報などは残します。
 */
function shouldMoveToArchive_(row, sheetName, today) {
  const dateValue = getArchiveDateValue_(row, sheetName);
  if (!dateValue) return false;

  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return false;
  date.setHours(0, 0, 0, 0);

  if (date.getTime() >= today.getTime()) return false;

  const status = String(row["状態"] || "").trim();
  const phone = String(row["電話対応"] || "").trim();
  const returned = String(row["返却済"] || "").trim();

  if (sheetName === "出先予定") return ["完了", "中止"].includes(status);
  if (sheetName === "会議予定") return ["完了", "中止"].includes(status);
  if (sheetName === "行事予定") return ["完了", "中止"].includes(status);
  if (sheetName === "工事予定") return ["完了", "中止"].includes(status);

  if (sheetName === "電話履歴") {
    return phone === "完了" || status === "完了";
  }

  if (sheetName === "社用車予約") {
    return ["返却済", "完了", "中止"].includes(status);
  }

  if (sheetName === "個人ToDo") {
    return ["完了", "中止"].includes(status);
  }

  if (sheetName === "日報") {
    return ["提出済", "確認済", "完了"].includes(status);
  }

  if (sheetName === "備品修理管理") {
    return returned === "済" || ["返却済", "完了", "中止"].includes(status);
  }

  return false;
}

function getArchiveDateValue_(row, sheetName) {
  if (sheetName === "行事予定") return row["終了日"] || row["開始日"] || row["日付"];
  if (sheetName === "工事予定") return row["終了日"] || row["開始日"];
  if (sheetName === "電話履歴") return row["日時"];
  if (sheetName === "社用車予約") return row["日付"];
  if (sheetName === "個人ToDo") return row["期限"] || row["登録日"];
  if (sheetName === "日報") return row["日付"];
  if (sheetName === "備品修理管理") return row["返却予定日"] || row["修理依頼日"] || row["登録日"];
  return row["日付"] || row["開始日"] || row["登録日"];
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
    dateValue = row["開始日"] || row["日付"];
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
  } else if (sheetName === "日報") {
    dateValue = row["日付"];
    content = joinText(row["現場"], row["作業内容"]);
    staff = row["担当"];
    note = row["日報文章"] || row["備考"] || "";
  } else if (sheetName === "免許資格管理") {
    dateValue = row["更新期限"];
    content = joinText(row["資格名"], row["区分"]);
    staff = row["担当"];
    note = joinText("コピー有無: " + (row["コピー有無"] || ""), row["備考"] || "");
  } else if (sheetName === "備品修理管理") {
    dateValue = row["返却予定日"] || row["修理依頼日"] || row["登録日"];
    content = joinText(row["備品名"], joinText(row["場所"], row["内容"]));
    staff = row["担当"];
    note = joinText("返却済: " + (row["返却済"] || ""), row["備考"] || "");
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


/**
 * 月次メンテナンス
 *
 * 月1回だけ実行する重めの整理処理。
 * 要確認一覧は日次・編集時に更新し、過去一覧整理はこの関数で月次実行する。
 */
function monthlyMaintenance() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の更新処理中のため、月次メンテナンスをスキップしました");
    return;
  }

  try {
    movePastItemsToArchive();
    groupArchiveByMonthWithHeader();
    createDashboard();
    hideSupportSheets();
    SpreadsheetApp.getActiveSpreadsheet().toast("月次メンテナンスが完了しました");
  } finally {
    lock.releaseLock();
  }
}

/**
 * 裏方シートを非表示にする。
 * 利用者が触るシートを減らして、運用画面を見やすくする。
 */
function hideSupportSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeName = activeSheet ? activeSheet.getName() : "";

  const supportSheets = [
    "設定",
    "担当別未読",
    "既読率集計",
    "ダッシュボード",
    "過去一覧",
    "車検履歴",
    "帳簿PDF履歴",
    "PDF_日報_一時",
    "帳簿PDF_一時"
  ];

  supportSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    if (sheet.getName() === activeName) return;

    try {
      sheet.hideSheet();
    } catch (e) {}
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("裏方シートを非表示にしました");
}

/**
 * 月次メンテナンストリガーを設定する。
 */
function installMonthlyMaintenanceTrigger() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "monthlyMaintenance") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("monthlyMaintenance")
    .timeBased()
    .onMonthDay(1)
    .atHour(6)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("毎月1日 6時の月次メンテナンストリガーを設定しました");
}


function SHEET_ID(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return "";
  return sheet.getSheetId();
}



// サンプルデータ投入機能は社用運用では不要のため削除しています。
// 初回は createCompanySheets() を実行し、その後は実データを入力してください。

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
  // 要確認一覧・一覧スケジュール・未読/既読集計は毎日更新します。
  // 過去一覧への移動は月次トリガー movePastItemsToArchive で実行します。
  refreshAll();
  checkCarReservationConflicts();
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

/**
 * 既読率集計シートの入力規則エラー修正
 *
 * 既読率集計は自動集計シートなので、2行目以降の入力規則を解除します。
 * A列「担当」に通常の担当者プルダウンが残っている場合のエラー対策です。
 */
function fixReadRateSummaryValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("既読率集計");

  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("既読率集計シートが見つかりません");
    return;
  }

  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  if (maxRows < 2 || maxCols < 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast("既読率集計シートに解除対象がありません");
    return;
  }

  sheet.getRange(2, 1, maxRows - 1, maxCols).clearDataValidations();
  SpreadsheetApp.getActiveSpreadsheet().toast("既読率集計の入力規則を解除しました");
}

/**
 * 既読率集計を作成
 *
 * 設定シートの「既読確認者」をもとに、
 * 一覧スケジュール上の個人既読列を集計します。
 *
 * このシートは自動集計用なので、担当列のプルダウンは不要です。
 * setupSheet() / writeObjectsToSheet() により入力規則が付く場合があるため、
 * 最後に clearDataValidations() で入力規則を解除します。
 */
function createReadRateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("既読率集計");
  if (!sheet) {
    sheet = ss.insertSheet("既読率集計");
  }

  const headers = ["担当", "対象件数", "既読件数", "未読件数", "既読率"];

  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  const scheduleSheet = ss.getSheetByName("一覧スケジュール");
  const personalMembers = getPersonalMembers_();
  const rows = [];

  personalMembers.forEach(member => {
    let total = 0;
    let read = 0;

    if (scheduleSheet && scheduleSheet.getLastRow() >= 2) {
      const scheduleHeaders = scheduleSheet
        .getRange(1, 1, 1, scheduleSheet.getLastColumn())
        .getValues()[0];

      const memberCol = scheduleHeaders.indexOf(member) + 1;

      if (memberCol > 0) {
        const values = scheduleSheet
          .getRange(2, memberCol, scheduleSheet.getLastRow() - 1, 1)
          .getValues();

        values.forEach(row => {
          total++;

          if (row[0] === true || row[0] === "TRUE" || row[0] === "Y") {
            read++;
          }
        });
      }
    }

    const unread = total - read;
    const rate = total === 0 ? 0 : read / total;

    rows.push({
      "担当": member,
      "対象件数": total,
      "既読件数": read,
      "未読件数": unread,
      "既読率": rate
    });
  });

  writeObjectsToSheet(sheet, rows);

  const rateCol = headers.indexOf("既読率") + 1;
  if (rateCol > 0 && rows.length > 0) {
    sheet.getRange(2, rateCol, rows.length, 1).setNumberFormat("0.0%");
  }

  // 既読率集計は自動集計シートなので、2行目以降の入力規則を解除します。
  if (sheet.getMaxRows() >= 2) {
    sheet
      .getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns())
      .clearDataValidations();
  }

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


/****************************************************************
 * 追加機能 v8.0
 * - 日報シート
 * - GPT日報文章作成
 * - 選択行の日報PDF作成
 ****************************************************************/

function createDailyReportText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== "日報") {
    SpreadsheetApp.getUi().alert("日報シートで、文章を作成したい行を選択してください。");
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("見出し行ではなく、日報データの行を選択してください。");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const data = objectFromRow(headers, values);

  const reportCol = headers.indexOf("日報文章") + 1;
  if (reportCol <= 0) {
    SpreadsheetApp.getUi().alert("日報文章列が見つかりません。");
    return;
  }

  const prompt = buildDailyReportPrompt_(data);
  const text = callOpenAI(prompt);

  if (!text) {
    SpreadsheetApp.getUi().alert("日報文章を作成できませんでした。APIキーまたは通信状況を確認してください。");
    return;
  }

  sheet.getRange(row, reportCol).setValue(text);
  SpreadsheetApp.getActiveSpreadsheet().toast("選択行の日報文章を作成しました");
}

function buildDailyReportPrompt_(data) {
  return [
    "以下の内容から、社内提出用の日報文章を自然な日本語で作成してください。",
    "",
    "【入力内容】",
    "日付: " + formatDateForReport_(data["日付"]),
    "担当: " + (data["担当"] || ""),
    "現場: " + (data["現場"] || ""),
    "作業内容: " + (data["作業内容"] || ""),
    "進捗: " + (data["進捗"] || ""),
    "問題点: " + (data["問題点"] || ""),
    "明日の予定: " + (data["明日の予定"] || ""),
    "備考: " + (data["備考"] || ""),
    "",
    "【条件】",
    "・丁寧な社内向け文章にする",
    "・事実を勝手に追加しない",
    "・長すぎず、読みやすくする",
    "・問題点が空欄の場合は、問題点について無理に書かない",
    "・明日の予定が空欄の場合は、明日の予定について無理に書かない"
  ].join("\n");
}

function callOpenAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");

  if (!apiKey) {
    SpreadsheetApp.getUi().alert("OPENAI_API_KEY が未設定です。スクリプトプロパティにAPIキーを保存してください。");
    return "";
  }

  const payload = {
    model: "gpt-4.1-mini",
    input: prompt
  };

  const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + apiKey
    },
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

  if (json.output && json.output.length > 0) {
    const parts = [];
    json.output.forEach(item => {
      if (!item.content) return;
      item.content.forEach(content => {
        if (content.text) parts.push(content.text);
      });
    });
    return parts.join("\n").trim();
  }

  return "";
}

function createSelectedDailyReportPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== "日報") {
    SpreadsheetApp.getUi().alert("日報シートで、PDF化したい行を選択してください。");
    return;
  }

  const row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert("見出し行ではなく、日報データの行を選択してください。");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const data = objectFromRow(headers, values);

  const dateText = formatDateForReport_(data["日付"]) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const fileDateText = dateText.replace(/\D/g, "") || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const staffText = data["担当"] || "担当未設定";
  const siteText = data["現場"] || "現場未設定";

  const doc = DocumentApp.create("日報_" + fileDateText + "_" + staffText + "_" + siteText);
  const body = doc.getBody();

  body.appendParagraph("日報").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("日付: " + dateText);
  body.appendParagraph("担当: " + staffText);
  body.appendParagraph("現場: " + siteText);
  body.appendParagraph("");

  appendDailyReportParagraph_(body, "作業内容", data["作業内容"]);
  appendDailyReportParagraph_(body, "進捗", data["進捗"]);
  appendDailyReportParagraph_(body, "問題点", data["問題点"]);
  appendDailyReportParagraph_(body, "明日の予定", data["明日の予定"]);
  appendDailyReportParagraph_(body, "日報文章", data["日報文章"]);
  appendDailyReportParagraph_(body, "写真", data["写真"]);
  appendDailyReportParagraph_(body, "状態", data["状態"]);
  appendDailyReportParagraph_(body, "備考", data["備考"]);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const safeStaff = String(staffText).replace(/[\\/:*?"<>|]/g, "_");
  const safeSite = String(siteText).replace(/[\\/:*?"<>|]/g, "_");
  const pdfBlob = docFile.getBlob().getAs(MimeType.PDF).setName("日報_" + fileDateText + "_" + safeStaff + "_" + safeSite + ".pdf");

  const pdfFolder = getOrCreateChildFolder_("日報PDF");
  const pdfFile = pdfFolder.createFile(pdfBlob);

  docFile.setTrashed(true);

  SpreadsheetApp.getActiveSpreadsheet().toast("選択行の日報PDFを作成しました");
  return pdfFile.getUrl();
}

function appendDailyReportParagraph_(body, title, value) {
  body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(value ? String(value) : "");
  body.appendParagraph("");
}

function getOrCreateChildFolder_(folderName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parents = spreadsheetFile.getParents();

  let parentFolder = null;
  if (parents.hasNext()) {
    parentFolder = parents.next();
  }

  if (!parentFolder) {
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(folderName);
  }

  const childFolders = parentFolder.getFoldersByName(folderName);
  if (childFolders.hasNext()) return childFolders.next();

  return parentFolder.createFolder(folderName);
}

function setOpenAIKeyForSetup() {
  // 初回設定用です。
  // 1. 下の YOUR_OPENAI_API_KEY_HERE を実際のAPIキーに一時的に置き換える
  // 2. この関数を1回だけ実行する
  // 3. 実行後、APIキー文字列を必ず消して保存する
  PropertiesService.getScriptProperties().setProperty(
    "OPENAI_API_KEY",
    "YOUR_OPENAI_API_KEY_HERE"
  );
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
    "お知らせ",
    "日報"
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



// 重複していた function copyPersonalChecks は整理版では削除しました。

function testCreateDiarySheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("日報");

  if (!sheet) {
    sheet = ss.insertSheet("日報");
  }

  const headers = [
    "日付",
    "担当",
    "現場",
    "作業内容",
    "進捗",
    "問題点",
    "明日の予定",
    "写真",
    "日報文章",
    "状態",
    "備考",
    "既読",
    "安倍",
    "三好",
    "佐藤"
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}






















function rebuildDailyReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("日報");

  if (!sheet) {
    sheet = ss.insertSheet("日報");
  }

  const headers = [
    "日付",
    "担当",
    "現場",
    "作業内容",
    "進捗",
    "問題点",
    "明日の予定",
    "写真",
    "日報文章",
    "状態",
    "備考",
    "既読",
    "安倍",
    "三好",
    "佐藤"
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  const lastCol = headers.length;
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.getRange(2, 1, 500, 1).setNumberFormat("yyyy/mm/dd");

  sheet.getRange(2, 12, 500, 4).insertCheckboxes();

  sheet.autoResizeColumns(1, lastCol);
}







function setupDailyReportSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const dateCol = headers.indexOf("日付") + 1;
  const staffCol = headers.indexOf("担当") + 1;
  const statusCol = headers.indexOf("状態") + 1;

  if (dateCol > 0) {
    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, dateCol, 500, 1)
      .setNumberFormat("yyyy/mm/dd")
      .setDataValidation(dateRule);
  }

  if (staffCol > 0) {
    const staffRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, ...getStaffMembers_()], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, staffCol, 500, 1)
      .setDataValidation(staffRule);
  }

  if (statusCol > 0) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, "予定", "進行中", "施工中", "完了", "延期", "中止", "要確認"], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, statusCol, 500, 1)
      .setDataValidation(statusRule);
  }

  sheet.autoResizeColumns(1, sheet.getLastColumn());
}





function fixCarNumberColumnAsText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const carNumberCol = headers.indexOf("車番") + 1;
  if (carNumberCol <= 0) return;

  sheet.getRange(2, carNumberCol, 500, 1)
    .setNumberFormat("@");
}


function fixDailyReportCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const checkNames = ["既読", "安倍", "三好", "佐藤"];

  checkNames.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;

    const dataRange = sheet.getRange(2, 1, lastRow - 1, col - 1);
    const values = dataRange.getValues();

    for (let i = 0; i < values.length; i++) {
      const rowNumber = i + 2;
      const hasData = values[i].some(v => v !== "" && v !== null);

      const cell = sheet.getRange(rowNumber, col);

      if (hasData) {
        cell.insertCheckboxes();
        if (cell.getValue() === "") cell.setValue(false);
      } else {
        cell.clearDataValidations();
        cell.clearContent();
      }
    }
  });
}




// 重複していた function createSelectedDailyReportPdf は整理版では削除しました。

function cleanDailyReportEmptyRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  const maxRows = sheet.getMaxRows();
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = sheet.getLastColumn();

  // 3行目以降の空行の入力規則・チェックボックスを消す
  if (maxRows > 2) {
    sheet.getRange(3, 1, maxRows - 2, lastCol).clearDataValidations();
    sheet.getRange(3, 1, maxRows - 2, lastCol).clearContent();
  }

  // 2行目だけ入力規則を戻す
  setupDailyReportSheetOnly();

  // 行グループを解除
  resetRowGroups(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast("日報の空行を整理しました");
}





















/****************************************************************
 * 追加機能 v8.1
 * - 免許資格管理
 * - 備品修理管理
 * - 他部門共有用 一覧帳簿PDF
 ****************************************************************/

function createLedgerPdf() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  refreshAll();

  const docTitle = "社内共有管理_一覧帳簿_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  body.appendParagraph("社内共有管理 一覧帳簿").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("作成日時: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy年M月d日 H:mm"));
  body.appendParagraph("用途: 他部門共有・会議配布・紙提出用");
  body.appendParagraph("");

  appendLedgerSection_(body, ss, "要確認一覧", ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"], 50);
  appendLedgerSection_(body, ss, "一覧スケジュール", ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート"], 100);
  appendLedgerSection_(body, ss, "免許資格管理", ["担当", "資格名", "区分", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"], 100);
  appendLedgerSection_(body, ss, "備品修理管理", ["登録日", "備品名", "場所", "内容", "担当", "修理依頼日", "返却予定日", "返却済", "状態", "通知", "備考"], 100);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getBlob().getAs(MimeType.PDF).setName(docTitle + ".pdf");
  const folder = getOrCreateChildFolder_("一覧帳簿PDF");
  const pdfFile = folder.createFile(pdfBlob);
  docFile.setTrashed(true);

  recordLedgerPdfHistory_(pdfFile.getUrl(), "一覧スケジュール・要確認一覧・免許資格管理・備品修理管理", "作成日時基準");
  SpreadsheetApp.getUi().alert("一覧帳簿PDFを作成しました。\n\n" + pdfFile.getUrl());
  return pdfFile.getUrl();
}

function appendLedgerSection_(body, ss, sheetName, outputHeaders, maxRows) {
  const sheet = ss.getSheetByName(sheetName);
  body.appendParagraph(sheetName).setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (!sheet || sheet.getLastRow() < 2) {
    body.appendParagraph("対象データはありません。");
    body.appendParagraph("");
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const rows = values
    .map(valuesRow => objectFromRow(headers, valuesRow))
    .filter(row => {
      if (row["種類"] === "月見出し") return false;
      return outputHeaders.some(header => row[header] !== "" && row[header] !== null && row[header] !== undefined);
    })
    .slice(0, maxRows || 100);

  if (rows.length === 0) {
    body.appendParagraph("対象データはありません。");
    body.appendParagraph("");
    return;
  }

  const tableValues = [];
  tableValues.push(outputHeaders);

  rows.forEach(row => {
    tableValues.push(outputHeaders.map(header => formatLedgerValue_(row[header])));
  });

  const table = body.appendTable(tableValues);
  table.setBorderWidth(0.5);

  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).setBackgroundColor("#d9ead3");
    headerRow.getCell(i).editAsText().setBold(true);
  }

  body.appendParagraph("");
}

function formatLedgerValue_(value) {
  if (value === null || value === undefined) return "";

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }

  if (value === true) return "済";
  if (value === false) return "未";

  return String(value);
}

function recordLedgerPdfHistory_(url, target, periodText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("帳簿PDF履歴");

  if (!sheet) {
    sheet = ss.insertSheet("帳簿PDF履歴");
    setupSheet(sheet, getSheetHeaders_()["帳簿PDF履歴"]);
  }

  const row = {
    "作成日時": new Date(),
    "対象": target || "",
    "期間": periodText || "",
    "PDFリンク": url || "",
    "備考": "他部門共有用"
  };

  appendObjectsToSheet(sheet, [row]);
}

function setupLicenseAndRepairSheetsOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headersMap = getSheetHeaders_();

  ["免許資格管理", "備品修理管理", "帳簿PDF履歴"].forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    resetSheet(sheet, headersMap[name].length);
    setupSheet(sheet, headersMap[name]);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("免許資格管理・備品修理管理・帳簿PDF履歴を作成しました");
}




/**
 * 既存の車検管理シートに「次回車検期限」列を安全に追加する。
 * 既存データが右にずれないよう、車検期限の右に列を挿入する。
 */
function ensureVehicleInspectionNextColumn_(sheet) {
  if (!sheet || sheet.getName() !== "車検管理") return;
  if (sheet.getLastColumn() < 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nextCol = headers.indexOf("次回車検期限") + 1;
  if (nextCol > 0) return;

  const inspectionCol = headers.indexOf("車検期限") + 1;
  if (inspectionCol <= 0) return;

  sheet.insertColumnAfter(inspectionCol);
  sheet.getRange(1, inspectionCol + 1).setValue("次回車検期限");
}

/**
 * 車検管理シートと車検履歴シートを作成・調整する。
 * 既存の車検管理データは消さず、「次回車検期限」列を追加する。
 */
function setupVehicleInspectionHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headersMap = getSheetHeaders_();

  let vehicleSheet = ss.getSheetByName("車検管理");
  if (!vehicleSheet) vehicleSheet = ss.insertSheet("車検管理");
  ensureVehicleInspectionNextColumn_(vehicleSheet);
  setupSheetWithoutClearing_(vehicleSheet, headersMap["車検管理"]);

  let historySheet = ss.getSheetByName("車検履歴");
  if (!historySheet) historySheet = ss.insertSheet("車検履歴");
  setupSheetWithoutClearing_(historySheet, headersMap["車検履歴"]);

  SpreadsheetApp.getActiveSpreadsheet().toast("車検管理・車検履歴を確認しました");
}

/**
 * 車検更新完了処理
 *
 * 車検管理で状態を「完了」または「更新済」にした行について、
 * 次回車検期限を車検期限へ移し、車検履歴シートへ履歴を残す。
 *
 * 車検管理 = 現在有効な次回期限だけ管理
 * 車検履歴 = 過去の車検更新履歴を保存
 * 過去一覧 = 車検は基本的に移動しない
 */
function processCompletedVehicleInspections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");

  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getActiveSpreadsheet().toast("車検管理に処理対象がありません");
    return;
  }

  setupVehicleInspectionHistory();

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const carNameCol = headers.indexOf("車両名") + 1;
  const carNumberCol = headers.indexOf("車番") + 1;
  const inspectionCol = headers.indexOf("車検期限") + 1;
  const nextInspectionCol = headers.indexOf("次回車検期限") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const insuranceCol = headers.indexOf("保険期限") + 1;
  const statusCol = headers.indexOf("状態") + 1;
  const staffCol = headers.indexOf("担当") + 1;
  const noteCol = headers.indexOf("備考") + 1;

  if (carNameCol <= 0 || inspectionCol <= 0 || nextInspectionCol <= 0 || statusCol <= 0) {
    SpreadsheetApp.getUi().alert(
      "車検管理に必要な列が不足しています。\n\n必要列：車両名、車検期限、次回車検期限、状態"
    );
    return;
  }

  let historySheet = ss.getSheetByName("車検履歴");
  if (!historySheet) historySheet = ss.insertSheet("車検履歴");
  setupSheetWithoutClearing_(historySheet, getSheetHeaders_()["車検履歴"]);

  const historyHeaders = historySheet.getRange(1, 1, 1, historySheet.getLastColumn()).getValues()[0];
  const historyRows = [];

  const lastRow = sheet.getLastRow();
  let processedCount = 0;

  for (let row = 2; row <= lastRow; row++) {
    const status = sheet.getRange(row, statusCol).getValue();
    if (status !== "完了" && status !== "更新済") continue;

    const oldInspectionDate = sheet.getRange(row, inspectionCol).getValue();
    const nextInspectionDate = sheet.getRange(row, nextInspectionCol).getValue();

    if (!nextInspectionDate) {
      // 次回期限が空欄の場合は事故防止のため処理しない
      continue;
    }

    const rowObj = {
      "更新日": new Date(),
      "車両名": carNameCol > 0 ? sheet.getRange(row, carNameCol).getValue() : "",
      "車番": carNumberCol > 0 ? sheet.getRange(row, carNumberCol).getValue() : "",
      "旧車検期限": oldInspectionDate,
      "新車検期限": nextInspectionDate,
      "保険期限": insuranceCol > 0 ? sheet.getRange(row, insuranceCol).getValue() : "",
      "担当": staffCol > 0 ? sheet.getRange(row, staffCol).getValue() : "",
      "備考": noteCol > 0 ? sheet.getRange(row, noteCol).getValue() : ""
    };

    historyRows.push(historyHeaders.map(header => rowObj[header] !== undefined ? rowObj[header] : ""));

    // 次回期限を現在の車検期限へ反映
    sheet.getRange(row, inspectionCol).setValue(nextInspectionDate);

    // 次回車検期限は空欄へ戻す
    sheet.getRange(row, nextInspectionCol).clearContent();

    // 状態を空欄へ戻す
    sheet.getRange(row, statusCol).clearContent();

    // 通知を再計算
    if (noticeCol > 0) {
      sheet.getRange(row, noticeCol).setValue(getNoticeText(nextInspectionDate));
    }

    processedCount++;
  }

  if (historyRows.length > 0) {
    const startRow = historySheet.getLastRow() + 1;
    historySheet.getRange(startRow, 1, historyRows.length, historyHeaders.length).setValues(historyRows);
    historySheet.autoResizeColumns(1, historyHeaders.length);
  }

  sortVehicleInspectionSheet();

  // 一覧・要確認などへ反映
  try {
    refreshSummaryOnly_();
  } catch (e) {}

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "車検更新完了処理：" + processedCount + "件"
  );
}

/**
 * 車検管理シートを車検期限順に並び替える。
 * 状態が完了・更新済の行は下へ送る。
 */
function sortVehicleInspectionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");
  if (!sheet || sheet.getLastRow() < 3) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = headers.indexOf("状態") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const inspectionCol = headers.indexOf("車検期限") + 1;

  if (inspectionCol <= 0) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const noticePriority = {
    "期限切れ": 1,
    "今日": 2,
    "3日以内": 3,
    "7日以内": 4,
    "": 9
  };

  values.sort((a, b) => {
    const statusA = statusCol > 0 ? a[statusCol - 1] : "";
    const statusB = statusCol > 0 ? b[statusCol - 1] : "";

    const doneA = statusA === "完了" || statusA === "更新済";
    const doneB = statusB === "完了" || statusB === "更新済";

    if (doneA !== doneB) {
      return doneA ? 1 : -1;
    }

    if (noticeCol > 0) {
      const noticeA = a[noticeCol - 1] || "";
      const noticeB = b[noticeCol - 1] || "";

      const pa = noticePriority[noticeA] || 9;
      const pb = noticePriority[noticeB] || 9;

      if (pa !== pb) return pa - pb;
    }

    const dateA = new Date(a[inspectionCol - 1]);
    const dateB = new Date(b[inspectionCol - 1]);

    const timeA = isNaN(dateA.getTime()) ? 9999999999999 : dateA.getTime();
    const timeB = isNaN(dateB.getTime()) ? 9999999999999 : dateB.getTime();

    return timeA - timeB;
  });

  sheet.getRange(2, 1, values.length, lastCol).setValues(values);
  setCheckboxesForDataRows(sheet);
  applyColorRules(sheet);
  sheet.autoResizeColumns(1, lastCol);

  SpreadsheetApp.getActiveSpreadsheet().toast("車検管理を並び替えました");
}


/**
 * SQL対応シートをまとめて作成する。
 *
 * 外部DBには接続しない安全版。
 * Google Sheets版の業務管理システムを、将来的にRDBへ移行できるように
 * SQL設計・移行対応表・SQLサンプル集として整理する。
 */
function createSqlSupportSheets() {
  createSqlDesignSheet();
  createSqlMigrationMappingSheet();
  createSqlSampleQueries();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "SQL設計・移行対応表・SQLサンプル集を作成しました"
  );
}

/**
 * SQL設計シートを作成する。
 *
 * 各業務シートをRDBテーブルとして再設計した表。
 * ポートフォリオでは「Google Sheets版をRDB化する設計」として見せる。
 */
function createSqlDesignSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "SQL設計";
  const sheet = recreateSqlSheet_(ss, sheetName);

  const headers = [
    "テーブル名",
    "元シート",
    "カラム名",
    "型",
    "主キー",
    "外部キー",
    "NULL可否",
    "説明"
  ];

  const rows = [
    ["staff", "設定", "id", "INTEGER", "PK", "", "NOT NULL", "担当者ID"],
    ["staff", "設定", "name", "TEXT", "", "", "NOT NULL", "担当者名"],
    ["staff", "設定", "is_active", "BOOLEAN", "", "", "NOT NULL", "利用中かどうか"],

    ["read_members", "設定", "id", "INTEGER", "PK", "", "NOT NULL", "既読確認者ID"],
    ["read_members", "設定", "staff_id", "INTEGER", "", "staff.id", "NOT NULL", "担当者への参照"],
    ["read_members", "設定", "display_order", "INTEGER", "", "", "NULL", "既読列の表示順"],

    ["schedules", "一覧スケジュール", "id", "INTEGER", "PK", "", "NOT NULL", "予定ID"],
    ["schedules", "一覧スケジュール", "schedule_date", "DATE", "", "", "NOT NULL", "予定日・期限日"],
    ["schedules", "一覧スケジュール", "type", "TEXT", "", "", "NOT NULL", "種類"],
    ["schedules", "一覧スケジュール", "content", "TEXT", "", "", "NOT NULL", "内容"],
    ["schedules", "一覧スケジュール", "staff_id", "INTEGER", "", "staff.id", "NULL", "担当者ID"],
    ["schedules", "一覧スケジュール", "status", "TEXT", "", "", "NULL", "状態"],
    ["schedules", "一覧スケジュール", "notice", "TEXT", "", "", "NULL", "通知"],
    ["schedules", "一覧スケジュール", "phone_status", "TEXT", "", "", "NULL", "電話対応"],
    ["schedules", "一覧スケジュール", "remarks", "TEXT", "", "", "NULL", "備考"],
    ["schedules", "一覧スケジュール", "source_sheet", "TEXT", "", "", "NOT NULL", "元シート"],

    ["read_checks", "一覧スケジュール", "id", "INTEGER", "PK", "", "NOT NULL", "既読確認ID"],
    ["read_checks", "一覧スケジュール", "schedule_id", "INTEGER", "", "schedules.id", "NOT NULL", "予定ID"],
    ["read_checks", "一覧スケジュール", "staff_id", "INTEGER", "", "staff.id", "NOT NULL", "既読確認者ID"],
    ["read_checks", "一覧スケジュール", "is_read", "BOOLEAN", "", "", "NOT NULL", "既読済みか"],
    ["read_checks", "一覧スケジュール", "read_at", "DATETIME", "", "", "NULL", "既読日時"],

    ["vehicles", "車検管理", "id", "INTEGER", "PK", "", "NOT NULL", "車両ID"],
    ["vehicles", "車検管理", "vehicle_name", "TEXT", "", "", "NOT NULL", "車両名"],
    ["vehicles", "車検管理", "vehicle_number", "TEXT", "", "", "NULL", "車番"],
    ["vehicles", "車検管理", "inspection_due_date", "DATE", "", "", "NOT NULL", "車検期限"],
    ["vehicles", "車検管理", "next_inspection_due_date", "DATE", "", "", "NULL", "次回車検期限"],
    ["vehicles", "車検管理", "insurance_due_date", "DATE", "", "", "NULL", "保険期限"],
    ["vehicles", "車検管理", "status", "TEXT", "", "", "NULL", "状態"],
    ["vehicles", "車検管理", "photo_url", "TEXT", "", "", "NULL", "写真URL"],
    ["vehicles", "車検管理", "remarks", "TEXT", "", "", "NULL", "備考"],

    ["vehicle_inspection_logs", "車検履歴", "id", "INTEGER", "PK", "", "NOT NULL", "車検履歴ID"],
    ["vehicle_inspection_logs", "車検履歴", "vehicle_id", "INTEGER", "", "vehicles.id", "NOT NULL", "車両ID"],
    ["vehicle_inspection_logs", "車検履歴", "old_inspection_due_date", "DATE", "", "", "NOT NULL", "旧車検期限"],
    ["vehicle_inspection_logs", "車検履歴", "new_inspection_due_date", "DATE", "", "", "NOT NULL", "新車検期限"],
    ["vehicle_inspection_logs", "車検履歴", "processed_at", "DATETIME", "", "", "NOT NULL", "更新日時"],
    ["vehicle_inspection_logs", "車検履歴", "processed_by", "TEXT", "", "", "NULL", "処理担当"],
    ["vehicle_inspection_logs", "車検履歴", "remarks", "TEXT", "", "", "NULL", "備考"],

    ["daily_reports", "日報", "id", "INTEGER", "PK", "", "NOT NULL", "日報ID"],
    ["daily_reports", "日報", "report_date", "DATE", "", "", "NOT NULL", "日付"],
    ["daily_reports", "日報", "staff_id", "INTEGER", "", "staff.id", "NOT NULL", "担当者ID"],
    ["daily_reports", "日報", "site_name", "TEXT", "", "", "NULL", "現場"],
    ["daily_reports", "日報", "work_content", "TEXT", "", "", "NULL", "作業内容"],
    ["daily_reports", "日報", "progress", "TEXT", "", "", "NULL", "進捗"],
    ["daily_reports", "日報", "issue", "TEXT", "", "", "NULL", "問題点"],
    ["daily_reports", "日報", "next_plan", "TEXT", "", "", "NULL", "明日の予定"],
    ["daily_reports", "日報", "photo_url", "TEXT", "", "", "NULL", "写真URL"],
    ["daily_reports", "日報", "report_text", "TEXT", "", "", "NULL", "日報文章"],
    ["daily_reports", "日報", "status", "TEXT", "", "", "NULL", "状態"],
    ["daily_reports", "日報", "pdf_url", "TEXT", "", "", "NULL", "PDFリンク"],
    ["daily_reports", "日報", "remarks", "TEXT", "", "", "NULL", "備考"],

    ["equipment_repairs", "備品修理管理", "id", "INTEGER", "PK", "", "NOT NULL", "備品修理ID"],
    ["equipment_repairs", "備品修理管理", "registered_date", "DATE", "", "", "NULL", "登録日"],
    ["equipment_repairs", "備品修理管理", "equipment_name", "TEXT", "", "", "NOT NULL", "備品名"],
    ["equipment_repairs", "備品修理管理", "location", "TEXT", "", "", "NULL", "場所"],
    ["equipment_repairs", "備品修理管理", "content", "TEXT", "", "", "NULL", "内容"],
    ["equipment_repairs", "備品修理管理", "staff_id", "INTEGER", "", "staff.id", "NULL", "担当者ID"],
    ["equipment_repairs", "備品修理管理", "repair_requested_date", "DATE", "", "", "NULL", "修理依頼日"],
    ["equipment_repairs", "備品修理管理", "return_due_date", "DATE", "", "", "NULL", "返却予定日"],
    ["equipment_repairs", "備品修理管理", "returned_status", "TEXT", "", "", "NULL", "返却済"],
    ["equipment_repairs", "備品修理管理", "status", "TEXT", "", "", "NULL", "状態"],
    ["equipment_repairs", "備品修理管理", "notice", "TEXT", "", "", "NULL", "通知"],
    ["equipment_repairs", "備品修理管理", "remarks", "TEXT", "", "", "NULL", "備考"],

    ["licenses", "免許資格管理", "id", "INTEGER", "PK", "", "NOT NULL", "免許資格ID"],
    ["licenses", "免許資格管理", "staff_id", "INTEGER", "", "staff.id", "NOT NULL", "担当者ID"],
    ["licenses", "免許資格管理", "license_name", "TEXT", "", "", "NOT NULL", "資格名"],
    ["licenses", "免許資格管理", "category", "TEXT", "", "", "NULL", "区分"],
    ["licenses", "免許資格管理", "acquired_date", "DATE", "", "", "NULL", "取得日"],
    ["licenses", "免許資格管理", "renewal_due_date", "DATE", "", "", "NULL", "更新期限"],
    ["licenses", "免許資格管理", "copy_status", "TEXT", "", "", "NULL", "コピー有無"],
    ["licenses", "免許資格管理", "status", "TEXT", "", "", "NULL", "状態"],
    ["licenses", "免許資格管理", "notice", "TEXT", "", "", "NULL", "通知"],
    ["licenses", "免許資格管理", "remarks", "TEXT", "", "", "NULL", "備考"],

    ["phone_calls", "電話履歴", "id", "INTEGER", "PK", "", "NOT NULL", "電話履歴ID"],
    ["phone_calls", "電話履歴", "called_at", "DATETIME", "", "", "NOT NULL", "日時"],
    ["phone_calls", "電話履歴", "partner_name", "TEXT", "", "", "NULL", "相手"],
    ["phone_calls", "電話履歴", "content", "TEXT", "", "", "NULL", "内容"],
    ["phone_calls", "電話履歴", "phone_status", "TEXT", "", "", "NULL", "電話対応"],
    ["phone_calls", "電話履歴", "staff_id", "INTEGER", "", "staff.id", "NULL", "担当者ID"],
    ["phone_calls", "電話履歴", "status", "TEXT", "", "", "NULL", "状態"],
    ["phone_calls", "電話履歴", "notice", "TEXT", "", "", "NULL", "通知"],
    ["phone_calls", "電話履歴", "remarks", "TEXT", "", "", "NULL", "備考"]
  ];

  writeSqlSheet_(sheet, headers, rows);
  SpreadsheetApp.getActiveSpreadsheet().toast("SQL設計シートを作成しました");
}

/**
 * 移行対応表を作成する。
 *
 * Google Sheetsの列をSQLテーブルの列にどう変換するかを整理する。
 * 横持ちの個人既読列を read_checks に縦持ち化する説明が重要。
 */
function createSqlMigrationMappingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "移行対応表";
  const sheet = recreateSqlSheet_(ss, sheetName);

  const headers = [
    "元シート",
    "元列",
    "SQLテーブル",
    "SQLカラム",
    "変換ルール",
    "備考"
  ];

  const rows = [
    ["設定", "担当者", "staff", "name", "設定種別='担当者' の値を登録", "担当プルダウンのマスタ"],
    ["設定", "既読確認者", "read_members", "staff_id", "staffと紐づける", "個人既読の対象者"],

    ["一覧スケジュール", "日付", "schedules", "schedule_date", "DATE型へ変換", "予定日または期限日"],
    ["一覧スケジュール", "種類", "schedules", "type", "TEXT", "出先予定・車検・日報など"],
    ["一覧スケジュール", "内容", "schedules", "content", "TEXT", "一覧表示用の内容"],
    ["一覧スケジュール", "担当", "schedules", "staff_id", "staff.nameからstaff.idへ変換", "未設定の場合はNULL"],
    ["一覧スケジュール", "状態", "schedules", "status", "TEXT", "業務状態"],
    ["一覧スケジュール", "通知", "schedules", "notice", "TEXT", "期限切れ・今日・3日以内など"],
    ["一覧スケジュール", "電話対応", "schedules", "phone_status", "TEXT", "未対応・対応中・折返し・完了"],
    ["一覧スケジュール", "備考", "schedules", "remarks", "TEXT", "備考"],
    ["一覧スケジュール", "元シート", "schedules", "source_sheet", "TEXT", "元データのシート名"],

    ["一覧スケジュール", "既読確認者列", "read_checks", "is_read", "横持ち列を縦持ちに変換", "例：安倍/三好/佐藤列を行データにする"],
    ["一覧スケジュール", "既読確認者列名", "read_checks", "staff_id", "列名をstaffに紐づける", "DBでは列追加ではなく行追加で管理"],

    ["車検管理", "車両名", "vehicles", "vehicle_name", "TEXT", "車両マスタとして扱う"],
    ["車検管理", "車番", "vehicles", "vehicle_number", "TEXT", "数値化しない"],
    ["車検管理", "車検期限", "vehicles", "inspection_due_date", "DATE型へ変換", "現在の車検期限"],
    ["車検管理", "次回車検期限", "vehicles", "next_inspection_due_date", "DATE型へ変換", "更新予定日"],
    ["車検管理", "保険期限", "vehicles", "insurance_due_date", "DATE型へ変換", "保険期限"],
    ["車検管理", "状態", "vehicles", "status", "TEXT", "未対応・予約済・実施済・完了・更新済"],
    ["車検管理", "写真", "vehicles", "photo_url", "URL文字列", "写真リンク"],
    ["車検管理", "備考", "vehicles", "remarks", "TEXT", "備考"],

    ["車検履歴", "旧車検期限", "vehicle_inspection_logs", "old_inspection_due_date", "DATE型へ変換", "更新前の期限"],
    ["車検履歴", "新車検期限", "vehicle_inspection_logs", "new_inspection_due_date", "DATE型へ変換", "更新後の期限"],
    ["車検履歴", "更新日", "vehicle_inspection_logs", "processed_at", "DATETIME型へ変換", "処理日時"],

    ["日報", "日付", "daily_reports", "report_date", "DATE型へ変換", "日報日付"],
    ["日報", "担当", "daily_reports", "staff_id", "staff.nameからstaff.idへ変換", "担当者"],
    ["日報", "現場", "daily_reports", "site_name", "TEXT", "現場名"],
    ["日報", "作業内容", "daily_reports", "work_content", "TEXT", "作業内容"],
    ["日報", "進捗", "daily_reports", "progress", "TEXT", "進捗"],
    ["日報", "問題点", "daily_reports", "issue", "TEXT", "問題点"],
    ["日報", "明日の予定", "daily_reports", "next_plan", "TEXT", "翌日の予定"],
    ["日報", "写真", "daily_reports", "photo_url", "URL文字列", "写真リンク"],
    ["日報", "日報文章", "daily_reports", "report_text", "TEXT", "GPT生成文章"],
    ["日報", "状態", "daily_reports", "status", "TEXT", "下書き・提出済・確認済など"],
    ["日報", "PDFリンク", "daily_reports", "pdf_url", "URL文字列", "PDF出力先"],

    ["備品修理管理", "登録日", "equipment_repairs", "registered_date", "DATE型へ変換", "登録日"],
    ["備品修理管理", "備品名", "equipment_repairs", "equipment_name", "TEXT", "備品名"],
    ["備品修理管理", "場所", "equipment_repairs", "location", "TEXT", "保管場所・発生場所"],
    ["備品修理管理", "内容", "equipment_repairs", "content", "TEXT", "修理内容"],
    ["備品修理管理", "担当", "equipment_repairs", "staff_id", "staff.nameからstaff.idへ変換", "担当者"],
    ["備品修理管理", "修理依頼日", "equipment_repairs", "repair_requested_date", "DATE型へ変換", "依頼日"],
    ["備品修理管理", "返却予定日", "equipment_repairs", "return_due_date", "DATE型へ変換", "返却予定日"],
    ["備品修理管理", "返却済", "equipment_repairs", "returned_status", "済/未のTEXT", "未返却判定に使用"],
    ["備品修理管理", "状態", "equipment_repairs", "status", "TEXT", "返却済・完了・中止で完了扱い"],
    ["備品修理管理", "通知", "equipment_repairs", "notice", "TEXT", "期限通知"],
    ["備品修理管理", "備考", "equipment_repairs", "remarks", "TEXT", "備考"],

    ["免許資格管理", "担当", "licenses", "staff_id", "staff.nameからstaff.idへ変換", "保有者"],
    ["免許資格管理", "資格名", "licenses", "license_name", "TEXT", "資格名"],
    ["免許資格管理", "区分", "licenses", "category", "TEXT", "免許・資格など"],
    ["免許資格管理", "取得日", "licenses", "acquired_date", "DATE型へ変換", "取得日"],
    ["免許資格管理", "更新期限", "licenses", "renewal_due_date", "DATE型へ変換", "更新期限"],
    ["免許資格管理", "コピー有無", "licenses", "copy_status", "TEXT", "有・無・未確認"],
    ["免許資格管理", "状態", "licenses", "status", "TEXT", "有効・更新予定・更新済・期限切れなど"],

    ["電話履歴", "日時", "phone_calls", "called_at", "DATETIME型へ変換", "電話日時"],
    ["電話履歴", "相手", "phone_calls", "partner_name", "TEXT", "相手先"],
    ["電話履歴", "内容", "phone_calls", "content", "TEXT", "内容"],
    ["電話履歴", "電話対応", "phone_calls", "phone_status", "TEXT", "未対応・対応中・折返し・完了"],
    ["電話履歴", "担当", "phone_calls", "staff_id", "staff.nameからstaff.idへ変換", "担当者"]
  ];

  writeSqlSheet_(sheet, headers, rows);
  SpreadsheetApp.getActiveSpreadsheet().toast("移行対応表を作成しました");
}

/**
 * SQLサンプル集を作成する。
 *
 * 業務で使うSELECT / WHERE / ORDER BY / GROUP BY / JOINを説明する。
 * SQLite想定のSQLにしている。
 */
function createSqlSampleQueries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "SQLサンプル集";
  const sheet = recreateSqlSheet_(ss, sheetName);

  const headers = [
    "目的",
    "使用テーブル",
    "SQL",
    "説明"
  ];

  const rows = [
    [
      "今日の予定を抽出",
      "schedules",
      "SELECT *\nFROM schedules\nWHERE schedule_date = DATE('now')\nORDER BY schedule_date;",
      "一覧スケジュールのうち、今日分だけを抽出する。"
    ],
    [
      "7日以内の予定・期限を抽出",
      "schedules",
      "SELECT *\nFROM schedules\nWHERE schedule_date BETWEEN DATE('now') AND DATE('now', '+7 day')\nORDER BY schedule_date;",
      "今日から7日以内の予定・期限を抽出する。"
    ],
    [
      "期限切れを抽出",
      "schedules",
      "SELECT *\nFROM schedules\nWHERE schedule_date < DATE('now')\n  AND status NOT IN ('完了', '確認済', '更新済', '返却済', '中止')\nORDER BY schedule_date;",
      "完了していない期限切れデータを抽出する。"
    ],
    [
      "未対応電話を抽出",
      "phone_calls",
      "SELECT *\nFROM phone_calls\nWHERE phone_status IN ('未対応', '折返し')\n   OR status IN ('未対応', '対応中')\nORDER BY called_at;",
      "折返し忘れ・未対応電話を確認する。"
    ],
    [
      "未返却備品を抽出",
      "equipment_repairs",
      "SELECT *\nFROM equipment_repairs\nWHERE COALESCE(returned_status, '') <> '済'\n  AND status NOT IN ('返却済', '完了', '中止')\nORDER BY return_due_date;",
      "未返却の備品修理データを抽出する。過去一覧へ移動しない対象。"
    ],
    [
      "7日以内の車検期限を抽出",
      "vehicles",
      "SELECT *\nFROM vehicles\nWHERE inspection_due_date BETWEEN DATE('now') AND DATE('now', '+7 day')\n  AND status NOT IN ('完了', '更新済')\nORDER BY inspection_due_date;",
      "車検期限が近い車両を抽出する。"
    ],
    [
      "期限切れの免許資格を抽出",
      "licenses",
      "SELECT l.*, s.name AS staff_name\nFROM licenses l\nJOIN staff s ON l.staff_id = s.id\nWHERE l.renewal_due_date < DATE('now')\n  AND l.status NOT IN ('有効', '更新済')\nORDER BY l.renewal_due_date;",
      "担当者名をJOINして、期限切れ資格を確認する。"
    ],
    [
      "担当者別の未読件数を集計",
      "read_checks / staff",
      "SELECT\n  s.name AS staff_name,\n  COUNT(*) AS unread_count\nFROM read_checks rc\nJOIN staff s ON rc.staff_id = s.id\nWHERE rc.is_read = 0\nGROUP BY s.name\nORDER BY unread_count DESC;",
      "個人既読列を縦持ち化したread_checksから、担当者別未読件数を集計する。"
    ],
    [
      "担当者別の日報件数を集計",
      "daily_reports / staff",
      "SELECT\n  s.name AS staff_name,\n  COUNT(*) AS report_count\nFROM daily_reports dr\nJOIN staff s ON dr.staff_id = s.id\nGROUP BY s.name\nORDER BY report_count DESC;",
      "JOINとGROUP BYで担当者別の日報件数を出す。"
    ],
    [
      "車両ごとの車検履歴を表示",
      "vehicles / vehicle_inspection_logs",
      "SELECT\n  v.vehicle_name,\n  v.vehicle_number,\n  log.old_inspection_due_date,\n  log.new_inspection_due_date,\n  log.processed_at\nFROM vehicle_inspection_logs log\nJOIN vehicles v ON log.vehicle_id = v.id\nORDER BY v.vehicle_name, log.processed_at DESC;",
      "車検管理と車検履歴をJOINして、車両ごとの履歴を確認する。"
    ],
    [
      "要確認一覧相当をSQLで抽出",
      "schedules",
      "SELECT *\nFROM schedules\nWHERE notice IN ('期限切れ', '今日', '3日以内', '7日以内')\n  AND status NOT IN ('完了', '確認済', '更新済', '返却済', '中止')\nORDER BY\n  CASE notice\n    WHEN '期限切れ' THEN 1\n    WHEN '今日' THEN 2\n    WHEN '3日以内' THEN 3\n    WHEN '7日以内' THEN 4\n    ELSE 99\n  END,\n  schedule_date;",
      "Apps Scriptの要確認一覧に近い条件をSQLで再現する。"
    ],
    [
      "横持ち既読を縦持ちにする考え方",
      "read_checks",
      "/*\nSheets:\n  schedule_id / 内容 / 安倍 / 三好 / 佐藤\n\nRDB:\n  schedule_id / staff_id / is_read\n\n例:\nINSERT INTO read_checks (schedule_id, staff_id, is_read)\nVALUES\n  (1, 1, 1),\n  (1, 2, 0),\n  (1, 3, 1);\n*/",
      "スプレッドシートでは横持ち、SQLでは縦持ちにして管理する。"
    ]
  ];

  writeSqlSheet_(sheet, headers, rows);
  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 1, 220);
  sheet.setColumnWidths(3, 1, 620);
  sheet.setColumnWidths(4, 1, 360);

  SpreadsheetApp.getActiveSpreadsheet().toast("SQLサンプル集を作成しました");
}

/**
 * SQL関連シートを作り直す。
 * 既存の同名シートだけをクリアするので、業務データには触らない。
 */
function recreateSqlSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    const filter = sheet.getFilter();
    if (filter) filter.remove();

    sheet.clear();
    sheet.clearFormats();
    sheet
      .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .clearDataValidations();
    sheet.setConditionalFormatRules([]);

    try {
      sheet.showColumns(1, sheet.getMaxColumns());
    } catch (e) {}
  }

  return sheet;
}

/**
 * SQL関連シートの共通書き込み処理。
 */
function writeSqlSheet_(sheet, headers, rows) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);

  const lastRow = Math.max(rows.length + 1, 2);
  sheet.getRange(1, 1, lastRow, headers.length)
    .setVerticalAlignment("top")
    .setWrap(true);

  try {
    sheet.getRange(1, 1, lastRow, headers.length).createFilter();
  } catch (e) {}

  sheet.autoResizeColumns(1, headers.length);
}



/**
 * 入力用・確認用シートの表示幅をまとめて調整する。
 *
 * 長文が入りやすい列を広げ、折り返し表示をONにする。
 * 入力のたびに自動変更するのではなく、メニューから必要時に実行する安全版。
 */
function formatInputSheetsForLongText() {
  formatDailyReportSheet();
  formatWorkStatusSheet_();
  formatPhoneHistorySheet_();
  formatEquipmentRepairSheet_();
  formatNoticeSheet_();
  formatTodoSheet_();
  formatScheduleListSheet_();
  formatAlertListSheet_();
  formatLicenseSheet_();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "入力用・確認用シートの表示幅をまとめて調整しました"
  );
}

/**
 * 日報シートの表示幅を調整する。
 *
 * 日報文章・作業内容・問題点・明日の予定などの長文列を広げる。
 */
function formatDailyReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "日付", 110);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "現場", 180);
  setColumnWidthByHeader_(sheet, headers, "作業内容", 260);
  setColumnWidthByHeader_(sheet, headers, "進捗", 180);
  setColumnWidthByHeader_(sheet, headers, "問題点", 260);
  setColumnWidthByHeader_(sheet, headers, "明日の予定", 260);
  setColumnWidthByHeader_(sheet, headers, "写真", 180);
  setColumnWidthByHeader_(sheet, headers, "日報文章", 600);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "備考", 240);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 90);
}

/**
 * 作業状況シートの表示幅を調整する。
 */
function formatWorkStatusSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("作業状況");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "現場", 180);
  setColumnWidthByHeader_(sheet, headers, "作業内容", 340);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "写真", 180);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 280);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 80);
}

/**
 * 電話履歴シートの表示幅を調整する。
 */
function formatPhoneHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("電話履歴");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "日時", 150);
  setColumnWidthByHeader_(sheet, headers, "相手", 180);
  setColumnWidthByHeader_(sheet, headers, "内容", 420);
  setColumnWidthByHeader_(sheet, headers, "電話対応", 130);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 280);
  setColumnWidthByHeader_(sheet, headers, "メモ", 280);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 80);
}

/**
 * 備品修理管理シートの表示幅を調整する。
 */
function formatEquipmentRepairSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("備品修理管理");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "登録日", 120);
  setColumnWidthByHeader_(sheet, headers, "備品名", 180);
  setColumnWidthByHeader_(sheet, headers, "場所", 180);
  setColumnWidthByHeader_(sheet, headers, "内容", 360);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "修理依頼日", 130);
  setColumnWidthByHeader_(sheet, headers, "返却予定日", 130);
  setColumnWidthByHeader_(sheet, headers, "返却済", 100);
  setColumnWidthByHeader_(sheet, headers, "状態", 130);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 300);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 80);
}

/**
 * お知らせシートの表示幅を調整する。
 */
function formatNoticeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("お知らせ");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "投稿日", 120);
  setColumnWidthByHeader_(sheet, headers, "タイトル", 240);
  setColumnWidthByHeader_(sheet, headers, "内容", 520);
  setColumnWidthByHeader_(sheet, headers, "投稿者", 120);
  setColumnWidthByHeader_(sheet, headers, "重要度", 100);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 260);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 80);
}

/**
 * 個人ToDoシートの表示幅を調整する。
 */
function formatTodoSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("個人ToDo");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "登録日", 120);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "内容", 420);
  setColumnWidthByHeader_(sheet, headers, "期限", 120);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 280);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 80);
}

/**
 * 一覧スケジュールの表示幅を調整する。
 */
function formatScheduleListSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("一覧スケジュール");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "日付", 120);
  setColumnWidthByHeader_(sheet, headers, "種類", 130);
  setColumnWidthByHeader_(sheet, headers, "内容", 420);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "電話対応", 130);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 280);
  setColumnWidthByHeader_(sheet, headers, "元シート", 140);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 70);
}

/**
 * 要確認一覧の表示幅を調整する。
 */
function formatAlertListSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("要確認一覧");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "日付", 120);
  setColumnWidthByHeader_(sheet, headers, "種類", 130);
  setColumnWidthByHeader_(sheet, headers, "内容", 460);
  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "電話対応", 130);
  setColumnWidthByHeader_(sheet, headers, "備考", 300);
  setColumnWidthByHeader_(sheet, headers, "元シート", 140);

  applyReadableLongTextFormat_(sheet, 70);
}

/**
 * 免許資格管理シートの表示幅を調整する。
 */
function formatLicenseSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("免許資格管理");
  if (!sheet) return;

  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  setColumnWidthByHeader_(sheet, headers, "担当", 120);
  setColumnWidthByHeader_(sheet, headers, "資格名", 240);
  setColumnWidthByHeader_(sheet, headers, "区分", 120);
  setColumnWidthByHeader_(sheet, headers, "取得日", 120);
  setColumnWidthByHeader_(sheet, headers, "更新期限", 130);
  setColumnWidthByHeader_(sheet, headers, "コピー有無", 120);
  setColumnWidthByHeader_(sheet, headers, "状態", 120);
  setColumnWidthByHeader_(sheet, headers, "通知", 120);
  setColumnWidthByHeader_(sheet, headers, "既読", 80);
  setColumnWidthByHeader_(sheet, headers, "備考", 280);

  getPersonalMembers_().forEach(name => {
    setColumnWidthByHeader_(sheet, headers, name, 90);
  });

  applyReadableLongTextFormat_(sheet, 70);
}

/**
 * シートのヘッダーを取得する。
 */
function getHeadersForFormat_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * 指定したヘッダー名の列幅を変更する。
 */
function setColumnWidthByHeader_(sheet, headers, headerName, width) {
  const col = headers.indexOf(headerName) + 1;
  if (col > 0) {
    sheet.setColumnWidth(col, width);
  }
}

/**
 * 長文セルを見やすくする共通処理。
 */
function applyReadableLongTextFormat_(sheet, bodyRowHeight) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const targetRows = Math.max(sheet.getMaxRows(), 50);

  sheet.getRange(1, 1, targetRows, lastCol)
    .setWrap(true)
    .setVerticalAlignment("top");

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);

  const maxRowsToFormat = Math.min(targetRows, 100);
  for (let r = 2; r <= maxRowsToFormat; r++) {
    sheet.setRowHeight(r, bodyRowHeight);
  }
}　　　　　


































