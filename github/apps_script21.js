/**
 * 社内共有業務管理システム v9.6.14 安定版（社用説明文調整版）
 *
 * 目的:
 * - v9.6.6〜v9.6.13の運用安定版をベースに、社用運用版として説明文を整理。
 * - 設定シートの内部文言を整理。
 * - 個人確認列の縦書き、空欄に戻すプルダウン運用、フィルタ安全化を維持。
 * - 公開・社内引継ぎの両方で読めるようにコメントを整理。
 * - 社内管理説明シートを追加し、概要・機能・運用ルール・トラブル対応をシート内で確認できるようにする。
 * - 一覧スケジュールの個人確認列縦書き、シート順、設定/既読率集計の表示幅を調整。
 * - 社内利用時に不自然な社外説明向け文言を社用向けに整理。
 * - 日報レシートは管理・要確認までとし、集計用の追加シートは安定優先で追加しない。
 * - 日報レシート管理シートを追加し、紙レシート・精算状態・経理確認を管理。
 * - 日報レシートは一覧スケジュールへは入れず、未確認・差戻し・未精算だけ要確認一覧へ出す。
 *
 * 注意:
 * - サンプルデータは架空名・サンプル名です。
 * - 実運用時は「設定」シートの担当者・既読確認者を実名に変更してください。
 * - APIキーなどの秘密情報はコードに直書きせず、Script Propertiesで管理してください。
 */


/**
 * 社用メモ
 * - 担当者・既読確認者は「設定」シートから読み込みます。
 * - 設定シートが未作成の場合は DEFAULT_STAFF_MEMBERS / DEFAULT_PERSONAL_MEMBERS を使います。
 * - 社外共有・説明用では架空名を使用します。実運用時は「設定」シートの値を変更してください。
 */

const DEFAULT_STAFF_MEMBERS = [
  "山田",
  "鈴木",
  "田中",
  "高橋",
  "伊藤",
  "中村",
  "小林",
  "吉田",
  "山本",
  "佐々木",
  "松本",
  "井上",
  "木村",
  "林",
  "清水",
];

const DEFAULT_PERSONAL_MEMBERS = [
  "山田",
  "高橋",
  "鈴木"
];

const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";
const CALENDAR_ID_HEADER = "カレンダーID";

/**
 * AppSheet用の安定キー列。
 *
 * 既存データがある本番シートでも列ずれしにくいように、ID列は各シートの末尾に置きます。
 * AppSheet側では各ID列を Key にし、Initial value は UNIQUEID() を推奨します。
 */
const SHEET_ID_HEADERS = {
  "出先予定": "出先ID",
  "工事予定": "工事ID",
  "会議予定": "会議ID",
  "行事予定": "行事ID",
  "作業状況": "作業ID",
  "電話履歴": "電話ID",
  "車検管理": "車両ID",
  "車検履歴": "履歴ID",
  "社用車予約": "予約ID",
  "日報": "日報ID",
  "免許資格管理": "資格ID",
  "備品修理管理": "修理ID",
  "お知らせ": "お知らせID",
  "個人ToDo": "ToDo_ID",
  "日報レシート管理": "レシートID",
  "帳簿PDF履歴": "帳簿ID"
};

// 入力規則・日付形式を適用する最低行数。
// 500行固定だと本番運用で500行以降の入力規則が外れるため、最低1000行まで広げます。
const MIN_DATA_ROWS_FOR_VALIDATION = 300;

// 入力が最終行付近まで到達した時に追加する行数。
const AUTO_EXTEND_ROWS_BUFFER = 100;

// 最終行から何行以内に入力されたら自動で行を追加するか。
const AUTO_EXTEND_TRIGGER_MARGIN = 30;


const DATE_HEADERS = [
  "日付",
  "日時",
  "作成日時",
  "更新日",
  "移動日",
  "開始日",
  "終了日",
  "車検期限",
  "次回車検期限",
  "保険期限",
  "次回保険期限",
  "投稿日",
  "登録日",
  "期限",
  "取得日",
  "更新期限",
  "修理依頼日",
  "返却予定日"
];


function getApplyRowCount_(sheet) {
  if (!sheet) return MIN_DATA_ROWS_FOR_VALIDATION;

  const maxRows = sheet.getMaxRows();
  const lastRow = Math.max(sheet.getLastRow(), 1);

  // 全行に入力規則を毎回かけると重くなるため、
  // 「最低300行」または「最終行 + 余裕100行」までに抑える。
  const targetRows = Math.min(
    maxRows,
    Math.max(MIN_DATA_ROWS_FOR_VALIDATION + 1, lastRow + AUTO_EXTEND_ROWS_BUFFER)
  );

  return Math.max(targetRows - 1, 1);
}

function ensureSheetMinimumRows_(sheet) {
  if (!sheet) return;

  const targetRows = MIN_DATA_ROWS_FOR_VALIDATION + 1; // 1行目はヘッダー
  const maxRows = sheet.getMaxRows();

  if (maxRows < targetRows) {
    sheet.insertRowsAfter(maxRows, targetRows - maxRows);
  }
}

function ensureRowsAfterEdit_(sheet, row) {
  if (!sheet || row <= 1) return;

  const maxRows = sheet.getMaxRows();
  const remainingRows = maxRows - row;

  if (remainingRows > AUTO_EXTEND_TRIGGER_MARGIN) return;

  sheet.insertRowsAfter(maxRows, AUTO_EXTEND_ROWS_BUFFER);

  const headersMap = getSheetHeaders_();
  const sheetName = sheet.getName();

  if (headersMap[sheetName]) {
    setupSheetWithoutClearing_(sheet, headersMap[sheetName]);
  } else {
    applyColorRules(sheet);
  }
}

function getSheetHeaders_() {
  const personalMembers = getPersonalMembers_();

  return {
  "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "出先ID"],
  "車検管理": ["車両名", "車番", "車検期限", "次回車検期限", "通知", "保険期限", "次回保険期限", "状態", "写真", READ_HEADER, ...personalMembers, "備考", "車両ID"],
  "車検履歴": ["更新日", "車両名", "車番", "旧車検期限", "新車検期限", "旧保険期限", "新保険期限", "担当", "備考", "履歴ID"],
  "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "会議ID"],
  "行事予定": ["開始日", "終了日", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "行事ID"],
  "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...personalMembers, "備考", "作業ID"],
  "工事予定": ["工事名", "現場", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "工事ID"],
  "電話履歴": ["日時", "相手", "内容", "電話対応", "担当", "対応メモ", "通知", READ_HEADER, ...personalMembers, "備考", "電話ID"],
  "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", "元シート"],
  "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "備考", "元シート"],
  "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...personalMembers, "備考", "お知らせID"],
  "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "ToDo_ID"],
  "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "予約ID"],
  "日報": ["日付", "入力者", "担当", "現場", "作業内容", "進捗", "問題点", "明日の予定", "他現場状況", "写真", "日報文章", "状態", "備考", READ_HEADER, ...personalMembers, "PDFリンク", "日報ID"],
  "日報レシート管理": ["日付", "担当", "現場", "支払先", "内容", "区分", "金額", "支払方法", "レシート写真", "経理確認", "精算状態", "通知", "備考", "レシートID"],
  "免許資格管理": ["担当", "資格名", "区分", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "資格ID"],
  "備品修理管理": ["登録日", "備品名", "場所", "内容", "担当", "修理依頼日", "返却予定日", "返却済", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "修理ID"],
  "帳簿PDF履歴": ["作成日時", "対象", "期間", "PDFリンク", "備考", "帳簿ID"],
  "帳簿出力設定": ["出力する", "帳簿名", "シート名", "出力列", "最大行数", "備考"],
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
      rows.push(["担当者", name, "担当プルダウン用。実運用時はこの値を変更してください。"]);
    });

    DEFAULT_PERSONAL_MEMBERS.forEach(name => {
      rows.push(["既読確認者", name, "個人確認列用。実運用時は確認者名を変更してください。"]);
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  normalizeSettingsSheetNotes_(sheet);
  return sheet;
}

function normalizeSettingsSheetNotes_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  if (!sheet || sheet.getLastRow() < 2) return;

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  let changed = false;

  const updated = values.map(row => {
    const type = String(row[0] || "").trim();
    const currentNote = String(row[2] || "");
    let note = currentNote;

    if (type === "担当者") {
      note = "担当プルダウン用。実運用時はこの値を変更してください。";
    } else if (type === "既読確認者") {
      note = "個人確認列用。実運用時は確認者名を変更してください。";
    } else if (currentNote.indexOf("公開") >= 0 || currentNote.indexOf("内部メモ") >= 0) {
      note = "実運用時はこの値を変更してください。";
    }

    if (note !== currentNote) changed = true;
    return [row[0], row[1], note];
  });

  if (changed) {
    sheet.getRange(2, 1, updated.length, 3).setValues(updated);
  }
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

function getIdHeaderForSheet_(sheetName) {
  return SHEET_ID_HEADERS[sheetName] || "";
}

function buildRecordId_(sheetName) {
  const prefixMap = {
    "出先予定": "OUT",
    "工事予定": "CON",
    "会議予定": "MTG",
    "行事予定": "EVT",
    "作業状況": "WRK",
    "電話履歴": "TEL",
    "車検管理": "CAR",
    "車検履歴": "CARH",
    "社用車予約": "RES",
    "日報": "DR",
    "日報レシート管理": "RCT",
    "免許資格管理": "LIC",
    "備品修理管理": "REP",
    "お知らせ": "NEWS",
    "個人ToDo": "TODO",
    "帳簿PDF履歴": "LEDGER"
  };

  const prefix = prefixMap[sheetName] || "ID";
  const nowText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMddHHmmss"
  );

  return prefix + "-" + nowText + "-" + Utilities.getUuid().slice(0, 8);
}

function ensureIdsForSheet_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

  const idHeader = getIdHeaderForSheet_(sheet.getName());
  if (!idHeader) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf(idHeader) + 1;
  if (idCol <= 0) return;

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  let changed = false;

  const idValues = values.map(row => {
    const currentId = row[idCol - 1];
    const hasData = row.some((value, index) => {
      if (index === idCol - 1) return false;
      return value !== "" && value !== null;
    });

    if (hasData && !currentId) {
      changed = true;
      return [buildRecordId_(sheet.getName())];
    }

    return [currentId || ""];
  });

  if (changed) {
    sheet.getRange(2, idCol, idValues.length, 1).setValues(idValues);
  }
}

function ensureIdForEditedRow_(sheet, row) {
  if (!sheet || row <= 1 || sheet.getLastColumn() < 1) return;

  const idHeader = getIdHeaderForSheet_(sheet.getName());
  if (!idHeader) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf(idHeader) + 1;
  if (idCol <= 0) return;

  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const currentId = values[idCol - 1];

  if (currentId) return;

  const hasData = values.some((value, index) => {
    if (index === idCol - 1) return false;
    return value !== "" && value !== null;
  });

  if (hasData) {
    sheet.getRange(row, idCol).setValue(buildRecordId_(sheet.getName()));
  }
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
    "工事予定": [CLEAR_LABEL, "着工前", "施工前", "施工中", "完了", "延期", "中止"],
    "会議予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "行事予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "作業状況": [CLEAR_LABEL, "着工前", "施工前", "施工中", "完了", "中止", "要確認"],
    // 電話履歴は「電話対応」に一本化するため、状態プルダウンは使いません。
    "電話履歴": [CLEAR_LABEL],
    "社用車予約": [CLEAR_LABEL, "予定", "使用中", "返却済", "中止", "予約重複"],
    "日報": [CLEAR_LABEL, "下書き", "提出済", "確認済", "差戻し", "完了"],
    "日報レシート管理": [CLEAR_LABEL, "未確認", "確認中", "確認済", "差戻し", "未精算", "精算済"],
    "免許資格管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "備品修理管理": [CLEAR_LABEL, "未依頼", "依頼済", "修理中", "返却待ち", "返却済", "完了", "中止"],
    "車検管理": [CLEAR_LABEL, "未対応", "予約済", "実施済", "完了", "更新済"],
    "個人ToDo": [CLEAR_LABEL, "未着手", "対応中", "完了", "中止"],
    "お知らせ": [CLEAR_LABEL, "要確認", "確認済", "完了"],
    // 自動出力シートは基本的に入力規則を付けないが、参照用として広めに定義しておく。
    "一覧スケジュール": [CLEAR_LABEL, "予定", "移動中", "着工前", "施工前", "施工中", "進行中", "未対応", "対応中", "折返し", "要確認", "使用中", "返却待ち", "返却済", "未依頼", "依頼済", "修理中", "予約済", "実施済", "有効", "更新予定", "更新済", "期限切れ", "失効", "下書き", "提出済", "確認済", "差戻し", "未着手", "完了", "延期", "中止", "予約重複"],
    "要確認一覧": [CLEAR_LABEL, "予定", "移動中", "着工前", "施工前", "施工中", "進行中", "未対応", "対応中", "折返し", "要確認", "使用中", "返却待ち", "返却済", "未依頼", "依頼済", "修理中", "予約済", "実施済", "有効", "更新予定", "更新済", "期限切れ", "失効", "下書き", "提出済", "確認済", "差戻し", "未着手", "完了", "延期", "中止", "予約重複", "重要"],
    "過去一覧": [CLEAR_LABEL, "完了", "返却済", "更新済", "確認済", "中止", "期限切れ", "失効"]
  };

  return map[sheetName] || [CLEAR_LABEL, "予定", "進行中", "完了", "延期", "中止", "未対応", "対応中", "要確認"];
}

/**
 * 自動出力・集計系シートかどうか。
 *
 * これらのシートは手入力ではなく、各入力シートから値を集約して作るため、
 * 「状態」「担当」「電話対応」などに入力規則を付けると、
 * 元シート側の状態値が増えた時に setValues() が入力規則違反で止まる。
 * そのため、自動出力シートにはデータ入力規則を付けない。
 */
function isAutoOutputSheet_(sheetName) {
  return [
    "一覧スケジュール",
    "要確認一覧",
    "担当別未読",
    "既読率集計",
    "ダッシュボード",
    "過去一覧",
    "帳簿PDF履歴",
    "帳簿出力設定"
  ].includes(sheetName);
}

function getCopyStatusList_() {
  return [CLEAR_LABEL, "有", "無", "未確認"];
}

function getReturnedStatusList_() {
  return [CLEAR_LABEL, "未", "済"];
}

function getReceiptCategoryList_() {
  return [CLEAR_LABEL, "燃料", "駐車場", "高速代", "消耗品", "資材", "工具", "修理", "その他"];
}

function getReceiptPaymentMethodList_() {
  return [CLEAR_LABEL, "現金", "会社カード", "立替", "請求書", "その他"];
}

function getReceiptAccountingCheckList_() {
  return [CLEAR_LABEL, "未確認", "確認中", "確認済", "差戻し"];
}

function getReceiptSettlementStatusList_() {
  return [CLEAR_LABEL, "未精算", "精算中", "精算済", "対象外"];
}

function setupSettingsSheet() {
  const sheet = ensureSettingsSheet_();
  normalizeSettingsSheetNotes_(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast("設定シートを作成/確認しました");
}

function normalizeSettingsSheetNotesMenu() {
  const sheet = ensureSettingsSheet_();
  normalizeSettingsSheetNotes_(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast("設定シートの備考を整えました");
}


/**
 * 運用ガイドシートを作成する。
 * 社内引継ぎ用に「どのメニューをいつ押すか」をシート内に残す。
 */
function createOperationGuideSheet() {
  createProcedureSheet();
    createCompanyManagementExplanationSheet();
}

function createProcedureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("手順シート");
  if (!sheet) sheet = ss.insertSheet("手順シート");

  sheet.clear();
  sheet.clearFormats();

  const rows = [
    ["区分", "順番", "実行メニュー", "使うタイミング", "注意"],
    ["初期", 1, "本番用初期セットアップ（軽量）", "本番コピー・新規構築時", "サンプルデータは入りません。タイムアウト対策として主要入力シート中心に設定します。"],
    ["初期", 2, "全体更新（軽量）", "初期セットアップ後・通常更新", "普段はこれだけで運用します。"],
    ["初期", 3, "入力シートだけ設定反映（軽量）", "タイムアウト時・担当者変更時", "全シート反映で重い場合はこちらを先に使います。"],
    ["初期", 4, "AppSheetのRegenerate Structure", "列追加・列名変更後", "ID列はKey、PDFリンクはURL、既読列はYes/Noにします。"],
    ["テスト", 1, "テスト環境セットアップを実行", "テストシート作成時", "本番では実行しないでください。サンプルデータが入ります。"],
    ["テスト", 2, "過去移動ルールをチェック", "本番移行前", "実データは移動しない判定用です。"],
    ["日常", 1, "全体更新（軽量）", "通常更新", "一覧・要確認・集計・ダッシュボードを更新します。"],
    ["日常", 2, "一覧帳簿PDFを作成", "PDF帳簿を出したい時", "帳簿出力設定で対象を選びます。"],
    ["日常", 3, "選択行の日報PDFを作成", "日報PDFを作る時", "日報シートの選択行にPDFリンクを保存します。"],
    ["日常", 4, "日報レシート管理に入力", "紙レシートを受け取った時", "未確認・差戻し・未精算だけ要確認一覧に出します。"],
    ["保守", 1, "設定を各シートに反映", "担当者・確認者・列構成を変えた時", "普段は押さなくてよいです。"],
    ["保守", 2, "個人確認グループを作り直す", "既読確認者変更後", "既読列は表示、個人確認列だけ折りたたみます。"],
    ["保守", 3, "日報PDFリンク列を修復", "PDFリンクにTRUE/FALSEが出た時", "一度だけ実行すればよい修復処理です。"],
    ["保守", 4, "車検管理を並び替え", "車検期限順に見たい時", "期限切れ・今日・3日以内を上に寄せます。"],
    ["月次", 1, "過去予定を過去一覧へ移動", "月1回", "未返却備品・未対応電話・未完了ToDoは移動しません。"],
    ["月次", 2, "月次メンテナンス", "月1回", "過去移動を含むため、本番ではバックアップ後に実行します。"],
    ["非表示", 1, "裏方シートを非表示", "運用画面を整理したい時", "設定・SQL・集計系を隠します。"],
    ["社外共有前確認", 1, "匿名化確認", "社外共有・資料化前", "実名・社名・地名・車番・取引先名が入っていないか確認します。"]
  ];

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 440);
  sheet.getRange(1, 1, rows.length, rows[0].length).setWrap(true).setVerticalAlignment("middle");
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

  // 軽量化のため、設定反映では一覧・要確認・ダッシュボード更新までは自動実行しません。
  // 必要に応じてメニューから「全体更新（軽量）」を実行してください。
  SpreadsheetApp.getActiveSpreadsheet().toast("設定シートの担当者名を反映しました。必要に応じて全体更新（軽量）を実行してください。");
}


/**
 * タイムアウト対策：入力シートだけ設定反映する。
 * 本番用初期セットアップ（軽量）ではこちらを使う。
 */
function applySettingsToInputSheetsOnly() {
  clearSettingsCache_();
  const targetNames = getPrimaryInputSheetNamesForSetup_();
  applySettingsToSheetsByNames_(targetNames, "入力シートの設定を反映しました");
}

/**
 * タイムアウト対策：出力・履歴・裏方シートだけ設定反映する。
 * 必要な時だけ実行する。
 */
function applySettingsToOutputSheetsOnly() {
  clearSettingsCache_();
  const targetNames = getOutputAndSupportSheetNamesForSetup_();
  applySettingsToSheetsByNames_(targetNames, "出力・裏方シートの設定を反映しました");
}

function applySettingsToSheetsByNames_(sheetNames, toastMessage) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、設定反映をスキップしました");
    return;
  }

  try {
    clearSettingsCache_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSettingsSheet_();
    const headersMap = getSheetHeaders_();

    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || !headersMap[name]) return;
      setupSheetWithoutClearing_(sheet, headersMap[name]);
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(toastMessage || "設定を反映しました");
  } finally {
    lock.releaseLock();
  }
}

function getPrimaryInputSheetNamesForSetup_() {
  return [
    "出先予定",
    "工事予定",
    "電話履歴",
    "車検管理",
    "社用車予約",
    "日報",
    "日報レシート管理",
    "備品修理管理",
    "免許資格管理",
    "お知らせ",
    "個人ToDo",
    "会議予定",
    "行事予定",
    "作業状況"
  ];
}

function getOutputAndSupportSheetNamesForSetup_() {
  return [
    "一覧スケジュール",
    "要確認一覧",
    "車検履歴",
    "帳簿PDF履歴",
    "帳簿出力設定",
    "担当別未読",
    "既読率集計",
    "過去一覧",
    "ダッシュボード",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集",
    "手順シート",
    "社内管理説明"
  ];
}

function createFilterSafely_(sheet, colCount) {
  if (!sheet || !colCount || colCount < 1) return;

  try {
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
  } catch (e) {
    console.log(sheet.getName() + " の既存フィルタ削除をスキップ: " + e.message);
  }

  try {
    const lastRow = Math.max(sheet.getMaxRows(), 2);
    sheet.getRange(1, 1, lastRow, colCount).createFilter();
  } catch (e) {
    console.log(sheet.getName() + " のフィルタ作成をスキップ: " + e.message);
  }
}

function setupSheetWithoutClearing_(sheet, headers) {
  ensureSheetMinimumRows_(sheet);
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

  createFilterSafely_(sheet, headers.length);

  sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), getApplyRowCount_(sheet) + 1), headers.length).clearDataValidations();

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
  const receiptCategoryList = getReceiptCategoryList_();
  const receiptPaymentMethodList = getReceiptPaymentMethodList_();
  const receiptAccountingCheckList = getReceiptAccountingCheckList_();
  const receiptSettlementStatusList = getReceiptSettlementStatusList_();

  headers.forEach((header, i) => {
    const col = i + 1;

    // 一覧スケジュール・要確認一覧などの自動出力シートには入力規則を付けない。
    // 元シートから集約される状態値が増えても、setValues() が止まらないようにする。
    if (!isAutoOutputSheet_(sheetName)) {
      if (DATE_HEADERS.includes(header)) setDateColumn(sheet, col);
      if (header === "状態") setDropdown(sheet, col, statusList);
      if (header === "担当") setDropdown(sheet, col, staffList);
      if (header === "入力者") setDropdown(sheet, col, staffList);
      if (header === "利用者") setDropdown(sheet, col, staffList);
      if (header === "社用車") setDropdown(sheet, col, carList);
      if (header === "電話対応") setDropdown(sheet, col, phoneList);
      if (header === "重要度") setDropdown(sheet, col, importanceList);
      if (header === "コピー有無") setDropdown(sheet, col, copyList);
      if (header === "返却済") setDropdown(sheet, col, returnedList);
      if (header === "区分" && sheetName === "日報レシート管理") setDropdown(sheet, col, receiptCategoryList);
      if (header === "支払方法") setDropdown(sheet, col, receiptPaymentMethodList);
      if (header === "経理確認") setDropdown(sheet, col, receiptAccountingCheckList);
      if (header === "精算状態") setDropdown(sheet, col, receiptSettlementStatusList);
    }
  });

  setCheckboxesForDataRows(sheet);
  ensureIdsForSheet_(sheet);
  applyColorRules(sheet);
  // 列幅は formatInputSheetsForLongText() / formatDailyReportSheet() で固定調整する。
  // sheet.autoResizeColumns(1, headers.length);

  ensureIdsForSheet_(sheet);

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
  ensureIdsForSheet_(sheet);

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
  ensureSheetMinimumRows_(sheet);
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
  const receiptCategoryList = getReceiptCategoryList_();
  const receiptPaymentMethodList = getReceiptPaymentMethodList_();
  const receiptAccountingCheckList = getReceiptAccountingCheckList_();
  const receiptSettlementStatusList = getReceiptSettlementStatusList_();

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  createFilterSafely_(sheet, headers.length);

  headers.forEach((header, i) => {
    const col = i + 1;

    // 自動出力シートは各入力シートの値を受けるだけなので、入力規則を付けない。
    if (!isAutoOutputSheet_(sheetName)) {
      if (DATE_HEADERS.includes(header)) {
        setDateColumn(sheet, col);
      }

      if (header === "状態") setDropdown(sheet, col, statusList);
      if (header === "担当") setDropdown(sheet, col, staffList);
      if (header === "入力者") setDropdown(sheet, col, staffList);
      if (header === "利用者") setDropdown(sheet, col, staffList);
      if (header === "社用車") setDropdown(sheet, col, carList);
      if (header === "電話対応") setDropdown(sheet, col, phoneList);
      if (header === "重要度") setDropdown(sheet, col, importanceList);
      if (header === "コピー有無") setDropdown(sheet, col, copyList);
      if (header === "返却済") setDropdown(sheet, col, returnedList);
      if (header === "区分" && sheetName === "日報レシート管理") setDropdown(sheet, col, receiptCategoryList);
      if (header === "支払方法") setDropdown(sheet, col, receiptPaymentMethodList);
      if (header === "経理確認") setDropdown(sheet, col, receiptAccountingCheckList);
      if (header === "精算状態") setDropdown(sheet, col, receiptSettlementStatusList);
    }
  });

  ensureIdsForSheet_(sheet);
  applyColorRules(sheet);
  // 列幅は表示幅調整関数で固定する。
  // sheet.autoResizeColumns(1, headers.length);
  
  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    sheet.hideColumns(calendarIdCol);
  }
}

function setDateColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);

  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, rowCount)
    .setNumberFormat("yyyy/mm/dd")
    .setDataValidation(rule);
}

function setDropdown(sheet, col, list) {
  const rowCount = getApplyRowCount_(sheet);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, rowCount).setDataValidation(rule);
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
  ensureIdsForSheet_(sheet);
  applyColorRules(sheet);
  // 自動更新のたびに列幅が戻らないよう、ここでは autoResize しない。
  // sheet.autoResizeColumns(1, headers.length);

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
    // 電話履歴は「状態」と「電話対応」が重複しやすいため、
    // 一覧上の状態も電話対応を表示する。
    "状態": row["電話対応"] || row["対応メモ"] || row["状態"] || "",
    "通知": getNoticeText(row["日時"]),
    "電話対応": row["電話対応"],
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "備考": row["備考"] || row["対応メモ"] || row["メモ"] || "",
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
    "担当": row["担当"] || row["入力者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読":
      row["既読"] === true ||
      row["既読"] === "TRUE" ||
      row["既読"] === "Y",

    ...copyPersonalChecks(row),

    "備考": row["備考"] || row["他現場状況"] || "",
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
  // 列幅は refreshAll() 最後の表示幅調整で戻す。
  // sheet.autoResizeColumns(1, headers.length);

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

  const priority = {
    "期限切れ": 1,
    "今日": 2,
    "3日以内": 3,
    "7日以内": 4
  };

  let rows = [];

  if (source && source.getLastRow() >= 2) {
    const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
    const values = source.getRange(2, 1, source.getLastRow() - 1, sourceHeaders.length).getValues();

    rows = values
      .map(valuesRow => objectFromRow(sourceHeaders, valuesRow))
      .filter(row => shouldIncludeInAlertList_(row, priority))
      .sort((a, b) => {
        const pa = priority[a["通知"]] || 99;
        const pb = priority[b["通知"]] || 99;
        if (pa !== pb) return pa - pb;
        return new Date(a["日付"]) - new Date(b["日付"]);
      });
  }

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

  appendDailyReceiptAlerts_(ss, rows);

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
  // 列幅は refreshAll() 最後の表示幅調整で戻す。
  // sheet.autoResizeColumns(1, headers.length);

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

function appendDailyReceiptAlerts_(ss, rows) {
  const receiptSheet = ss.getSheetByName("日報レシート管理");
  if (!receiptSheet || receiptSheet.getLastRow() < 2) return;

  const headers = receiptSheet.getRange(1, 1, 1, receiptSheet.getLastColumn()).getValues()[0];
  const values = receiptSheet.getRange(2, 1, receiptSheet.getLastRow() - 1, headers.length).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    if (!shouldIncludeDailyReceiptInAlert_(row)) return;

    rows.unshift({
      "日付": row["日付"] || new Date(),
      "種類": "日報レシート",
      "内容": joinText(row["支払先"], joinText(row["内容"], row["金額"] ? String(row["金額"]) + "円" : "")),
      "担当": row["担当"],
      "状態": buildDailyReceiptStatusText_(row),
      "通知": getDailyReceiptNoticeText_(row),
      "電話対応": "",
      "備考": row["備考"] || joinText(row["現場"], row["区分"]),
      "元シート": "日報レシート管理"
    });
  });
}

function shouldIncludeDailyReceiptInAlert_(row) {
  const accounting = String(row["経理確認"] || "").trim();
  const settlement = String(row["精算状態"] || "").trim();

  if (accounting === "差戻し") return true;
  if (accounting === "未確認") return true;
  if (settlement === "未精算") return true;

  return false;
}

function buildDailyReceiptStatusText_(row) {
  const accounting = String(row["経理確認"] || "").trim();
  const settlement = String(row["精算状態"] || "").trim();
  return joinText(accounting, settlement) || "要確認";
}

function getDailyReceiptNoticeText_(row) {
  const accounting = String(row["経理確認"] || "").trim();
  const settlement = String(row["精算状態"] || "").trim();

  if (accounting === "差戻し") return "差戻し";
  if (settlement === "未精算") return "未精算";
  if (accounting === "未確認") return "未確認";
  return "要確認";
}

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
    return phone !== "完了";
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
    ["日報レシート件数", '=COUNTA(日報レシート管理!A2:A)'],
    ["未確認レシート", '=COUNTIF(日報レシート管理!J2:J,"未確認")+COUNTIF(日報レシート管理!J2:J,"差戻し")+COUNTIF(日報レシート管理!K2:K,"未精算")'],
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

    try {
      sortVehicleInspectionSheetSilent_();
    } catch (e) {
      console.log("車検管理の自動並び替えをスキップ: " + e.message);
    }

    // 更新後に列幅・折り返しを見やすい状態へ戻す.
    // onEditではなく、手動/定期の全体更新時だけ実行する。
    try {
      formatInputSheetsForLongText();
    } catch (e) {
      console.log("表示幅調整をスキップ: " + e.message);
    }
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

    try {
      formatInputSheetsForLongText();
    } catch (e) {
      console.log("表示幅調整をスキップ: " + e.message);
    }

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
    "日報レシート管理",
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

    ensureRowsAfterEdit_(sheet, row);

    if (e.range.getValue() === CLEAR_LABEL) {
      e.range.clearContent();
      return;
    }

    ensureIdForEditedRow_(sheet, row);
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
  const priority = [
    "作成日時",
    "更新日",
    "移動日",
    "日付",
    "日時",
    "開始日",
    "終了日",
    "期限",
    "投稿日",
    "登録日",
    "車検期限",
    "次回車検期限",
    "保険期限",
    "次回保険期限",
    "取得日",
    "更新期限",
    "修理依頼日",
    "返却予定日"
  ];

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
    "日報レシート管理",
    "免許資格管理",
    "備品修理管理",
    "過去一覧"
  ];
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("社内管理")
    .addSubMenu(
      ui.createMenu("1. 初期セットアップ")
        .addItem("本番用初期セットアップ（軽量）", "initialSetupForProduction")
        .addItem("本番用初期セットアップ（完全）", "initialSetupForProductionFull")
        .addItem("テスト環境セットアップを実行", "initialSetupForTest")
        .addSeparator()
        .addItem("シート構成を作成（軽量）", "createCompanySheets")
        .addItem("設定シートを作成/確認", "setupSettingsSheet")
        .addItem("設定を各シートに反映", "applySettingsToSheets")
        .addItem("入力シートだけ設定反映（軽量）", "applySettingsToInputSheetsOnly")
        .addItem("出力・裏方シートだけ設定反映", "applySettingsToOutputSheetsOnly")
        .addItem("帳簿出力設定を作成/確認", "setupLedgerOutputSettingsSheet")
        .addItem("SQL関連シートをまとめて作成", "createSqlSupportSheets")
        .addItem("移行前チェックを実行", "runPreMigrationCheck")
        .addItem("手順シートを作成", "createProcedureSheet")
        .addItem("社内管理説明シートを作成", "createCompanyManagementExplanationSheet")
    )
    .addSubMenu(
      ui.createMenu("2. 日常運用")
        .addItem("全体更新（軽量）", "refreshAll")
        .addItem("要確認一覧を更新", "createAlertList")
        .addItem("担当別未読を更新", "createAssigneeUnreadSummary")
        .addItem("既読率集計を更新", "createReadRateSummary")
        .addSeparator()
        .addItem("一覧帳簿PDFを作成", "createLedgerPdf")
        .addItem("選択行の日報 他現場状況を作成", "fillDailyReportOtherSituationForSelectedRow")
        .addItem("選択行の日報文章をGPT作成", "createDailyReportText")
        .addItem("選択行の日報PDFを作成", "createSelectedDailyReportPdf")
        .addItem("当日サマリーPDFを作成", "createDailyReportPdf")
        .addSeparator()
        .addItem("Googleカレンダーへ反映", "syncCalendarEvents")
        .addItem("社用車予約の重複チェック", "checkCarReservationConflicts")
        .addItem("車検更新完了処理", "processCompletedVehicleInspections")
        .addItem("車検管理を並び替え", "sortVehicleInspectionSheet")
        .addItem("主要シートを日時順に並び替え", "sortAllOperationalSheets")
    )
    .addSubMenu(
      ui.createMenu("3. 表示・保守")
        .addItem("入力用シートの表示幅を調整", "formatInputSheetsForLongText")
        .addItem("全シート表示幅を強制調整", "formatAllSheetsFixedWidthsNow")
        .addItem("日報シートの表示幅を調整", "formatDailyReportSheet")
        .addItem("個人確認グループを作り直す", "rebuildCheckGroups")
        .addItem("日報PDFリンク列を修復", "fixDailyReportPdfLinkColumn")
        .addItem("裏方シートを非表示", "hideSupportSheets")
        .addItem("スプレッドシートをバックアップ", "backupSpreadsheet")
    )
    .addSubMenu(
      ui.createMenu("4. 月次・履歴処理")
        .addItem("全体更新（重い：グループ/過去一覧も再作成）", "refreshAllHeavy")
        .addItem("現在のシートを月ごとに折りたたみ", "groupRowsByMonth")
        .addItem("全シートを月ごとに折りたたみ", "groupAllRowsByMonth")
        .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
        .addItem("月次メンテナンス", "monthlyMaintenance")
    )
    .addSubMenu(
      ui.createMenu("5. テスト用")
        .addItem("テスト用サンプルデータを追加", "addSampleDataForTest")
        .addItem("過去移動ルールをチェック", "checkArchiveRulesForTest")
        .addItem("移行前チェックを実行", "runPreMigrationCheck")
    )
    .addSubMenu(
      ui.createMenu("6. トリガー設定")
        .addItem("毎日自動更新トリガーを設定", "installDailyRefreshTrigger")
        .addItem("毎月過去一覧移動トリガーを設定", "installMonthlyArchiveTrigger")
        .addItem("毎月メンテナンストリガーを設定", "installMonthlyMaintenanceTrigger")
        .addItem("自動化トリガーをまとめて設定", "installAllAutomationTriggers")
    )
    .addToUi();
}

/**
 * 本番用初期セットアップ。
 *
 * 目的：後付けメニューを個別に何度も押さなくても、
 * 新しいスプレッドシートや本番コピーで再現できるようにする。
 *
 * 注意：サンプルデータは追加しません。
 */
function initialSetupForProduction() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、本番用初期セットアップをスキップしました");
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("本番用初期セットアップ（軽量）を開始します");

    // v9.6.3：軽量版はタイムアウト対策として段階を絞ります。
    // 全シート設定反映・全体更新・表示幅調整・SQL関連・移行前チェック・裏方非表示は別メニューで実行します。
    createCompanySheets();
    setupSettingsSheet();
    setupLedgerOutputSettingsSheet();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    applySettingsToInputSheetsOnly();
    fixDailyReportPdfLinkColumn();

    SpreadsheetApp.getActiveSpreadsheet().toast("本番用初期セットアップ（軽量）が完了しました。次に 全体更新（軽量）→必要なら個人確認グループ作り直し を実行してください。");
  } finally {
    lock.releaseLock();
  }
}

function initialSetupForProductionFull() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、本番用初期セットアップ（完全）をスキップしました");
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("本番用初期セットアップ（完全）を開始します");

    // 完全版も段階的に実行しますが、データ量が多い場合は時間がかかります。
    createCompanySheets();
    setupSettingsSheet();
    setupLedgerOutputSettingsSheet();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    applySettingsToInputSheetsOnly();
    applySettingsToOutputSheetsOnly();
    createSqlSupportSheets();
    fixDailyReportPdfLinkColumn();
    rebuildCheckGroups();
    refreshAll();

    try {
      formatInputSheetsForLongText();
    } catch (e) {
      console.log("表示幅調整をスキップ: " + e.message);
    }

    runPreMigrationCheck();
    hideSupportSheets();

    SpreadsheetApp.getActiveSpreadsheet().toast("本番用初期セットアップ（完全）が完了しました");
  } finally {
    lock.releaseLock();
  }
}

/**
 * テスト環境セットアップ。
 *
 * 本番用初期セットアップに加えて、架空サンプルデータを追加し、
 * 一覧・要確認・過去移動ルールの確認まで行う。
 * 本番シートでは実行しないでください。
 */
function initialSetupForTest() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、テスト環境セットアップをスキップしました");
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("テスト環境セットアップを開始します");

    // v9.6.3：テスト環境もタイムアウトを避けるため、重い処理を減らします。
    createCompanySheets();
    setupSettingsSheet();
    setupLedgerOutputSettingsSheet();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    applySettingsToInputSheetsOnly();
    fixDailyReportPdfLinkColumn();
    addSampleDataForTest();
    refreshAll();
    checkArchiveRulesForTest();

    // SQL関連・移行前チェック・表示幅調整は必要に応じて個別メニューで実行してください。

    SpreadsheetApp.getActiveSpreadsheet().toast("テスト環境セットアップが完了しました");
  } finally {
    lock.releaseLock();
  }
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

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const personalMembers = getPersonalMembers_();

    const personalCols = personalMembers
      .map(member => headers.indexOf(member) + 1)
      .filter(col => col > 0);

    if (personalCols.length === 0) return;

    const startCol = Math.min(...personalCols);
    const endCol = Math.max(...personalCols);

    // 既読列は全体既読として常時見える方が便利なので、折りたたみ対象にしない。
    // 折りたたむのは個人確認列だけ。PDFリンク・ID列も常に表示する。
    try {
      sheet.getRange(1, startCol, sheet.getMaxRows(), endCol - startCol + 1)
        .shiftColumnGroupDepth(1);
    } catch (e) {
      console.log(name + " の個人確認グループ作成をスキップ: " + e.message);
    }
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
  const ruleRowCount = getApplyRowCount_(sheet);

  function addTextRule(col, text, bg) {
    if (col <= 0) return;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(bg)
        .setRanges([sheet.getRange(2, col, ruleRowCount)])
        .build()
    );
  }

  addTextRule(statusCol, "完了", "#b6d7a8");
  addTextRule(statusCol, "進行中", "#cfe2f3");
  addTextRule(statusCol, "施工中", "#cfe2f3");
  addTextRule(statusCol, "着工前", "#d9d2e9");
  addTextRule(statusCol, "施工前", "#d9d2e9");
  addTextRule(statusCol, "要確認", "#f4cccc");
  addTextRule(statusCol, "予定", "#eeeeee");
  addTextRule(statusCol, "延期", "#ffe599");
  addTextRule(statusCol, "中止", "#ea9999");
  addTextRule(statusCol, "未対応", "#f4cccc");
  addTextRule(statusCol, "対応中", "#fff2cc");
  addTextRule(statusCol, "使用中", "#cfe2f3");
  addTextRule(statusCol, "返却待ち", "#fce5cd");
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
  addTextRule(statusCol, "実施済", "#d9ead3");
  addTextRule(statusCol, "未依頼", "#f4cccc");
  addTextRule(statusCol, "依頼済", "#fff2cc");
  addTextRule(statusCol, "修理中", "#cfe2f3");

  addTextRule(phoneCol, "未対応", "#f4cccc");
  addTextRule(phoneCol, "折返し", "#fff2cc");
  addTextRule(phoneCol, "対応中", "#fff2cc");
  addTextRule(phoneCol, "完了", "#b6d7a8");

  if (readCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied("=$" + columnLetter(readCol) + "2=FALSE")
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(2, readCol, ruleRowCount)])
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
  // sheet.autoResizeColumns(1, headers.length);
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
  const inputterCol = headers.indexOf("入力者") + 1;
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


/**
 * 日報のPDFリンク列に、古いチェックボックス設定の名残で TRUE/FALSE が入った場合の修復用。
 * PDFリンク列はチェックボックスではなく、空欄またはURLとして扱う。
 */
function fixDailyReportPdfLinkColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const pdfCol = headers.indexOf("PDFリンク") + 1;
  if (pdfCol <= 0) return;

  const rowCount = getApplyRowCount_(sheet);
  const range = sheet.getRange(2, pdfCol, rowCount, 1);

  range.clearDataValidations();
  range.setNumberFormat("@");
  range.setWrap(true);

  const values = range.getValues().map(row => {
    const v = row[0];
    if (v === true || v === false || v === "TRUE" || v === "FALSE") return [""];
    return [v];
  });

  range.setValues(values);
  SpreadsheetApp.getActiveSpreadsheet().toast("日報PDFリンク列を修復しました");
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
    "工事予定",
    "会議予定",
    "行事予定",
    "電話履歴",
    "車検管理",
    "車検履歴",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理",
    "お知らせ",
    "個人ToDo",
    "一覧スケジュール",
    "要確認一覧",
    "過去一覧",
    "帳簿PDF履歴"
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
    "日報レシート管理",
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

  if (sheetName === "日報レシート管理") {
    return isDailyReceiptCompletedForArchive_(row);
  }

  if (sheetName === "備品修理管理") {
    // 備品修理管理は、返却予定日・修理依頼日・登録日が過ぎただけでは移動しません。
    // 未返却・依頼中・修理中・返却待ちのものを過去一覧へ移すと管理漏れになるため、
    // 「返却済=済」または「状態=返却済/完了/中止」の場合だけ移動します。
    return isEquipmentRepairCompletedForArchive_(row);
  }

  return false;
}

function isEquipmentRepairCompletedForArchive_(row) {
  const status = String(row["状態"] || "").trim();
  const returned = String(row["返却済"] || "").trim();

  if (returned === "済") return true;
  if (["返却済", "完了", "中止"].includes(status)) return true;

  return false;
}

function isDailyReceiptCompletedForArchive_(row) {
  const accounting = String(row["経理確認"] || "").trim();
  const settlement = String(row["精算状態"] || "").trim();

  return accounting === "確認済" && ["精算済", "対象外"].includes(settlement);
}

function getArchiveDateValue_(row, sheetName) {
  if (sheetName === "行事予定") return row["終了日"] || row["開始日"] || row["日付"];
  if (sheetName === "工事予定") return row["終了日"] || row["開始日"];
  if (sheetName === "電話履歴") return row["日時"];
  if (sheetName === "社用車予約") return row["日付"];
  if (sheetName === "個人ToDo") return row["期限"] || row["登録日"];
  if (sheetName === "日報") return row["日付"];
  if (sheetName === "日報レシート管理") return row["日付"];
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
    staff = row["担当"] || row["入力者"];
    note = row["日報文章"] || row["他現場状況"] || row["備考"] || "";
  } else if (sheetName === "日報レシート管理") {
    dateValue = row["日付"];
    content = joinText(row["支払先"], joinText(row["内容"], row["金額"] ? String(row["金額"]) + "円" : ""));
    staff = row["担当"];
    status = buildDailyReceiptStatusText_(row);
    notice = getDailyReceiptNoticeText_(row);
    note = joinText(joinText("現場: " + (row["現場"] || ""), "区分: " + (row["区分"] || "")), row["備考"] || "");
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
  ensureIdsForSheet_(sheet);
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
    "帳簿出力設定",
    "担当別未読",
    "既読率集計",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集",
    "機能チェック結果",
    "移行前チェック",
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

  let sheet = ss.getSheetByName('既読率集計');
  if (!sheet) {
    sheet = ss.insertSheet('既読率集計');
  }

  const headers = ['担当', '対象件数', '既読件数', '未読件数', '既読率'];

  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  const scheduleSheet = ss.getSheetByName('一覧スケジュール');
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
          if (row[0] === true || row[0] === 'TRUE' || row[0] === 'Y') {
            read++;
          }
        });
      }
    }

    const unread = total - read;
    const rate = total === 0 ? 0 : read / total;

    rows.push({
      '担当': member,
      '対象件数': total,
      '既読件数': read,
      '未読件数': unread,
      '既読率': rate
    });
  });

  writeObjectsToSheet(sheet, rows);

  const rateCol = headers.indexOf('既読率') + 1;
  if (rateCol > 0 && rows.length > 0) {
    sheet.getRange(2, rateCol, rows.length, 1).setNumberFormat('0.0%');
  }

  if (sheet.getMaxRows() >= 2) {
    sheet
      .getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns())
      .clearDataValidations();
  }

  formatReadRateSummarySheet_(sheet);
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


function getDailyReportPeopleToExclude_(data) {
  const people = [];
  [data["入力者"], data["担当"]].forEach(value => {
    const name = String(value || "").trim();
    if (name && !people.includes(name)) people.push(name);
  });
  return people;
}

function fillDailyReportOtherSituationForSelectedRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (!sheet || sheet.getName() !== "日報") {
    SpreadsheetApp.getUi().alert("日報シートで、他現場状況を作成したい行を選択してください。");
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
  const otherCol = headers.indexOf("他現場状況") + 1;

  if (otherCol <= 0) {
    SpreadsheetApp.getUi().alert("他現場状況列が見つかりません。社内管理 → 設定を各シートに反映 を実行してください。");
    return;
  }

  const text = buildDailyReportOtherSiteSituationText_(data);
  sheet.getRange(row, otherCol).setValue(text);
  SpreadsheetApp.getActiveSpreadsheet().toast("選択行の日報 他現場状況を作成しました");
}

/**
 * 日報用の「他現場状況」を作成する。
 *
 * ここでいう「他現場状況」は、入力者本人のToDoや電話ではなく、
 * 入力者・担当者以外が関わっている当日/進行中の現場・工事・作業状況です。
 * 日報に「担当現場以外で、会社全体として動いている主な現場」を補足する目的で使います。
 */
function buildDailyReportOtherSiteSituationText_(data) {
  const excludePeople = getDailyReportPeopleToExclude_(data);
  const reportDate = data["日付"] || new Date();
  const reportSite = String(data["現場"] || "").trim();

  const sections = [];

  const sameDayProjects = getOtherSiteProjectScheduleSummaries_(excludePeople, reportDate, reportSite);
  if (sameDayProjects.length > 0) {
    sections.push("他現場の工事予定: " + sameDayProjects.join(" / "));
  }

  const activeWorks = getOtherSiteActiveWorkSummaries_(excludePeople, reportSite);
  if (activeWorks.length > 0) {
    sections.push("他現場の作業状況: " + activeWorks.join(" / "));
  }

  const outsideSchedules = getOtherSiteOutsideScheduleSummaries_(excludePeople, reportDate);
  if (outsideSchedules.length > 0) {
    sections.push("他担当の出先予定: " + outsideSchedules.join(" / "));
  }

  const carReservations = getOtherSiteCarReservationSummaries_(excludePeople, reportDate);
  if (carReservations.length > 0) {
    sections.push("他担当の社用車予約: " + carReservations.join(" / "));
  }

  return sections.length > 0
    ? sections.join("\n")
    : "入力者・担当者以外の当日工事予定、進行中作業、出先予定、社用車予約はありません。";
}

function getOtherSiteProjectScheduleSummaries_(excludePeople, reportDate, reportSite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("工事予定");
  const result = [];
  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    const staff = String(row["担当"] || "").trim();
    const status = String(row["状態"] || "").trim();
    const site = String(row["現場"] || "").trim();

    if (staff && excludePeople.includes(staff)) return;
    if (["完了", "中止"].includes(status)) return;
    if (!isReportDateWithinRange_(reportDate, row["開始日"], row["終了日"])) return;

    const sameSiteMark = reportSite && site === reportSite ? "同現場・他担当" : "他現場";
    const content = joinText(row["工事名"], site);
    result.push(sameSiteMark + "「" + content + "」" + (staff ? " 担当:" + staff : "") + (status ? "（" + status + "）" : ""));
  });

  return result.slice(0, 8);
}

function getOtherSiteActiveWorkSummaries_(excludePeople, reportSite) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("作業状況");
  const result = [];
  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    const staff = String(row["担当"] || "").trim();
    const status = String(row["状態"] || "").trim();
    const site = String(row["現場"] || "").trim();

    if (staff && excludePeople.includes(staff)) return;
    if (["完了", "中止"].includes(status)) return;

    const sameSiteMark = reportSite && site === reportSite ? "同現場・他担当" : "他現場";
    result.push(sameSiteMark + "「" + joinText(site, row["作業内容"]) + "」" + (staff ? " 担当:" + staff : "") + (status ? "（" + status + "）" : ""));
  });

  return result.slice(0, 8);
}

function getOtherSiteOutsideScheduleSummaries_(excludePeople, reportDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("出先予定");
  const result = [];
  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    const staff = String(row["担当"] || "").trim();
    const status = String(row["状態"] || "").trim();

    if (staff && excludePeople.includes(staff)) return;
    if (!isSameDateOnly_(row["日付"], reportDate)) return;
    if (["完了", "中止"].includes(status)) return;

    result.push(joinText(row["行き先"], row["用件"]) + (staff ? " 担当:" + staff : "") + (status ? "（" + status + "）" : ""));
  });

  return result.slice(0, 5);
}

function getOtherSiteCarReservationSummaries_(excludePeople, reportDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("社用車予約");
  const result = [];
  if (!sheet || sheet.getLastRow() < 2) return result;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  values.forEach(valuesRow => {
    const row = objectFromRow(headers, valuesRow);
    const user = String(row["利用者"] || "").trim();
    const status = String(row["状態"] || "").trim();

    if (user && excludePeople.includes(user)) return;
    if (!isSameDateOnly_(row["日付"], reportDate)) return;
    if (["返却済", "完了", "中止"].includes(status)) return;

    result.push(joinText(row["社用車"], joinText(row["行き先"], row["用途"])) + (user ? " 利用者:" + user : "") + (status ? "（" + status + "）" : ""));
  });

  return result.slice(0, 5);
}

function isReportDateWithinRange_(targetDateValue, startDateValue, endDateValue) {
  if (!targetDateValue || !startDateValue) return false;

  const target = new Date(targetDateValue);
  const start = new Date(startDateValue);
  const end = endDateValue ? new Date(endDateValue) : new Date(startDateValue);

  if (isNaN(target.getTime()) || isNaN(start.getTime()) || isNaN(end.getTime())) return false;

  target.setHours(0, 0, 0, 0);
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  return target.getTime() >= start.getTime() && target.getTime() <= end.getTime();
}

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

  const otherCol = headers.indexOf("他現場状況") + 1;
  if (otherCol > 0 && !data["他現場状況"]) {
    const otherSituation = buildDailyReportOtherSiteSituationText_(data);
    sheet.getRange(row, otherCol).setValue(otherSituation);
    data["他現場状況"] = otherSituation;
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
    "入力者: " + (data["入力者"] || ""),
    "担当: " + (data["担当"] || ""),
    "現場: " + (data["現場"] || ""),
    "作業内容: " + (data["作業内容"] || ""),
    "進捗: " + (data["進捗"] || ""),
    "問題点: " + (data["問題点"] || ""),
    "明日の予定: " + (data["明日の予定"] || ""),
    "入力者以外の現場状況: " + (data["他現場状況"] || ""),
    "備考: " + (data["備考"] || ""),
    "",
    "【条件】",
    "・丁寧な社内向け文章にする",
    "・事実を勝手に追加しない",
    "・入力者以外の現場状況は、必要な場合だけ簡潔に補足する",
    "・他現場の工事予定、作業状況、出先予定、社用車予約がある場合は、日報末尾に確認事項として自然に含める",
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

  const otherCol = headers.indexOf("他現場状況") + 1;
  if (otherCol > 0 && !data["他現場状況"]) {
    const otherSituation = buildDailyReportOtherSiteSituationText_(data);
    sheet.getRange(row, otherCol).setValue(otherSituation);
    data["他現場状況"] = otherSituation;
  }

  const dateText = formatDateForReport_(data["日付"]) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const fileDateText = dateText.replace(/\D/g, "") || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd");
  const inputterText = data["入力者"] || data["担当"] || "入力者未設定";
  const staffText = data["担当"] || data["入力者"] || "担当未設定";
  const siteText = data["現場"] || "現場未設定";

  const doc = DocumentApp.create("日報_" + fileDateText + "_" + staffText + "_" + siteText);
  const body = doc.getBody();

  body.appendParagraph("日報").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("日付: " + dateText);
  body.appendParagraph("入力者: " + inputterText);
  body.appendParagraph("担当: " + staffText);
  body.appendParagraph("現場: " + siteText);
  body.appendParagraph("");

  appendDailyReportParagraph_(body, "作業内容", data["作業内容"]);
  appendDailyReportParagraph_(body, "進捗", data["進捗"]);
  appendDailyReportParagraph_(body, "問題点", data["問題点"]);
  appendDailyReportParagraph_(body, "明日の予定", data["明日の予定"]);
  appendDailyReportParagraph_(body, "入力者以外の現場状況", data["他現場状況"]);
  appendDailyReportParagraph_(body, "日報文章", data["日報文章"]);
  appendDailyReportParagraph_(body, "写真", data["写真"]);
  appendDailyReportParagraph_(body, "状態", data["状態"]);
  appendDailyReportParagraph_(body, "備考", data["備考"]);

  doc.saveAndClose();

  const docFile = DriveApp.getFileById(doc.getId());
  const safeStaff = String(staffText).replace(/[\/:*?"<>|]/g, "_");
  const safeSite = String(siteText).replace(/[\/:*?"<>|]/g, "_");
  const pdfBlob = docFile.getBlob().getAs(MimeType.PDF).setName("日報_" + fileDateText + "_" + safeStaff + "_" + safeSite + ".pdf");

  const pdfFolder = getOrCreateChildFolder_("日報PDF");
  const pdfFile = pdfFolder.createFile(pdfBlob);
  const pdfUrl = pdfFile.getUrl();

  const pdfLinkCol = headers.indexOf("PDFリンク") + 1;
  if (pdfLinkCol > 0) {
    sheet.getRange(row, pdfLinkCol).setValue(pdfUrl);
  }

  docFile.setTrashed(true);

  SpreadsheetApp.getActiveSpreadsheet().toast("選択行の日報PDFを作成し、PDFリンクを保存しました");
  return pdfUrl;
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
    "日報",
    "日報レシート管理",
    "免許資格管理",
    "備品修理管理"
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
    ensureIdsForSheet_(sheet);
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
    "入力者",
    "担当",
    "現場",
    "作業内容",
    "進捗",
    "問題点",
    "明日の予定",
    "他現場状況",
    "写真",
    "日報文章",
    "状態",
    "備考",
    "既読",
    "山田",
    "高橋",
    "鈴木"
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
    "入力者",
    "担当",
    "現場",
    "作業内容",
    "進捗",
    "問題点",
    "明日の予定",
    "他現場状況",
    "写真",
    "日報文章",
    "状態",
    "備考",
    "既読",
    "山田",
    "高橋",
    "鈴木"
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  const lastCol = headers.length;
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.getRange(2, 1, getApplyRowCount_(sheet), 1).setNumberFormat("yyyy/mm/dd");

  sheet.getRange(2, 14, getApplyRowCount_(sheet), 4).insertCheckboxes();

  sheet.autoResizeColumns(1, lastCol);
}







function setupDailyReportSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const dateCol = headers.indexOf("日付") + 1;
  const staffCol = headers.indexOf("担当") + 1;
  const inputterCol = headers.indexOf("入力者") + 1;
  const statusCol = headers.indexOf("状態") + 1;

  if (dateCol > 0) {
    const dateRule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, dateCol, getApplyRowCount_(sheet), 1)
      .setNumberFormat("yyyy/mm/dd")
      .setDataValidation(dateRule);
  }

  if (staffCol > 0) {
    const staffRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, ...getStaffMembers_()], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, staffCol, getApplyRowCount_(sheet), 1)
      .setDataValidation(staffRule);
  }

  if (inputterCol > 0) {
    const inputterRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, ...getStaffMembers_()], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, inputterCol, getApplyRowCount_(sheet), 1)
      .setDataValidation(inputterRule);
  }

  if (statusCol > 0) {
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([CLEAR_LABEL, "予定", "進行中", "施工中", "完了", "延期", "中止", "要確認"], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, statusCol, getApplyRowCount_(sheet), 1)
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

  sheet.getRange(2, carNumberCol, getApplyRowCount_(sheet), 1)
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

  const checkNames = ["既読", "山田", "高橋", "鈴木"];

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

  // 帳簿PDFは共有用なので、出力前に集計系だけ軽く更新する。
  try {
    refreshSummaryOnly_();
  } catch (e) {
    console.log("帳簿PDF作成前の軽量更新をスキップ: " + e.message);
  }

  const sections = getSelectedLedgerSections_();
  if (sections.length === 0) {
    SpreadsheetApp.getUi().alert("帳簿出力設定で出力対象が選択されていません。\n出力する列にチェックを入れてください。");
    return "";
  }

  const tempSheetName = "PDF_一覧帳簿_一時";
  let temp = ss.getSheetByName(tempSheetName);
  if (!temp) temp = ss.insertSheet(tempSheetName);

  try {
    temp.showSheet();
  } catch (e) {}

  temp.clear();
  temp.clearFormats();
  temp.getRange(1, 1, temp.getMaxRows(), temp.getMaxColumns()).clearDataValidations();
  temp.setConditionalFormatRules([]);

  let row = 1;
  const outputDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm"
  );

  temp.getRange(row, 1, 1, 10).merge();
  temp.getRange(row, 1)
    .setValue("社内共有業務管理システム 一覧帳簿")
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBackground("#d9ead3");
  temp.setRowHeight(row, 38);
  row += 2;

  temp.getRange(row, 1).setValue("出力日時").setFontWeight("bold");
  temp.getRange(row, 2, 1, 3).merge().setValue(outputDate);
  temp.getRange(row, 5).setValue("用途").setFontWeight("bold");
  temp.getRange(row, 6, 1, 4).merge().setValue("他部門共有・経理確認・管理者確認用");
  row += 2;

  sections.forEach((section, index) => {
    row = writeLedgerSection_(
      temp,
      row,
      (index + 1) + ". " + section.title,
      section.sheetName,
      section.outputHeaders,
      section.maxRows
    );
  });

  formatLedgerPdfSheet_(temp);

  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  const fileName = "一覧帳簿_" + Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd_HHmm"
  );

  const pdfBlob = exportLedgerSheetToPdfBlob_(ss, temp, fileName);
  const folder = getOrCreateChildFolder_("一覧帳簿PDF");
  const pdfFile = folder.createFile(pdfBlob);

  recordLedgerPdfHistory_(
    pdfFile.getUrl(),
    sections.map(section => section.title).join("・"),
    "出力時点"
  );

  try {
    temp.hideSheet();
  } catch (e) {}

  SpreadsheetApp.getUi().alert("一覧帳簿PDFを作成しました。\n\n" + pdfFile.getUrl());
  return pdfFile.getUrl();
}

function setupLedgerOutputSettingsSheet() {
  ensureLedgerOutputSettingsSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast("帳簿出力設定を作成/確認しました");
}

function ensureLedgerOutputSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("帳簿出力設定");
  if (!sheet) sheet = ss.insertSheet("帳簿出力設定");

  const headers = ["出力する", "帳簿名", "シート名", "出力列", "最大行数", "備考"];
  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeader = currentHeaders[0] === "出力する" && currentHeaders[1] === "帳簿名";

  if (!hasHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#d9ead3")
      .setHorizontalAlignment("center");

    const rows = getDefaultLedgerOutputSettings_().map(item => [
      item.enabled,
      item.title,
      item.sheetName,
      item.outputHeaders.join(","),
      item.maxRows,
      item.note || ""
    ]);

    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const checkboxRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const normalized = checkboxRange.getValues().map(row => {
    const v = row[0];
    return [v === true || v === "TRUE" || v === "Y" || v === "1"];
  });
  checkboxRange.setValues(normalized);
  checkboxRange.insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 520);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 260);
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setWrap(true).setVerticalAlignment("middle");

  return sheet;
}

function getDefaultLedgerOutputSettings_() {
  return [
    { enabled: true, title: "要確認一覧", sheetName: "要確認一覧", outputHeaders: ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "備考", "元シート"], maxRows: 40, note: "期限切れ・今日・3日以内など" },
    { enabled: true, title: "日報一覧", sheetName: "日報", outputHeaders: ["日付", "入力者", "担当", "現場", "作業内容", "進捗", "他現場状況", "状態", "備考"], maxRows: 40, note: "入力者以外の現場状況も出力" },
    { enabled: true, title: "日報レシート一覧", sheetName: "日報レシート管理", outputHeaders: ["日付", "担当", "現場", "支払先", "内容", "区分", "金額", "支払方法", "経理確認", "精算状態", "備考"], maxRows: 40, note: "紙レシート・精算確認用" },
    { enabled: true, title: "備品修理一覧", sheetName: "備品修理管理", outputHeaders: ["登録日", "備品名", "場所", "内容", "担当", "修理依頼日", "返却予定日", "返却済", "状態", "備考"], maxRows: 40, note: "未返却・修理中確認" },
    { enabled: true, title: "車検・保険期限一覧", sheetName: "車検管理", outputHeaders: ["車両名", "車番", "車検期限", "保険期限", "状態", "通知", "備考"], maxRows: 40, note: "車検・保険期限確認" },
    { enabled: true, title: "車検更新履歴", sheetName: "車検履歴", outputHeaders: ["更新日", "車両名", "車番", "旧車検期限", "新車検期限", "旧保険期限", "新保険期限", "担当", "備考"], maxRows: 40, note: "更新ログ" },
    { enabled: false, title: "電話履歴一覧", sheetName: "電話履歴", outputHeaders: ["日時", "相手", "内容", "電話対応", "担当", "対応メモ", "通知", "備考"], maxRows: 40, note: "電話対応確認用" },
    { enabled: false, title: "免許資格一覧", sheetName: "免許資格管理", outputHeaders: ["担当", "資格名", "区分", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"], maxRows: 40, note: "資格期限確認用" },
    { enabled: false, title: "個人ToDo一覧", sheetName: "個人ToDo", outputHeaders: ["登録日", "担当", "内容", "期限", "状態", "通知", "備考"], maxRows: 40, note: "未完了ToDo確認用" },
    { enabled: false, title: "工事予定一覧", sheetName: "工事予定", outputHeaders: ["工事名", "現場", "開始日", "終了日", "状態", "担当", "通知", "電話対応", "備考"], maxRows: 40, note: "工事予定確認用" },
    { enabled: false, title: "出先予定一覧", sheetName: "出先予定", outputHeaders: ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", "備考"], maxRows: 40, note: "出先確認用" }
  ];
}

function getSelectedLedgerSections_() {
  const sheet = ensureLedgerOutputSettingsSheet_();
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const sections = [];

  values.forEach(row => {
    const enabled = row[0] === true || row[0] === "TRUE" || row[0] === "Y";
    const title = String(row[1] || "").trim();
    const sheetName = String(row[2] || "").trim();
    const headersText = String(row[3] || "").trim();
    const maxRows = Number(row[4]) || 40;

    if (!enabled || !title || !sheetName || !headersText) return;

    sections.push({
      title: title,
      sheetName: sheetName,
      outputHeaders: headersText.split(",").map(value => value.trim()).filter(Boolean),
      maxRows: maxRows
    });
  });

  return sections;
}


function writeLedgerSection_(temp, startRow, title, sourceSheetName, outputHeaders, maxRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(sourceSheetName);
  let row = startRow;

  temp.getRange(row, 1, 1, Math.max(outputHeaders.length, 1))
    .merge()
    .setValue(title)
    .setFontWeight("bold")
    .setBackground("#cfe2f3")
    .setVerticalAlignment("middle");
  row++;

  temp.getRange(row, 1, 1, outputHeaders.length)
    .setValues([outputHeaders])
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);
  row++;

  if (!source || source.getLastRow() < 2) {
    temp.getRange(row, 1).setValue("データなし").setFontStyle("italic");
    return row + 2;
  }

  const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
  const values = source.getRange(2, 1, source.getLastRow() - 1, sourceHeaders.length).getValues();

  const rows = values
    .map(valuesRow => objectFromRow(sourceHeaders, valuesRow))
    .filter(obj => {
      if (obj["種類"] === "月見出し") return false;
      return outputHeaders.some(header => {
        const value = obj[header];
        return value !== "" && value !== null && value !== undefined;
      });
    })
    .slice(0, maxRows || 40)
    .map(obj => outputHeaders.map(header => formatLedgerValue_(obj[header])));

  if (rows.length === 0) {
    temp.getRange(row, 1).setValue("データなし").setFontStyle("italic");
    return row + 2;
  }

  temp.getRange(row, 1, rows.length, outputHeaders.length)
    .setValues(rows)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment("middle")
    .setWrap(true);

  return row + rows.length + 2;
}

function formatLedgerValue_(value) {
  if (value === null || value === undefined) return "";

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy/MM/dd");
  }

  if (value === true) return "済";
  if (value === false) return "";

  const text = String(value || "");
  if (text.length > 90) return text.substring(0, 90) + "…";
  return text;
}

function formatLedgerPdfSheet_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 10);

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily("Arial")
    .setFontSize(9)
    .setVerticalAlignment("middle")
    .setWrap(true);

  const widths = [90, 100, 180, 90, 90, 95, 95, 180, 110, 140];
  widths.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });

  if (lastRow >= 1) {
    sheet.setRowHeights(1, lastRow, 32);
  }

  try {
    sheet.setFrozenRows(0);
  } catch (e) {}
}

function exportLedgerSheetToPdfBlob_(ss, sheet, fileName) {
  const spreadsheetId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export" +
    "?format=pdf" +
    "&gid=" + sheetId +
    "&size=A4" +
    "&portrait=false" +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=true" +
    "&gridlines=false" +
    "&fzr=false" +
    "&top_margin=0.50" +
    "&bottom_margin=0.50" +
    "&left_margin=0.40" +
    "&right_margin=0.40";

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  return response.getBlob().setName(fileName + ".pdf");
}

function recordLedgerPdfHistory_(url, target, periodText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("帳簿PDF履歴");

  if (!sheet) {
    sheet = ss.insertSheet("帳簿PDF履歴");
    setupSheetWithoutClearing_(sheet, getSheetHeaders_()["帳簿PDF履歴"]);
  }

  const row = {
    "作成日時": new Date(),
    "対象": target || "一覧帳簿",
    "期間": periodText || "出力時点",
    "PDFリンク": url || "",
    "備考": "他部門・経理・管理者共有用"
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

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 車検期限の右に「次回車検期限」を追加
  let nextInspectionCol = headers.indexOf("次回車検期限") + 1;
  if (nextInspectionCol <= 0) {
    const inspectionCol = headers.indexOf("車検期限") + 1;
    if (inspectionCol > 0) {
      sheet.insertColumnAfter(inspectionCol);
      sheet.getRange(1, inspectionCol + 1).setValue("次回車検期限");
    }
  }

  headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 保険期限の右に「次回保険期限」を追加
  let nextInsuranceCol = headers.indexOf("次回保険期限") + 1;
  if (nextInsuranceCol <= 0) {
    const insuranceCol = headers.indexOf("保険期限") + 1;
    if (insuranceCol > 0) {
      sheet.insertColumnAfter(insuranceCol);
      sheet.getRange(1, insuranceCol + 1).setValue("次回保険期限");
    }
  }
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
  const nextInsuranceCol = headers.indexOf("次回保険期限") + 1;
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

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const historyRows = [];

  let processedCount = 0;
  let skippedCount = 0;

  values.forEach(valuesRow => {
    const status = valuesRow[statusCol - 1];
    if (status !== "完了" && status !== "更新済") return;

    const oldInspectionDate = valuesRow[inspectionCol - 1];
    const nextInspectionDate = valuesRow[nextInspectionCol - 1];
    const oldInsuranceDate = insuranceCol > 0 ? valuesRow[insuranceCol - 1] : "";
    const nextInsuranceDate = nextInsuranceCol > 0 ? valuesRow[nextInsuranceCol - 1] : "";
    const currentNote = noteCol > 0 ? String(valuesRow[noteCol - 1] || "") : "";

    // 次回車検期限がない場合は、車検更新そのものをしない。
    if (!nextInspectionDate) {
      if (noteCol > 0) {
        valuesRow[noteCol - 1] = appendNote_(
          currentNote,
          "次回車検期限が未入力のため、車検更新処理をスキップ"
        );
      }
      skippedCount++;
      return;
    }

    // 保険期限が入力されている車両は、保険も管理対象とみなす。
    // その場合、次回保険期限が空欄なら保険更新漏れ防止のため処理しない。
    // 保険期限が空欄の車両は、保険管理なしとして車検だけ更新できる。
    if (oldInsuranceDate && !nextInsuranceDate) {
      if (noteCol > 0) {
        valuesRow[noteCol - 1] = appendNote_(
          currentNote,
          "保険期限は入力済みですが、次回保険期限が未入力のため、更新処理をスキップ"
        );
      }
      skippedCount++;
      return;
    }

    const rowObj = {
      "更新日": new Date(),
      "車両名": carNameCol > 0 ? valuesRow[carNameCol - 1] : "",
      "車番": carNumberCol > 0 ? valuesRow[carNumberCol - 1] : "",
      "旧車検期限": oldInspectionDate,
      "新車検期限": nextInspectionDate,
      "旧保険期限": oldInsuranceDate,
      "新保険期限": nextInsuranceDate || "",
      "担当": staffCol > 0 ? valuesRow[staffCol - 1] : "",
      "備考": currentNote
    };

    historyRows.push(historyHeaders.map(header => rowObj[header] !== undefined ? rowObj[header] : ""));

    valuesRow[inspectionCol - 1] = nextInspectionDate;
    valuesRow[nextInspectionCol - 1] = "";

    if (insuranceCol > 0 && nextInsuranceCol > 0) {
      if (nextInsuranceDate) {
        valuesRow[insuranceCol - 1] = nextInsuranceDate;
      }
      valuesRow[nextInsuranceCol - 1] = "";
    }

    valuesRow[statusCol - 1] = "";

    if (noticeCol > 0) {
      valuesRow[noticeCol - 1] = getNoticeText(nextInspectionDate);
    }

    processedCount++;
  });

  if (processedCount > 0 || skippedCount > 0) {
    sheet.getRange(2, 1, values.length, lastCol).setValues(values);
  }

  if (historyRows.length > 0) {
    const startRow = historySheet.getLastRow() + 1;
    historySheet.getRange(startRow, 1, historyRows.length, historyHeaders.length).setValues(historyRows);
    try {
      formatVehicleHistorySheet_();
    } catch (e) {
      console.log("車検履歴の表示幅調整をスキップ: " + e.message);
    }
  }

  sortVehicleInspectionSheet();

  try {
    formatVehicleInspectionSheet_();
  } catch (e) {
    console.log("車検管理の表示幅調整をスキップ: " + e.message);
  }

  try {
    refreshSummaryOnly_();
  } catch (e) {}

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "車検更新完了処理：" + processedCount + "件 / スキップ：" + skippedCount + "件"
  );
}

function appendNote_(currentNote, message) {
  const dateText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy/MM/dd HH:mm"
  );

  const addText = "[" + dateText + "] " + message;
  if (!currentNote) return addText;
  return currentNote + "\n" + addText;
}


/**
 * 車検管理シートを車検期限順に並び替える。
 * 状態が完了・更新済の行は下へ送る。
 */
function sortVehicleInspectionSheetSilent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");
  if (!sheet || sheet.getLastRow() < 3) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("状態") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const inspectionCol = headers.indexOf("車検期限") + 1;
  const insuranceCol = headers.indexOf("保険期限") + 1;
  const vehicleCol = headers.indexOf("車両名") + 1;
  if (inspectionCol <= 0) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const noticePriority = {"期限切れ": 1, "今日": 2, "3日以内": 3, "7日以内": 4, "": 9};

  values.sort((a, b) => {
    const statusA = statusCol > 0 ? a[statusCol - 1] : "";
    const statusB = statusCol > 0 ? b[statusCol - 1] : "";
    const doneA = statusA === "完了" || statusA === "更新済";
    const doneB = statusB === "完了" || statusB === "更新済";
    if (doneA !== doneB) return doneA ? 1 : -1;

    if (noticeCol > 0) {
      const pa = noticePriority[a[noticeCol - 1] || ""] || 9;
      const pb = noticePriority[b[noticeCol - 1] || ""] || 9;
      if (pa !== pb) return pa - pb;
    }

    const timeA = getSortableTime_(a[inspectionCol - 1]);
    const timeB = getSortableTime_(b[inspectionCol - 1]);
    if (timeA !== timeB) return timeA - timeB;

    if (insuranceCol > 0) {
      const insA = getSortableTime_(a[insuranceCol - 1]);
      const insB = getSortableTime_(b[insuranceCol - 1]);
      if (insA !== insB) return insA - insB;
    }

    if (vehicleCol > 0) {
      return String(a[vehicleCol - 1] || "").localeCompare(String(b[vehicleCol - 1] || ""), "ja");
    }
    return 0;
  });

  sheet.getRange(2, 1, values.length, lastCol).setValues(values);
}

function getSortableTime_(value) {
  const date = new Date(value);
  return isNaN(date.getTime()) ? 9999999999999 : date.getTime();
}

function sortVehicleInspectionSheet() {
  sortVehicleInspectionSheetSilent_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");
  if (sheet) {
    setCheckboxesForDataRows(sheet);
    applyColorRules(sheet);
    try {
      formatVehicleInspectionSheet_();
    } catch (e) {
      console.log("車検管理の表示幅調整をスキップ: " + e.message);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("車検管理を車検期限順に並び替えました");
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
 * Google Sheets版の業務データを、将来的にRDB化できる設計として整理する。
 */


/**
 * テスト用サンプルデータを追加する。
 * 社外共有・テスト用の架空名/架空現場のみを使います。
 * 既存データは消さず、各シートの末尾へ追記します。
 */
function addSampleDataForTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_();
  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    setupSheetWithoutClearing_(sheet, headersMap[name]);
  });

  const today = new Date();
  const staff = getStaffMembers_();
  const a = staff[0] || "山田";
  const b = staff[1] || "鈴木";
  const c = staff[2] || "田中";

  appendObjectsToSheet(ss.getSheetByName("出先予定"), [
    {"日付": today, "行き先": "役所サンプル", "用件": "書類提出", "担当": a, "社用車": "軽トラ②", "状態": "予定", "通知": getNoticeText(today), "電話対応": "", "備考": "サンプル"},
    {"日付": addDaysForSample_(today, -3), "行き先": "仮設資材置場", "用件": "資材確認", "担当": b, "社用車": "軽ワゴン③", "状態": "完了", "通知": getNoticeText(addDaysForSample_(today, -3)), "電話対応": "", "備考": "過去移動テスト用"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("工事予定"), [
    {"工事名": "道路補修サンプル工事", "現場": "中央町サンプル現場", "開始日": today, "終了日": addDaysForSample_(today, 7), "状態": "施工中", "担当": a, "通知": getNoticeText(today), "電話対応": "", "備考": "進行中サンプル"},
    {"工事名": "側溝清掃サンプル工事", "現場": "北町サンプル現場", "開始日": addDaysForSample_(today, -10), "終了日": addDaysForSample_(today, -2), "状態": "完了", "担当": c, "通知": getNoticeText(addDaysForSample_(today, -10)), "電話対応": "", "備考": "過去移動テスト用"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("会議予定"), [
    {"日付": today, "会議名": "工程会議", "内容": "今週の工程確認", "担当": a, "状態": "予定", "通知": getNoticeText(today), "資料": "", "備考": "サンプル"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("行事予定"), [
    {"開始日": addDaysForSample_(today, 3), "終了日": addDaysForSample_(today, 3), "行事名": "安全講習", "内容": "社内安全講習", "担当": b, "状態": "予定", "通知": getNoticeText(addDaysForSample_(today, 3)), "備考": "サンプル"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("作業状況"), [
    {"現場": "南町サンプル現場", "作業内容": "舗装準備", "状態": "施工中", "担当": a, "写真": "", "通知": "", "備考": "サンプル"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("電話履歴"), [
    {"日時": today, "相手": "サンプル取引先A", "内容": "見積確認の折返し依頼", "電話対応": "未対応", "担当": a, "対応メモ": "折返し必要", "通知": getNoticeText(today), "備考": "要確認に出る想定"},
    {"日時": addDaysForSample_(today, -2), "相手": "サンプル取引先B", "内容": "納期確認", "電話対応": "完了", "担当": b, "対応メモ": "対応済", "通知": getNoticeText(addDaysForSample_(today, -2)), "備考": "過去移動テスト用"},
    {"日時": today, "相手": "サンプル取引先C", "内容": "担当者不在", "電話対応": "折返し", "担当": c, "対応メモ": "午後対応", "通知": getNoticeText(today), "備考": "要確認に出る想定"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("車検管理"), [
    {"車両名": "サンプル1号車", "車番": "地域000あ0001", "車検期限": addDaysForSample_(today, 20), "次回車検期限": addYearsForSample_(addDaysForSample_(today, 20), 1), "通知": getNoticeText(addDaysForSample_(today, 20)), "保険期限": addDaysForSample_(today, 20), "次回保険期限": addYearsForSample_(addDaysForSample_(today, 20), 1), "状態": "予約済", "写真": "", "備考": "車検テスト用"},
    {"車両名": "サンプル2号車", "車番": "地域000あ0002", "車検期限": addDaysForSample_(today, -5), "次回車検期限": addYearsForSample_(today, 1), "通知": getNoticeText(addDaysForSample_(today, -5)), "保険期限": "", "次回保険期限": "", "状態": "未対応", "写真": "", "備考": "保険なし車両テスト用"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("社用車予約"), [
    {"日付": today, "開始時刻": "09:00", "終了時刻": "11:00", "社用車": "軽トラ②", "利用者": a, "行き先": "中央町サンプル現場", "用途": "現場確認", "状態": "予定", "通知": getNoticeText(today), "備考": "サンプル"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("日報"), [
    {"日付": today, "入力者": a, "担当": a, "現場": "中央町サンプル現場", "作業内容": "舗装前準備", "進捗": "予定通り", "問題点": "特になし", "明日の予定": "資材搬入", "他現場状況": "北町現場は完了確認中", "写真": "", "日報文章": "", "状態": "下書き", "備考": "PDFリンク列確認用"},
    {"日付": addDaysForSample_(today, -2), "入力者": b, "担当": b, "現場": "北町サンプル現場", "作業内容": "清掃作業", "進捗": "完了", "問題点": "なし", "明日の予定": "なし", "他現場状況": "", "写真": "", "日報文章": "作業完了済みのサンプル日報です。", "状態": "提出済", "備考": "過去移動テスト用"}
  ]);



  appendObjectsToSheet(ss.getSheetByName("日報レシート管理"), [
    {"日付": today, "担当": a, "現場": "中央町サンプル現場", "支払先": "サンプル給油所", "内容": "軽油代", "区分": "燃料", "金額": 5200, "支払方法": "立替", "レシート写真": "", "経理確認": "未確認", "精算状態": "未精算", "通知": "未精算", "備考": "要確認に出る想定"},
    {"日付": addDaysForSample_(today, -3), "担当": b, "現場": "北町サンプル現場", "支払先": "サンプルホームセンター", "内容": "消耗品購入", "区分": "消耗品", "金額": 1800, "支払方法": "現金", "レシート写真": "", "経理確認": "確認済", "精算状態": "精算済", "通知": "", "備考": "過去移動テスト用"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("免許資格管理"), [
    {"担当": a, "資格名": "サンプル技能講習", "区分": "技能講習", "取得日": addYearsForSample_(today, -3), "更新期限": addDaysForSample_(today, 10), "コピー有無": "有", "状態": "更新予定", "通知": getNoticeText(addDaysForSample_(today, 10)), "備考": "期限確認用"},
    {"担当": b, "資格名": "サンプル特別教育", "区分": "特別教育", "取得日": addYearsForSample_(today, -1), "更新期限": addYearsForSample_(today, 4), "コピー有無": "有", "状態": "有効", "通知": getNoticeText(addYearsForSample_(today, 4)), "備考": "有効サンプル"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("備品修理管理"), [
    {"登録日": addDaysForSample_(today, -7), "備品名": "サンプル発電機", "場所": "倉庫A", "内容": "エンジン不調", "担当": a, "修理依頼日": addDaysForSample_(today, -6), "返却予定日": addDaysForSample_(today, -1), "返却済": "未", "状態": "修理中", "通知": getNoticeText(addDaysForSample_(today, -1)), "備考": "未返却なので過去移動しない想定"},
    {"登録日": addDaysForSample_(today, -12), "備品名": "サンプル投光器", "場所": "倉庫B", "内容": "ランプ交換", "担当": b, "修理依頼日": addDaysForSample_(today, -11), "返却予定日": addDaysForSample_(today, -3), "返却済": "済", "状態": "返却済", "通知": getNoticeText(addDaysForSample_(today, -3)), "備考": "過去移動する想定"},
    {"登録日": addDaysForSample_(today, -15), "備品名": "サンプル水中ポンプ", "場所": "倉庫C", "内容": "修理不可", "担当": c, "修理依頼日": addDaysForSample_(today, -14), "返却予定日": addDaysForSample_(today, -4), "返却済": "未", "状態": "中止", "通知": getNoticeText(addDaysForSample_(today, -4)), "備考": "中止なので過去移動する想定"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("お知らせ"), [
    {"投稿日": today, "タイトル": "サンプルお知らせ", "内容": "明日の朝礼場所を確認してください。", "投稿者": a, "重要度": "高", "通知": "重要", "備考": "未読なら要確認"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("個人ToDo"), [
    {"登録日": today, "担当": a, "内容": "サンプル見積確認", "期限": today, "状態": "未着手", "通知": getNoticeText(today), "備考": "要確認に出る想定"},
    {"登録日": addDaysForSample_(today, -5), "担当": b, "内容": "サンプル書類提出", "期限": addDaysForSample_(today, -1), "状態": "完了", "通知": getNoticeText(addDaysForSample_(today, -1)), "備考": "過去移動テスト用"}
  ]);

  appendObjectsToSheet(ss.getSheetByName("過去一覧"), [
    {"日付": addDaysForSample_(today, -40), "種類": "出先予定", "内容": "仮設資材置場 書類回収", "担当": a, "状態": "完了", "通知": "期限切れ", "電話対応": "", "元シート": "出先予定", "備考": "過去一覧サンプル", "移動日": today},
    {"日付": addDaysForSample_(today, -25), "種類": "電話履歴", "内容": "サンプル取引先B 納期確認", "担当": b, "状態": "完了", "通知": "期限切れ", "電話対応": "完了", "元シート": "電話履歴", "備考": "完了電話の過去一覧サンプル", "移動日": today},
    {"日付": addDaysForSample_(today, -10), "種類": "備品修理", "内容": "サンプル投光器 ランプ交換", "担当": b, "状態": "返却済", "通知": "期限切れ", "電話対応": "", "元シート": "備品修理管理", "備考": "返却済備品の過去一覧サンプル", "移動日": today}
  ]);

  refreshAll();
  SpreadsheetApp.getActiveSpreadsheet().toast("テスト用サンプルデータを追加しました");
}

function addDaysForSample_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function addYearsForSample_(date, years) {
  const d = new Date(date);
  d.setFullYear(d.getFullYear() + years);
  return d;
}

/**
 * 過去一覧へ移動してよい/いけない判定を、実データ削除なしで検査します。
 */
function checkArchiveRulesForTest() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const cases = [
    {name: "未返却備品は移動しない", sheetName: "備品修理管理", expected: false, row: {"登録日": addDaysForSample_(today, -7), "修理依頼日": addDaysForSample_(today, -6), "返却予定日": addDaysForSample_(today, -1), "返却済": "未", "状態": "修理中"}},
    {name: "返却済備品は移動する", sheetName: "備品修理管理", expected: true, row: {"登録日": addDaysForSample_(today, -7), "返却予定日": addDaysForSample_(today, -1), "返却済": "済", "状態": "返却済"}},
    {name: "完了電話は移動する", sheetName: "電話履歴", expected: true, row: {"日時": addDaysForSample_(today, -1), "電話対応": "完了"}},
    {name: "未対応電話は移動しない", sheetName: "電話履歴", expected: false, row: {"日時": addDaysForSample_(today, -1), "電話対応": "未対応"}},
    {name: "完了ToDoは移動する", sheetName: "個人ToDo", expected: true, row: {"登録日": addDaysForSample_(today, -5), "期限": addDaysForSample_(today, -1), "状態": "完了"}},
    {name: "未着手ToDoは移動しない", sheetName: "個人ToDo", expected: false, row: {"登録日": addDaysForSample_(today, -5), "期限": addDaysForSample_(today, -1), "状態": "未着手"}},
    {name: "提出済日報は移動する", sheetName: "日報", expected: true, row: {"日付": addDaysForSample_(today, -1), "状態": "提出済"}},
    {name: "下書き日報は移動しない", sheetName: "日報", expected: false, row: {"日付": addDaysForSample_(today, -1), "状態": "下書き"}},
    {name: "確認済・精算済レシートは移動する", sheetName: "日報レシート管理", expected: true, row: {"日付": addDaysForSample_(today, -3), "経理確認": "確認済", "精算状態": "精算済"}},
    {name: "未精算レシートは移動しない", sheetName: "日報レシート管理", expected: false, row: {"日付": addDaysForSample_(today, -3), "経理確認": "未確認", "精算状態": "未精算"}}
  ];

  const results = cases.map(c => {
    const actual = shouldMoveToArchive_(c.row, c.sheetName, today);
    return {
      "チェック項目": c.name,
      "対象シート": c.sheetName,
      "期待値": c.expected ? "移動する" : "移動しない",
      "実結果": actual ? "移動する" : "移動しない",
      "判定": actual === c.expected ? "OK" : "NG",
      "確認日時": new Date()
    };
  });

  writeCheckResultSheet_("機能チェック結果", ["チェック項目", "対象シート", "期待値", "実結果", "判定", "確認日時"], results);

  const ngCount = results.filter(r => r["判定"] === "NG").length;
  SpreadsheetApp.getActiveSpreadsheet().toast("過去移動ルールチェック完了: NG " + ngCount + "件");
}

/**
 * 本番移行前に最低限見る項目をチェックシートへ出します。
 */
function runPreMigrationCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headersMap = getSheetHeaders_();
  const results = [];

  Object.keys(headersMap).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      results.push(buildPreMigrationResult_(sheetName, "シート存在", "NG", "シートがありません"));
      return;
    }

    const expectedHeaders = headersMap[sheetName];
    const actualHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expectedHeaders.length)).getValues()[0];
    const missing = expectedHeaders.filter(h => !actualHeaders.includes(h));
    results.push(buildPreMigrationResult_(sheetName, "ヘッダー確認", missing.length === 0 ? "OK" : "NG", missing.length === 0 ? "必要列あり" : "不足: " + missing.join(" / ")));

    const idHeader = getIdHeaderForSheet_(sheetName);
    if (idHeader) {
      results.push(buildPreMigrationResult_(sheetName, "ID列確認", actualHeaders.includes(idHeader) ? "OK" : "NG", idHeader));
    }
  });

  const dailyHeaders = headersMap["日報"] || [];
  results.push(buildPreMigrationResult_("日報", "PDFリンク列確認", dailyHeaders.includes("PDFリンク") ? "OK" : "NG", "PDFリンクは既読グループ外に置く"));
  results.push(buildPreMigrationResult_("電話履歴", "状態/電話対応整理", (headersMap["電話履歴"] || []).includes("対応メモ") ? "OK" : "NG", "電話履歴は電話対応に一本化、旧状態は対応メモ"));

  writeCheckResultSheet_("移行前チェック", ["対象シート", "チェック項目", "判定", "備考", "確認日時"], results);
  SpreadsheetApp.getActiveSpreadsheet().toast("移行前チェックを作成しました");
}

function buildPreMigrationResult_(sheetName, item, result, note) {
  return {
    "対象シート": sheetName,
    "チェック項目": item,
    "判定": result,
    "備考": note,
    "確認日時": new Date()
  };
}

function writeCheckResultSheet_(sheetName, headers, objects) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  resetSheet(sheet, headers.length);
  setupHeaderOnly_(sheet, headers);

  if (objects.length > 0) {
    const values = objects.map(obj => headers.map(h => obj[h] !== undefined ? obj[h] : ""));
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }

  sheet.autoResizeColumns(1, headers.length);
}
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
    ["vehicles", "車検管理", "next_insurance_due_date", "DATE", "", "", "NULL", "次回保険期限"],
    ["vehicles", "車検管理", "status", "TEXT", "", "", "NULL", "状態"],
    ["vehicles", "車検管理", "photo_url", "TEXT", "", "", "NULL", "写真URL"],
    ["vehicles", "車検管理", "remarks", "TEXT", "", "", "NULL", "備考"],

    ["vehicle_inspection_logs", "車検履歴", "id", "INTEGER", "PK", "", "NOT NULL", "車検履歴ID"],
    ["vehicle_inspection_logs", "車検履歴", "vehicle_id", "INTEGER", "", "vehicles.id", "NOT NULL", "車両ID"],
    ["vehicle_inspection_logs", "車検履歴", "old_inspection_due_date", "DATE", "", "", "NOT NULL", "旧車検期限"],
    ["vehicle_inspection_logs", "車検履歴", "new_inspection_due_date", "DATE", "", "", "NOT NULL", "新車検期限"],
    ["vehicle_inspection_logs", "車検履歴", "old_insurance_due_date", "DATE", "", "", "NULL", "旧保険期限"],
    ["vehicle_inspection_logs", "車検履歴", "new_insurance_due_date", "DATE", "", "", "NULL", "新保険期限"],
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

    ["一覧スケジュール", "既読確認者列", "read_checks", "is_read", "横持ち列を縦持ちに変換", "例：山田/高橋/鈴木列を行データにする"],
    ["一覧スケジュール", "既読確認者列名", "read_checks", "staff_id", "列名をstaffに紐づける", "DBでは列追加ではなく行追加で管理"],

    ["車検管理", "車両名", "vehicles", "vehicle_name", "TEXT", "車両マスタとして扱う"],
    ["車検管理", "車番", "vehicles", "vehicle_number", "TEXT", "数値化しない"],
    ["車検管理", "車検期限", "vehicles", "inspection_due_date", "DATE型へ変換", "現在の車検期限"],
    ["車検管理", "次回車検期限", "vehicles", "next_inspection_due_date", "DATE型へ変換", "更新予定日"],
    ["車検管理", "保険期限", "vehicles", "insurance_due_date", "DATE型へ変換", "保険期限"],
    ["車検管理", "次回保険期限", "vehicles", "next_insurance_due_date", "DATE型へ変換", "更新予定の保険期限"],
    ["車検管理", "状態", "vehicles", "status", "TEXT", "未対応・予約済・実施済・完了・更新済"],
    ["車検管理", "写真", "vehicles", "photo_url", "URL文字列", "写真リンク"],
    ["車検管理", "備考", "vehicles", "remarks", "TEXT", "備考"],

    ["車検履歴", "旧車検期限", "vehicle_inspection_logs", "old_inspection_due_date", "DATE型へ変換", "更新前の期限"],
    ["車検履歴", "新車検期限", "vehicle_inspection_logs", "new_inspection_due_date", "DATE型へ変換", "更新後の期限"],
    ["車検履歴", "旧保険期限", "vehicle_inspection_logs", "old_insurance_due_date", "DATE型へ変換", "更新前の保険期限"],
    ["車検履歴", "新保険期限", "vehicle_inspection_logs", "new_insurance_due_date", "DATE型へ変換", "更新後の保険期限"],
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
      "/*\nSheets:\n  schedule_id / 内容 / 山田 / 高橋 / 鈴木\n\nRDB:\n  schedule_id / staff_id / is_read\n\n例:\nINSERT INTO read_checks (schedule_id, staff_id, is_read)\nVALUES\n  (1, 1, 1),\n  (1, 2, 0),\n  (1, 3, 1);\n*/",
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

  createFilterSafely_(sheet, headers.length);

  sheet.autoResizeColumns(1, headers.length);
}



/**
 * 入力用・確認用シートの表示幅をまとめて調整する。
 *
 * - すべての入力用シートと主要な確認用シートを対象にする
 * - 列幅は固定、折り返し表示ON、行高さは少し大きめにする
 * - onEditのたびには実行せず、メニューまたはrefreshAll後だけ実行する
 * - autoResizeColumnsで列幅が戻る問題を避けるため、この関数で最終的な表示幅を決める
 */
function formatInputSheetsForLongText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targets = [
    "出先予定",
    "工事予定",
    "会議予定",
    "行事予定",
    "作業状況",
    "電話履歴",
    "車検管理",
    "車検履歴",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理",
    "お知らせ",
    "個人ToDo",
    "一覧スケジュール",
    "要確認一覧",
    "過去一覧",
    "帳簿PDF履歴",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集"
  ];

  targets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    formatSheetBase_(sheet);
    formatSheetByName_(sheet);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("入力用・確認用シートの表示幅を調整しました");
}

/**
 * 旧メニュー互換用。
 * メニュー「日報シートの表示幅を調整」から呼ばれる。
 */
function formatDailyReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");
  if (!sheet) return;

  formatSheetBase_(sheet);
  formatSheetByName_(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast("日報シートの表示幅を調整しました");
}

/**
 * 表示調整の共通処理。
 */
function formatSheetBase_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const maxRows = sheet.getMaxRows();
  const formatRows = Math.max(maxRows, 50);

  sheet.getRange(1, 1, formatRows, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);

  const bodyRows = Math.min(Math.max(maxRows - 1, 1), 200);
  if (bodyRows > 0) {
    sheet.setRowHeights(2, bodyRows, 42);
  }
}

/**
 * シート名ごとの列幅調整。
 */
function formatSheetByName_(sheet) {
  const name = sheet.getName();
  const headers = getHeadersForFormat_(sheet);
  if (headers.length === 0) return;

  function setWidth(header, width) {
    setColumnWidthByHeader_(sheet, headers, header, width);
  }

  function setPersonalCheckWidth() {
    setWidth(READ_HEADER, 70);
    getPersonalMembers_().forEach(member => {
      setWidth(member, 70);
    });
  }

  if (name === "車検管理") {
    setWidth("車両名", 130);
    setWidth("車番", 100);
    setWidth("車検期限", 110);
    setWidth("次回車検期限", 130);
    setWidth("通知", 100);
    setWidth("保険期限", 110);
    setWidth("次回保険期限", 130);
    setWidth("状態", 100);
    setWidth("写真", 150);
    setWidth("備考", 220);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 42);
    return;
  }

  if (name === "車検履歴") {
    setWidth("更新日", 110);
    setWidth("車両名", 130);
    setWidth("車番", 100);
    setWidth("旧車検期限", 120);
    setWidth("新車検期限", 120);
    setWidth("旧保険期限", 120);
    setWidth("新保険期限", 120);
    setWidth("担当", 100);
    setWidth("備考", 240);
    setBodyRowHeight_(sheet, 42);
    return;
  }

  if (name === "日報") {
    setWidth("日付", 110);
    setWidth("入力者", 100);
    setWidth("担当", 100);
    setWidth("現場", 160);
    setWidth("作業内容", 260);
    setWidth("進捗", 120);
    setWidth("問題点", 260);
    setWidth("明日の予定", 260);
    setWidth("他現場状況", 360);
    setWidth("写真", 150);
    setWidth("日報文章", 420);
    setWidth("PDFリンク", 220);
    setWidth("状態", 100);
    setWidth("備考", 220);
    setWidth("日報ID", 150);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 70);
    return;
  }


  if (name === "日報レシート管理") {
    setWidth("日付", 110);
    setWidth("担当", 100);
    setWidth("現場", 180);
    setWidth("支払先", 180);
    setWidth("内容", 260);
    setWidth("区分", 100);
    setWidth("金額", 100);
    setWidth("支払方法", 110);
    setWidth("レシート写真", 170);
    setWidth("経理確認", 110);
    setWidth("精算状態", 110);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setWidth("レシートID", 150);
    setBodyRowHeight_(sheet, 58);
    return;
  }

  if (name === "出先予定") {
    setWidth("日付", 110);
    setWidth("行き先", 180);
    setWidth("用件", 280);
    setWidth("担当", 100);
    setWidth("社用車", 120);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("電話対応", 110);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "工事予定") {
    setWidth("工事名", 220);
    setWidth("現場", 180);
    setWidth("開始日", 110);
    setWidth("終了日", 110);
    setWidth("状態", 100);
    setWidth("担当", 100);
    setWidth("通知", 100);
    setWidth("電話対応", 110);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "会議予定") {
    setWidth("日付", 110);
    setWidth("会議名", 200);
    setWidth("内容", 320);
    setWidth("担当", 100);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("資料", 160);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "行事予定") {
    setWidth("開始日", 110);
    setWidth("終了日", 110);
    setWidth("行事名", 220);
    setWidth("内容", 320);
    setWidth("担当", 100);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "作業状況") {
    setWidth("現場", 180);
    setWidth("作業内容", 320);
    setWidth("状態", 100);
    setWidth("担当", 100);
    setWidth("写真", 150);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 50);
    return;
  }

  if (name === "電話履歴") {
    setWidth("日時", 130);
    setWidth("相手", 160);
    setWidth("内容", 320);
    setWidth("電話対応", 110);
    setWidth("担当", 100);
    setWidth("対応メモ", 220);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setWidth("電話ID", 150);
        setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 50);
    return;
  }

  if (name === "社用車予約") {
    setWidth("日付", 110);
    setWidth("開始時刻", 90);
    setWidth("終了時刻", 90);
    setWidth("社用車", 130);
    setWidth("利用者", 100);
    setWidth("行き先", 180);
    setWidth("用途", 240);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "免許資格管理") {
    setWidth("担当", 100);
    setWidth("資格名", 220);
    setWidth("区分", 120);
    setWidth("取得日", 110);
    setWidth("更新期限", 110);
    setWidth("コピー有無", 100);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "備品修理管理") {
    setWidth("登録日", 110);
    setWidth("備品名", 180);
    setWidth("場所", 160);
    setWidth("内容", 320);
    setWidth("担当", 100);
    setWidth("修理依頼日", 120);
    setWidth("返却予定日", 120);
    setWidth("返却済", 90);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 50);
    return;
  }

  if (name === "お知らせ") {
    setWidth("投稿日", 110);
    setWidth("タイトル", 240);
    setWidth("内容", 420);
    setWidth("投稿者", 100);
    setWidth("重要度", 90);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 56);
    return;
  }

  if (name === "個人ToDo") {
    setWidth("登録日", 110);
    setWidth("担当", 100);
    setWidth("内容", 360);
    setWidth("期限", 110);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("備考", 240);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 50);
    return;
  }

  if (name === "一覧スケジュール") {
    setWidth("日付", 110);
    setWidth("種類", 110);
    setWidth("内容", 360);
    setWidth("担当", 100);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("電話対応", 110);
    setWidth("備考", 240);
    setWidth("元シート", 130);
    setPersonalCheckWidth();
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "要確認一覧") {
    setWidth("日付", 110);
    setWidth("種類", 110);
    setWidth("内容", 380);
    setWidth("担当", 100);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("電話対応", 110);
    setWidth("備考", 240);
    setWidth("元シート", 130);
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "過去一覧") {
    setWidth("日付", 110);
    setWidth("種類", 110);
    setWidth("内容", 360);
    setWidth("担当", 100);
    setWidth("状態", 100);
    setWidth("通知", 100);
    setWidth("電話対応", 110);
    setWidth("元シート", 130);
    setWidth("備考", 240);
    setWidth("移動日", 110);
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "帳簿PDF履歴") {
    setWidth("作成日時", 150);
    setWidth("対象", 160);
    setWidth("期間", 160);
    setWidth("PDFリンク", 220);
    setWidth("備考", 260);
    setBodyRowHeight_(sheet, 46);
    return;
  }

  if (name === "SQL設計") {
    setWidth("分類", 130);
    setWidth("テーブル名", 190);
    setWidth("カラム名", 190);
    setWidth("型", 120);
    setWidth("説明", 460);
    setBodyRowHeight_(sheet, 56);
    return;
  }

  if (name === "移行対応表") {
    setWidth("現行シート", 160);
    setWidth("現行列", 160);
    setWidth("移行先テーブル", 190);
    setWidth("移行先カラム", 190);
    setWidth("移行メモ", 460);
    setBodyRowHeight_(sheet, 56);
    return;
  }

  if (name === "SQLサンプル集") {
    setWidth("用途", 220);
    setWidth("SQL", 640);
    setWidth("説明", 420);
    setBodyRowHeight_(sheet, 80);
    return;
  }
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
 * 本文行の高さをまとめて変更する。
 */
function setBodyRowHeight_(sheet, height) {
  const maxRows = sheet.getMaxRows();
  const bodyRows = Math.min(Math.max(maxRows - 1, 1), 200);
  if (bodyRows > 0) {
    sheet.setRowHeights(2, bodyRows, height);
  }
}

/**
 * 旧関数互換用。
 * 古い個別フォーマット関数が残っていても動くようにする。
 */
function applyReadableLongTextFormat_(sheet, bodyRowHeight) {
  if (!sheet) return;
  formatSheetBase_(sheet);
  setBodyRowHeight_(sheet, bodyRowHeight || 42);
}

/**
 * v8.5 表示幅調整 強制上書き版
 *
 * 既存の formatInputSheetsForLongText / formatDailyReportSheet /
 * formatSheetBase_ / formatSheetByName_ が上に残っていても、
 * Apps Script では後から定義した関数が優先されるため、
 * このブロックをファイル末尾に置くことで表示幅ルールを統一する。
 *
 * 方針：
 * - 文章列だけ広め
 * - 日付・担当・状態・通知・チェック列は小さめ
 * - 入力シート、一覧系、SQL設計系、集計系もまとめて調整
 */
function formatInputSheetsForLongText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    "出先予定",
    "工事予定",
    "会議予定",
    "行事予定",
    "作業状況",
    "電話履歴",
    "車検管理",
    "車検履歴",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理",
    "お知らせ",
    "個人ToDo",
    "一覧スケジュール",
    "要確認一覧",
    "過去一覧",
    "帳簿PDF履歴",
    "帳簿出力設定",
    "設定",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集",
    "担当別未読",
    "既読率集計",
    "ダッシュボード"
  ];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    formatSheetBase_(sheet);
    formatSheetByName_(sheet);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("全シートの表示幅を調整しました");
}

/**
 * メニュー用：全シート表示幅を強制調整
 */
function formatAllSheetsFixedWidthsNow() {
  formatInputSheetsForLongText();
}

/**
 * 日報だけ表示幅を調整する。
 */
function formatDailyReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");

  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("日報シートが見つかりません");
    return;
  }

  formatSheetBase_(sheet);
  formatSheetByName_(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast("日報シートの表示幅を調整しました");
}

/**
 * 全シート共通の基本表示。
 */
function formatSheetBase_(sheet) {
  const lastCol = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();

  if (lastCol < 1 || maxRows < 1) return;

  sheet.getRange(1, 1, maxRows, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.getRange(1, 1, 1, lastCol)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");

  const rowCount = Math.min(Math.max(sheet.getLastRow() - 1, 20), Math.max(maxRows - 1, 1));
  if (maxRows >= 2 && rowCount > 0) {
    sheet.setRowHeights(2, rowCount, 34);
  }
}

/**
 * シート別の列幅設定。
 */
function formatSheetByName_(sheet) {
  const name = sheet.getName();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function setWidth(header, width) {
    const col = headers.indexOf(header) + 1;
    if (col > 0) {
      sheet.setColumnWidth(col, width);
    }
  }

  function setCheckWidths() {
    setWidth(READ_HEADER, 60);

    getPersonalMembers_().forEach(member => {
      setWidth(member, 60);
    });
  }

  // 共通で短くする列
  setWidth("日付", 105);
  setWidth("日時", 125);
  setWidth("開始日", 105);
  setWidth("終了日", 105);
  setWidth("登録日", 105);
  setWidth("投稿日", 105);
  setWidth("期限", 105);
  setWidth("取得日", 105);
  setWidth("更新期限", 105);
  setWidth("修理依頼日", 110);
  setWidth("返却予定日", 110);

  setWidth("担当", 80);
  setWidth("利用者", 80);
  setWidth("状態", 95);
  setWidth("通知", 90);
  setWidth("電話対応", 95);
  setWidth("重要度", 80);
  setWidth("コピー有無", 90);
  setWidth("返却済", 75);
  setWidth("PDFリンク", 220);
  Object.keys(SHEET_ID_HEADERS).forEach(sheetName => {
    setWidth(SHEET_ID_HEADERS[sheetName], 150);
  });

  setCheckWidths();

  if (name === "出先予定") {
    setWidth("日付", 105);
    setWidth("行き先", 150);
    setWidth("用件", 240);
    setWidth("担当", 80);
    setWidth("社用車", 105);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("対応メモ", 220);
    setWidth("備考", 220);
    setWidth("電話ID", 150);
    return;
  }

  if (name === "工事予定") {
    setWidth("工事名", 180);
    setWidth("現場", 145);
    setWidth("開始日", 105);
    setWidth("終了日", 105);
    setWidth("状態", 95);
    setWidth("担当", 80);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("備考", 230);
    return;
  }

  if (name === "会議予定") {
    setWidth("日付", 105);
    setWidth("会議名", 170);
    setWidth("内容", 280);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("資料", 150);
    setWidth("備考", 220);
    return;
  }

  if (name === "行事予定") {
    setWidth("開始日", 105);
    setWidth("終了日", 105);
    setWidth("行事名", 170);
    setWidth("内容", 280);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "作業状況") {
    setWidth("現場", 150);
    setWidth("作業内容", 280);
    setWidth("状態", 95);
    setWidth("担当", 80);
    setWidth("写真", 130);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "電話履歴") {
    setWidth("日時", 125);
    setWidth("相手", 140);
    setWidth("内容", 280);
    setWidth("電話対応", 95);
    setWidth("担当", 80);
    setWidth("対応メモ", 220);
    setWidth("通知", 90);
    setWidth("備考", 220);
    setWidth("電話ID", 150);
    return;
  }

  if (name === "車検管理") {
    setWidth("車両名", 110);
    setWidth("車番", 85);
    setWidth("車検期限", 105);
    setWidth("次回車検期限", 115);
    setWidth("通知", 90);
    setWidth("保険期限", 105);
    setWidth("次回保険期限", 115);
    setWidth("状態", 95);
    setWidth("写真", 130);
    setWidth("備考", 210);
    return;
  }

  if (name === "車検履歴") {
    setWidth("更新日", 105);
    setWidth("車両名", 110);
    setWidth("車番", 85);
    setWidth("旧車検期限", 110);
    setWidth("新車検期限", 110);
    setWidth("旧保険期限", 110);
    setWidth("新保険期限", 110);
    setWidth("担当", 80);
    setWidth("備考", 220);
    return;
  }

  if (name === "社用車予約") {
    setWidth("日付", 105);
    setWidth("開始時刻", 80);
    setWidth("終了時刻", 80);
    setWidth("社用車", 115);
    setWidth("利用者", 80);
    setWidth("行き先", 150);
    setWidth("用途", 220);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "日報") {
    setWidth("日付", 105);
    setWidth("担当", 80);
    setWidth("現場", 145);
    setWidth("作業内容", 260);
    setWidth("進捗", 100);
    setWidth("問題点", 260);
    setWidth("明日の予定", 260);
    setWidth("写真", 130);
    setWidth("日報文章", 390);
    setWidth("状態", 95);
    setWidth("備考", 220);
    setWidth("日報ID", 150);

    if (sheet.getMaxRows() >= 2) {
      sheet.setRowHeights(2, sheet.getMaxRows() - 1, 64);
    }
    return;
  }

  if (name === "免許資格管理") {
    setWidth("担当", 80);
    setWidth("資格名", 180);
    setWidth("区分", 90);
    setWidth("取得日", 105);
    setWidth("更新期限", 105);
    setWidth("コピー有無", 90);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 210);
    return;
  }

  if (name === "備品修理管理") {
    setWidth("登録日", 105);
    setWidth("備品名", 150);
    setWidth("場所", 130);
    setWidth("内容", 280);
    setWidth("担当", 80);
    setWidth("修理依頼日", 110);
    setWidth("返却予定日", 110);
    setWidth("返却済", 75);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "お知らせ") {
    setWidth("投稿日", 105);
    setWidth("タイトル", 210);
    setWidth("内容", 360);
    setWidth("投稿者", 80);
    setWidth("重要度", 80);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "個人ToDo") {
    setWidth("登録日", 105);
    setWidth("担当", 80);
    setWidth("内容", 320);
    setWidth("期限", 105);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "一覧スケジュール") {
    setWidth("日付", 105);
    setWidth("種類", 105);
    setWidth("内容", 330);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("備考", 220);
    setWidth("元シート", 120);
    return;
  }

  if (name === "要確認一覧") {
    setWidth("日付", 105);
    setWidth("種類", 105);
    setWidth("内容", 330);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("備考", 220);
    setWidth("元シート", 120);
    return;
  }

  if (name === "過去一覧") {
    setWidth("日付", 105);
    setWidth("種類", 105);
    setWidth("内容", 330);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("元シート", 120);
    setWidth("備考", 220);
    setWidth("移動日", 105);
    return;
  }

  if (name === "帳簿PDF履歴") {
    setWidth("作成日時", 140);
    setWidth("対象", 140);
    setWidth("期間", 160);
    setWidth("PDFリンク", 220);
    setWidth("備考", 220);
    return;
  }

  if (name === "設定") {
    setWidth("設定種別", 110);
    setWidth("値", 140);
    setWidth("備考", 360);
    return;
  }

  if (name === "SQL設計") {
    setWidth("分類", 130);
    setWidth("テーブル名", 160);
    setWidth("項目名", 160);
    setWidth("型", 110);
    setWidth("説明", 360);
    setWidth("備考", 260);
    return;
  }

  if (name === "移行対応表") {
    setWidth("Google Sheetsシート", 160);
    setWidth("Google Sheets列", 150);
    setWidth("SQLテーブル", 160);
    setWidth("SQLカラム", 160);
    setWidth("変換方針", 360);
    setWidth("備考", 260);
    return;
  }

  if (name === "SQLサンプル集") {
    setWidth("用途", 190);
    setWidth("SQL例", 560);
    setWidth("説明", 320);
    return;
  }

  if (name === "担当別未読") {
    setWidth("担当", 90);
    setWidth("未読件数", 90);
    setWidth("未完了ToDo", 110);
    setWidth("未対応電話", 110);
    setWidth("今日の予定", 100);
    return;
  }

  if (name === "既読率集計") {
    setWidth("対象", 140);
    setWidth("全件数", 90);
    setWidth("既読件数", 90);
    setWidth("未読件数", 90);
    setWidth("既読率", 90);
    return;
  }

  if (name === "ダッシュボード") {
    setWidth("項目", 180);
    setWidth("件数", 90);
    return;
  }
}

/**
 * v8.6 表示幅・ダッシュボード最終調整版
 *
 * - 日付/日時列をさらに小さめに調整
 * - ダッシュボードは入力シートとは別レイアウトで整形
 * - createDashboard() 直後にもダッシュボード専用表示調整を実行
 * - このブロックはファイル末尾に置く。Apps Scriptでは後から定義した関数が優先される。
 */
function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName("ダッシュボード");
  if (!sheet) {
    sheet = ss.insertSheet("ダッシュボード");
  }

  if (sheet.getMaxColumns() < 4) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 4 - sheet.getMaxColumns());
  }

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
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.setConditionalFormatRules([]);

  // タイトルはA〜Dを結合して、縦に割れないようにする。
  sheet.getRange("A1:D1")
    .merge()
    .setValue("社内共有ダッシュボード")
    .setBackground("#d9ead3")
    .setFontWeight("bold")
    .setFontSize(22)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setWrap(false);

  sheet.getRange("A3:B3").setValues([["項目", "件数"]]);
  sheet.getRange("A3:B3")
    .setBackground("#d9ead3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

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

  sheet.setFrozenRows(3);

  try {
    formatSheetByName_(sheet);
  } catch (e) {
    console.log("ダッシュボード表示調整をスキップ: " + e.message);
  }
}

function formatInputSheetsForLongText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    "出先予定",
    "工事予定",
    "会議予定",
    "行事予定",
    "作業状況",
    "電話履歴",
    "車検管理",
    "車検履歴",
    "社用車予約",
    "日報",
    "免許資格管理",
    "備品修理管理",
    "お知らせ",
    "個人ToDo",
    "一覧スケジュール",
    "要確認一覧",
    "過去一覧",
    "帳簿PDF履歴",
    "帳簿出力設定",
    "設定",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集",
    "担当別未読",
    "既読率集計",
    "ダッシュボード"
  ];

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    formatSheetBase_(sheet);
    formatSheetByName_(sheet);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("全シートの表示幅を調整しました");
}

function formatAllSheetsFixedWidthsNow() {
  formatInputSheetsForLongText();
}

function formatDailyReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("日報");

  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("日報シートが見つかりません");
    return;
  }

  formatSheetBase_(sheet);
  formatSheetByName_(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast("日報シートの表示幅を調整しました");
}

function formatSheetBase_(sheet) {
  const lastCol = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();

  if (lastCol < 1 || maxRows < 1) return;

  sheet.getRange(1, 1, maxRows, lastCol)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.getRange(1, 1, 1, lastCol)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold");

  const rowCount = Math.min(Math.max(sheet.getLastRow() - 1, 20), Math.max(maxRows - 1, 1));
  if (maxRows >= 2 && rowCount > 0) {
    sheet.setRowHeights(2, rowCount, 34);
  }
}

function formatSheetByName_(sheet) {
  const name = sheet.getName();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  function setWidth(header, width) {
    const col = headers.indexOf(header) + 1;
    if (col > 0) {
      sheet.setColumnWidth(col, width);
    }
  }

  function setCheckWidths() {
    setWidth(READ_HEADER, 60);
    getPersonalMembers_().forEach(member => {
      setWidth(member, 60);
    });
  }

  // 共通で短くする列。日付は yyyy/mm/dd が読める最小寄りにする。
  setWidth("日付", 90);
  setWidth("日時", 115);
  setWidth("開始日", 90);
  setWidth("終了日", 90);
  setWidth("登録日", 90);
  setWidth("投稿日", 90);
  setWidth("期限", 90);
  setWidth("取得日", 90);
  setWidth("更新期限", 95);
  setWidth("修理依頼日", 100);
  setWidth("返却予定日", 100);

  setWidth("担当", 80);
  setWidth("利用者", 80);
  setWidth("状態", 95);
  setWidth("通知", 90);
  setWidth("電話対応", 95);
  setWidth("重要度", 80);
  setWidth("コピー有無", 90);
  setWidth("返却済", 75);
  setWidth("PDFリンク", 220);
  Object.keys(SHEET_ID_HEADERS).forEach(sheetName => {
    setWidth(SHEET_ID_HEADERS[sheetName], 150);
  });

  setCheckWidths();

  if (name === "出先予定") {
    setWidth("日付", 90);
    setWidth("行き先", 150);
    setWidth("用件", 240);
    setWidth("担当", 80);
    setWidth("社用車", 105);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("対応メモ", 220);
    setWidth("備考", 220);
    setWidth("電話ID", 150);
    return;
  }

  if (name === "工事予定") {
    setWidth("工事名", 180);
    setWidth("現場", 145);
    setWidth("開始日", 90);
    setWidth("終了日", 90);
    setWidth("状態", 95);
    setWidth("担当", 80);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("備考", 230);
    return;
  }

  if (name === "会議予定") {
    setWidth("日付", 90);
    setWidth("会議名", 170);
    setWidth("内容", 280);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("資料", 150);
    setWidth("備考", 220);
    return;
  }

  if (name === "行事予定") {
    setWidth("開始日", 90);
    setWidth("終了日", 90);
    setWidth("行事名", 170);
    setWidth("内容", 280);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "作業状況") {
    setWidth("現場", 150);
    setWidth("作業内容", 280);
    setWidth("状態", 95);
    setWidth("担当", 80);
    setWidth("写真", 130);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "電話履歴") {
    setWidth("日時", 115);
    setWidth("相手", 140);
    setWidth("内容", 280);
    setWidth("電話対応", 95);
    setWidth("担当", 80);
    setWidth("対応メモ", 220);
    setWidth("通知", 90);
    setWidth("備考", 220);
    setWidth("電話ID", 150);
    return;
  }

  if (name === "車検管理") {
    setWidth("車両名", 110);
    setWidth("車番", 85);
    setWidth("車検期限", 95);
    setWidth("次回車検期限", 105);
    setWidth("通知", 90);
    setWidth("保険期限", 95);
    setWidth("次回保険期限", 105);
    setWidth("状態", 95);
    setWidth("写真", 130);
    setWidth("備考", 210);
    return;
  }

  if (name === "車検履歴") {
    setWidth("更新日", 90);
    setWidth("車両名", 110);
    setWidth("車番", 85);
    setWidth("旧車検期限", 105);
    setWidth("新車検期限", 105);
    setWidth("旧保険期限", 105);
    setWidth("新保険期限", 105);
    setWidth("担当", 80);
    setWidth("備考", 220);
    return;
  }

  if (name === "社用車予約") {
    setWidth("日付", 90);
    setWidth("開始時刻", 80);
    setWidth("終了時刻", 80);
    setWidth("社用車", 115);
    setWidth("利用者", 80);
    setWidth("行き先", 150);
    setWidth("用途", 220);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "日報") {
    // PDFは専用の一時シートで整形するため、入力用の日報シートは横に長くしすぎない。
    setWidth("日付", 85);
    setWidth("入力者", 80);
    setWidth("担当", 80);
    setWidth("現場", 120);
    setWidth("作業内容", 220);
    setWidth("進捗", 80);
    setWidth("問題点", 220);
    setWidth("明日の予定", 220);
    setWidth("他現場状況", 240);
    setWidth("写真", 115);
    setWidth("日報文章", 300);
    setWidth("PDFリンク", 170);
    setWidth("状態", 90);
    setWidth("備考", 180);
    setWidth("日報ID", 125);

    const rowCount = Math.min(Math.max(sheet.getLastRow() - 1, 20), Math.max(sheet.getMaxRows() - 1, 1));
    if (sheet.getMaxRows() >= 2 && rowCount > 0) {
      sheet.setRowHeights(2, rowCount, 72);
    }
    return;
  }

  if (name === "免許資格管理") {
    setWidth("担当", 80);
    setWidth("資格名", 180);
    setWidth("区分", 90);
    setWidth("取得日", 90);
    setWidth("更新期限", 95);
    setWidth("コピー有無", 90);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 210);
    return;
  }

  if (name === "備品修理管理") {
    setWidth("登録日", 90);
    setWidth("備品名", 150);
    setWidth("場所", 130);
    setWidth("内容", 280);
    setWidth("担当", 80);
    setWidth("修理依頼日", 100);
    setWidth("返却予定日", 100);
    setWidth("返却済", 75);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "お知らせ") {
    setWidth("投稿日", 90);
    setWidth("タイトル", 210);
    setWidth("内容", 360);
    setWidth("投稿者", 80);
    setWidth("重要度", 80);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "個人ToDo") {
    setWidth("登録日", 90);
    setWidth("担当", 80);
    setWidth("内容", 320);
    setWidth("期限", 90);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("備考", 220);
    return;
  }

  if (name === "一覧スケジュール") {
    setWidth("日付", 90);
    setWidth("種類", 105);
    setWidth("内容", 330);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("備考", 220);
    setWidth("元シート", 120);
    return;
  }

  if (name === "要確認一覧") {
    setWidth("日付", 90);
    setWidth("種類", 105);
    setWidth("内容", 330);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("備考", 220);
    setWidth("元シート", 120);
    return;
  }

  if (name === "過去一覧") {
    setWidth("日付", 90);
    setWidth("種類", 105);
    setWidth("内容", 330);
    setWidth("担当", 80);
    setWidth("状態", 95);
    setWidth("通知", 90);
    setWidth("電話対応", 95);
    setWidth("元シート", 120);
    setWidth("備考", 220);
    setWidth("移動日", 90);
    return;
  }

  if (name === "帳簿PDF履歴") {
    setWidth("作成日時", 130);
    setWidth("対象", 140);
    setWidth("期間", 140);
    setWidth("PDFリンク", 320);
    setWidth("備考", 220);
    return;
  }

  if (name === "設定") {
    setWidth("設定種別", 110);
    setWidth("値", 140);
    setWidth("備考", 360);
    return;
  }

  if (name === "SQL設計") {
    setWidth("分類", 130);
    setWidth("テーブル名", 160);
    setWidth("項目名", 160);
    setWidth("型", 110);
    setWidth("説明", 360);
    setWidth("備考", 260);
    return;
  }

  if (name === "移行対応表") {
    setWidth("Google Sheetsシート", 160);
    setWidth("Google Sheets列", 150);
    setWidth("SQLテーブル", 160);
    setWidth("SQLカラム", 160);
    setWidth("変換方針", 360);
    setWidth("備考", 260);
    return;
  }

  if (name === "SQLサンプル集") {
    setWidth("用途", 190);
    setWidth("SQL例", 560);
    setWidth("説明", 320);
    return;
  }

  if (name === "担当別未読") {
    setWidth("担当", 90);
    setWidth("未読件数", 90);
    setWidth("未完了ToDo", 110);
    setWidth("未対応電話", 110);
    setWidth("今日の予定", 100);
    return;
  }

  if (name === "既読率集計") {
    setWidth("対象", 140);
    setWidth("全件数", 90);
    setWidth("既読件数", 90);
    setWidth("未読件数", 90);
    setWidth("既読率", 90);
    return;
  }

  if (name === "ダッシュボード") {
    sheet.setColumnWidth(1, 240);
    sheet.setColumnWidth(2, 90);
    sheet.setColumnWidth(3, 20);
    sheet.setColumnWidth(4, 300);

    try {
      sheet.getRange("A1:D1")
        .setWrap(false)
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("left");
    } catch (e) {}

    sheet.setRowHeight(1, 90);
    sheet.setRowHeight(2, 24);
    sheet.setRowHeight(3, 42);

    if (sheet.getMaxRows() >= 4) {
      sheet.setRowHeights(4, sheet.getMaxRows() - 3, 38);
    }
    return;
  }
}












/**
 * v9.6.4 追加修正（onEdit軽量化・日時順ソート・色分け整理）
 *
 * このブロックは既存関数を安全に上書きします。
 * - onEditでは一覧/要確認/集計を再作成しない
 * - 全体更新（軽量）で入力シートを日付順に並び替えてから集約する
 * - 状態/通知/種類の色を意味ごとに整理する
 * - ロック待ちを短くし、スキップ時のメッセージを分かりやすくする
 */

function markSummaryDirty_() {
  try {
    PropertiesService.getDocumentProperties().setProperty(
      "SUMMARY_DIRTY",
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")
    );
  } catch (e) {}
}

function clearSummaryDirty_() {
  try {
    PropertiesService.getDocumentProperties().deleteProperty("SUMMARY_DIRTY");
  } catch (e) {}
}

function onEdit(e) {
  if (!e || !e.range) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) {
    // 編集中の二重実行を防ぐ。頻繁に出る場合は、少し待ってから手動更新する。
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("別の処理中のため、この編集時処理をスキップしました。必要なら少し待って全体更新（軽量）を実行してください。");
    } catch (err) {}
    return;
  }

  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (row === 1) return;

    ensureRowsAfterEditLight_(sheet, row);

    if (e.range.getValue() === CLEAR_LABEL) {
      e.range.clearContent();
      markSummaryDirty_();
      return;
    }

    ensureIdForEditedRow_(sheet, row);
    updateNoticeForEditedRow(sheet, row, col);
    updateCheckboxesForEditedRow(sheet, row);

    if (sheetName === "社用車予約") {
      const conflictMessage = getCarReservationConflictMessage(sheet, row);
      if (conflictMessage) {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const noticeCol = headers.indexOf("通知") + 1;
        if (noticeCol > 0) sheet.getRange(row, noticeCol).setValue("予約重複");
        try {
          SpreadsheetApp.getActiveSpreadsheet().toast(conflictMessage);
        } catch (err) {}
      }
    }

    if (getInputSheetNames_().includes(sheetName)) {
      // v9.6.4：onEditでは重い一覧/要確認/集計作成を走らせない。
      // 入力後に「全体更新（軽量）」でまとめて反映する。
      markSummaryDirty_();
    }
  } finally {
    lock.releaseLock();
  }
}

function ensureRowsAfterEditLight_(sheet, row) {
  if (!sheet || row <= 1) return;

  const maxRows = sheet.getMaxRows();
  const remainingRows = maxRows - row;
  if (remainingRows > AUTO_EXTEND_TRIGGER_MARGIN) return;

  sheet.insertRowsAfter(maxRows, AUTO_EXTEND_ROWS_BUFFER);

  // v9.6.4：行追加時に全ヘッダー/全入力規則を再設定しない。
  // 追加行に最低限の入力規則を必要時だけ適用する。
  try {
    const headersMap = getSheetHeaders_();
    const headers = headersMap[sheet.getName()];
    if (!headers) return;

    const startRow = maxRows + 1;
    const rowCount = AUTO_EXTEND_ROWS_BUFFER;
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sheetName = sheet.getName();

    const statusList = getStatusListForSheet_(sheetName);
    const staffList = [CLEAR_LABEL, ...getStaffMembers_()];
    const carList = [CLEAR_LABEL, "4t", "軽トラ②", "8tユニック", "軽ワゴン③", "ローダー", "P.BOX①", "P.BOX②", "3t", "キャブ③", "軽トラ③", "2t①", "軽ワゴン⑥", "軽バン⑤", "軽ダンプ②"];
    const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
    const importanceList = [CLEAR_LABEL, "高", "中", "低"];

    currentHeaders.forEach((header, i) => {
      const col = i + 1;
      if (isAutoOutputSheet_(sheetName)) return;

      if (DATE_HEADERS.includes(header)) {
        const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
        sheet.getRange(startRow, col, rowCount, 1).setNumberFormat("yyyy/mm/dd").setDataValidation(rule);
      }
      if (header === "状態") setDropdownForRange_(sheet, startRow, col, rowCount, statusList);
      if (header === "担当" || header === "入力者" || header === "利用者") setDropdownForRange_(sheet, startRow, col, rowCount, staffList);
      if (header === "社用車") setDropdownForRange_(sheet, startRow, col, rowCount, carList);
      if (header === "電話対応") setDropdownForRange_(sheet, startRow, col, rowCount, phoneList);
      if (header === "重要度") setDropdownForRange_(sheet, startRow, col, rowCount, importanceList);
      if (header === "コピー有無") setDropdownForRange_(sheet, startRow, col, rowCount, getCopyStatusList_());
      if (header === "返却済") setDropdownForRange_(sheet, startRow, col, rowCount, getReturnedStatusList_());
    });
  } catch (e) {
    console.log("追加行の入力規則設定をスキップ: " + e.message);
  }
}

function setDropdownForRange_(sheet, startRow, col, rowCount, list) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, col, rowCount, 1).setDataValidation(rule);
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("別の処理中のため、全体更新をスキップしました。少し待ってから再実行してください。");
    return;
  }

  try {
    // v9.6.4：軽量更新でも、集約前に主要シートを日付順へ整える。
    refreshInputSheetsLight_();
    sortAllOperationalSheetsSilent_();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    clearSummaryDirty_();

    SpreadsheetApp.getActiveSpreadsheet().toast("全体更新（軽量）が完了しました");
  } finally {
    lock.releaseLock();
  }
}

function refreshInputSheetsLight_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    try {
      updateNoticeForSheet_(sheet);
      ensureIdsForSheet_(sheet);
      setCheckboxesForDataRows(sheet);
    } catch (e) {
      console.log(name + " の軽量更新をスキップ: " + e.message);
    }
  });
}

function sortAllOperationalSheetsSilent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configs = getSortConfigs_();

  Object.keys(configs).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 3) return;

    try {
      sortSheetByConfig_(sheet, configs[sheetName]);
    } catch (e) {
      console.log(sheetName + " の並び替えをスキップ: " + e.message);
    }
  });
}

function sortAllOperationalSheets() {
  sortAllOperationalSheetsSilent_();
  SpreadsheetApp.getActiveSpreadsheet().toast("主要シートを日時順に並び替えました");
}

function getSortConfigs_() {
  return {
    "出先予定": {date: "日付", direction: "asc", status: "状態"},
    "工事予定": {date: "開始日", direction: "asc", status: "状態"},
    "会議予定": {date: "日付", direction: "asc", status: "状態"},
    "行事予定": {date: "開始日", direction: "asc", status: "状態"},
    "電話履歴": {date: "日時", direction: "desc", status: "電話対応"},
    "社用車予約": {date: "日付", direction: "asc", status: "状態"},
    "日報": {date: "日付", direction: "desc", status: "状態"},
    "日報レシート管理": {date: "日付", direction: "desc", status: "精算状態"},
    "免許資格管理": {date: "更新期限", direction: "asc", status: "状態"},
    "備品修理管理": {date: "返却予定日", fallbackDate: "修理依頼日", direction: "asc", status: "状態"},
    "お知らせ": {date: "投稿日", direction: "desc", status: "重要度"},
    "個人ToDo": {date: "期限", fallbackDate: "登録日", direction: "asc", status: "状態"},
    "車検管理": {date: "車検期限", direction: "asc", status: "状態", notice: "通知", secondaryDate: "保険期限", name: "車両名"},
    "車検履歴": {date: "更新日", direction: "desc", name: "車両名"},
    "一覧スケジュール": {date: "日付", direction: "asc", status: "状態", notice: "通知", type: "種類"},
    "過去一覧": {date: "日付", direction: "desc", type: "種類"},
    "帳簿PDF履歴": {date: "作成日時", direction: "desc", type: "対象"}
  };
}

function sortSheetByConfig_(sheet, config) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dateCol = headers.indexOf(config.date) + 1;
  const fallbackDateCol = config.fallbackDate ? headers.indexOf(config.fallbackDate) + 1 : 0;
  if (dateCol <= 0 && fallbackDateCol <= 0) return;

  const statusCol = config.status ? headers.indexOf(config.status) + 1 : 0;
  const noticeCol = config.notice ? headers.indexOf(config.notice) + 1 : 0;
  const secondaryDateCol = config.secondaryDate ? headers.indexOf(config.secondaryDate) + 1 : 0;
  const nameCol = config.name ? headers.indexOf(config.name) + 1 : 0;
  const typeCol = config.type ? headers.indexOf(config.type) + 1 : 0;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const noticePriority = {"重要": 0, "期限切れ": 1, "今日": 2, "3日以内": 3, "7日以内": 4, "予約重複": 5, "": 9};
  const statusPriority = {
    "要確認": 1, "未対応": 1, "期限切れ": 1, "失効": 1,
    "折返し": 2, "対応中": 2, "返却待ち": 2, "修理中": 2,
    "予定": 3, "予約済": 3, "着工前": 3, "施工前": 3, "未着手": 3, "更新予定": 3,
    "施工中": 4, "進行中": 4, "使用中": 4, "依頼済": 4,
    "提出済": 5, "実施済": 5,
    "完了": 8, "更新済": 8, "確認済": 8, "返却済": 8, "有効": 8,
    "中止": 9, "延期": 9,
    "": 7
  };

  values.sort((a, b) => {
    if (noticeCol > 0) {
      const pa = noticePriority[String(a[noticeCol - 1] || "")] || 9;
      const pb = noticePriority[String(b[noticeCol - 1] || "")] || 9;
      if (pa !== pb) return pa - pb;
    }

    if (statusCol > 0) {
      const sa = statusPriority[String(a[statusCol - 1] || "")] || 7;
      const sb = statusPriority[String(b[statusCol - 1] || "")] || 7;
      if (sa !== sb) return sa - sb;
    }

    const rawA = dateCol > 0 ? a[dateCol - 1] : (fallbackDateCol > 0 ? a[fallbackDateCol - 1] : "");
    const rawB = dateCol > 0 ? b[dateCol - 1] : (fallbackDateCol > 0 ? b[fallbackDateCol - 1] : "");
    const timeA = getSortableTime_(rawA || (fallbackDateCol > 0 ? a[fallbackDateCol - 1] : ""));
    const timeB = getSortableTime_(rawB || (fallbackDateCol > 0 ? b[fallbackDateCol - 1] : ""));
    if (timeA !== timeB) {
      return config.direction === "desc" ? timeB - timeA : timeA - timeB;
    }

    if (secondaryDateCol > 0) {
      const subA = getSortableTime_(a[secondaryDateCol - 1]);
      const subB = getSortableTime_(b[secondaryDateCol - 1]);
      if (subA !== subB) return subA - subB;
    }

    if (typeCol > 0) {
      const ta = String(a[typeCol - 1] || "");
      const tb = String(b[typeCol - 1] || "");
      if (ta !== tb) return ta.localeCompare(tb, "ja");
    }

    if (nameCol > 0) {
      return String(a[nameCol - 1] || "").localeCompare(String(b[nameCol - 1] || ""), "ja");
    }

    return 0;
  });

  sheet.getRange(2, 1, values.length, lastCol).setValues(values);
}

function applyColorRules(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = headers.indexOf("状態") + 1;
  const phoneCol = headers.indexOf("電話対応") + 1;
  const readCol = headers.indexOf(READ_HEADER) + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const typeCol = headers.indexOf("種類") + 1;
  const importanceCol = headers.indexOf("重要度") + 1;
  const returnedCol = headers.indexOf("返却済") + 1;
  const copyCol = headers.indexOf("コピー有無") + 1;

  const rules = [];
  const ruleRowCount = getApplyRowCount_(sheet);

  function addTextRule(col, text, bg) {
    if (col <= 0) return;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(bg)
        .setRanges([sheet.getRange(2, col, ruleRowCount)])
        .build()
    );
  }

  // 状態：意味ごとに色を分ける
  ["要確認", "未対応", "期限切れ", "失効", "差戻し", "未依頼"].forEach(v => addTextRule(statusCol, v, "#f4cccc"));
  ["対応中", "折返し", "更新予定", "依頼済"].forEach(v => addTextRule(statusCol, v, "#fff2cc"));
  ["返却待ち", "修理中"].forEach(v => addTextRule(statusCol, v, "#fce5cd"));
  ["着工前", "施工前", "未着手"].forEach(v => addTextRule(statusCol, v, "#d9d2e9"));
  ["施工中", "進行中", "使用中"].forEach(v => addTextRule(statusCol, v, "#cfe2f3"));
  ["予定", "予約済", "実施済", "下書き"].forEach(v => addTextRule(statusCol, v, "#eeeeee"));
  ["完了", "更新済", "確認済", "返却済", "有効", "提出済"].forEach(v => addTextRule(statusCol, v, "#d9ead3"));
  ["延期"].forEach(v => addTextRule(statusCol, v, "#ffe599"));
  ["中止"].forEach(v => addTextRule(statusCol, v, "#ea9999"));

  // 電話対応
  addTextRule(phoneCol, "未対応", "#f4cccc");
  addTextRule(phoneCol, "折返し", "#fff2cc");
  addTextRule(phoneCol, "対応中", "#fff2cc");
  addTextRule(phoneCol, "完了", "#d9ead3");

  // 返却済・コピー有無
  addTextRule(returnedCol, "未", "#fff2cc");
  addTextRule(returnedCol, "済", "#d9ead3");
  addTextRule(copyCol, "有", "#d9ead3");
  addTextRule(copyCol, "無", "#f4cccc");
  addTextRule(copyCol, "未確認", "#fff2cc");

  if (readCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied("=$" + columnLetter(readCol) + "2=FALSE")
        .setBackground("#fff2cc")
        .setRanges([sheet.getRange(2, readCol, ruleRowCount)])
        .build()
    );
  }

  // 通知：緊急度で色分け
  addTextRule(noticeCol, "重要", "#ea9999");
  addTextRule(noticeCol, "期限切れ", "#ea9999");
  addTextRule(noticeCol, "今日", "#fce5cd");
  addTextRule(noticeCol, "3日以内", "#fff2cc");
  addTextRule(noticeCol, "7日以内", "#d9ead3");
  addTextRule(noticeCol, "予約重複", "#ea9999");

  // 種類：状態や通知と被りすぎない淡色にする
  addTextRule(typeCol, "お知らせ", "#fff8e1");
  addTextRule(typeCol, "車検", "#fce4ec");
  addTextRule(typeCol, "備品修理", "#f3e5f5");
  addTextRule(typeCol, "免許資格", "#e8eaf6");
  addTextRule(typeCol, "会議", "#ede7f6");
  addTextRule(typeCol, "行事", "#fff3e0");
  addTextRule(typeCol, "工事予定", "#e3f2fd");
  addTextRule(typeCol, "出先予定", "#e8f5e9");
  addTextRule(typeCol, "電話履歴", "#fffde7");
  addTextRule(typeCol, "作業状況", "#f5f5f5");
  addTextRule(typeCol, "ToDo", "#e8f5e9");
  addTextRule(typeCol, "社用車予約", "#e0f2f1");
  addTextRule(typeCol, "日報", "#f1f8e9");

  addTextRule(importanceCol, "高", "#ea9999");
  addTextRule(importanceCol, "中", "#ffe599");
  addTextRule(importanceCol, "低", "#d9ead3");

  sheet.setConditionalFormatRules(rules);
}

function sortVehicleInspectionSheetSilent_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");
  if (!sheet || sheet.getLastRow() < 3) return;
  sortSheetByConfig_(sheet, getSortConfigs_()["車検管理"]);
}



/**
 * v9.6.5 追加修正（実運用メニュー最小化・ロック通知整理版）
 *
 * 目的：
 * - メニューを「初回セットアップ」「日常運用」「保守・テスト」に寄せ、手順迷子を減らす
 * - refreshAll / onEdit のロック通知を連発しない
 * - onEdit は入力行の最低限処理だけにする
 * - 全体更新はユーザーが押した時だけ集約処理を行う
 */

function showSkipToastOnce_(message, key, seconds) {
  try {
    const cache = CacheService.getDocumentCache();
    const cacheKey = key || "SKIP_TOAST_LOCK";
    const exists = cache.get(cacheKey);
    if (exists) return;
    cache.put(cacheKey, "1", seconds || 10);
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
  } catch (e) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(message);
    } catch (err) {}
  }
}

function showDoneToast_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
  } catch (e) {}
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("社内管理")
    .addSubMenu(
      ui.createMenu("1. 初回セットアップ")
        .addItem("初回セットアップ（軽量）", "initialSetupForProduction")
        .addItem("シート構成を作成", "createCompanySheets")
        .addItem("設定シートを作成/確認", "setupSettingsSheet")
        .addItem("設定シートの備考を整える", "normalizeSettingsSheetNotesMenu")
        .addItem("入力シートだけ設定反映", "applySettingsToInputSheetsOnly")
        .addItem("帳簿出力設定を作成/確認", "setupLedgerOutputSettingsSheet")
        .addItem("手順シートを作成", "createProcedureSheet")
        .addItem("社内管理説明シートを作成", "createCompanyManagementExplanationSheet")
    )
    .addSubMenu(
      ui.createMenu("2. 日常運用")
        .addItem("全体更新（軽量）", "refreshAll")
        .addItem("主要シートを日時順に並び替え", "sortAllOperationalSheets")
        .addItem("一覧帳簿PDFを作成", "createLedgerPdf")
        .addSeparator()
        .addItem("選択行の日報 他現場状況を作成", "fillDailyReportOtherSituationForSelectedRow")
        .addItem("選択行の日報文章をGPT作成", "createDailyReportText")
        .addItem("選択行の日報PDFを作成", "createSelectedDailyReportPdf")
        .addItem("車検更新完了処理", "processCompletedVehicleInspections")
    )
    .addSubMenu(
      ui.createMenu("3. 保守・テスト")
        .addItem("個人確認グループを作り直す", "rebuildCheckGroups")
        .addItem("日報PDFリンク列を修復", "fixDailyReportPdfLinkColumn")
        .addItem("表示幅を調整", "formatInputSheetsForLongText")
        .addItem("裏方シートを非表示", "hideSupportSheets")
        .addItem("スプレッドシートをバックアップ", "backupSpreadsheet")
        .addSeparator()
        .addItem("テスト用サンプルデータを追加", "addSampleDataForTest")
        .addItem("過去移動ルールをチェック", "checkArchiveRulesForTest")
        .addItem("移行前チェックを実行", "runPreMigrationCheck")
        .addSeparator()
        .addItem("SQL関連シートを作成", "createSqlSupportSheets")
        .addItem("全体更新（重い）", "refreshAllHeavy")
        .addItem("全シートを月ごとに折りたたみ", "groupAllRowsByMonth")
        .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
    )
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(500)) {
    // onEdit は頻発するため、スキップ通知を出さない。
    // 必要な場合は手動で「全体更新（軽量）」を実行する。
    return;
  }

  try {
    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const row = e.range.getRow();
    const col = e.range.getColumn();

    if (row === 1) return;

    ensureRowsAfterEditLight_(sheet, row);

    if (e.range.getValue() === CLEAR_LABEL) {
      e.range.clearContent();
      markSummaryDirty_();
      return;
    }

    ensureIdForEditedRow_(sheet, row);
    updateNoticeForEditedRow(sheet, row, col);
    updateCheckboxesForEditedRow(sheet, row);

    if (sheetName === "社用車予約") {
      const conflictMessage = getCarReservationConflictMessage(sheet, row);
      if (conflictMessage) {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const noticeCol = headers.indexOf("通知") + 1;
        if (noticeCol > 0) sheet.getRange(row, noticeCol).setValue("予約重複");
        showSkipToastOnce_(conflictMessage, "CAR_RESERVATION_CONFLICT", 8);
      }
    }

    if (getInputSheetNames_().includes(sheetName)) {
      markSummaryDirty_();
    }
  } finally {
    lock.releaseLock();
  }
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    showSkipToastOnce_(
      '別の処理中のため、全体更新をスキップしました。少し待ってから再実行してください。',
      'REFRESH_ALL_SKIP_TOAST',
      10
    );
    return;
  }

  try {
    refreshInputSheetsLight_();
    sortAllOperationalSheetsSilent_();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();

    // 一覧スケジュールは再作成されるため、更新後に必ず表示幅・縦書き・シート順を戻す。
    try {
      formatInputSheetsForLongText();
    } catch (e) {
      console.log('全体更新後の表示幅調整をスキップ: ' + e.message);
    }

    try {
      reorderSheetsForOperation();
    } catch (e) {
      console.log('全体更新後のシート順調整をスキップ: ' + e.message);
    }

    clearSummaryDirty_();
    showDoneToast_('全体更新（軽量）が完了しました');
  } finally {
    lock.releaseLock();
  }
}

function initialSetupForProduction() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    showSkipToastOnce_(
      "別の処理中のため、初回セットアップをスキップしました。少し待ってから再実行してください。",
      "INITIAL_SETUP_SKIP_TOAST",
      10
    );
    return;
  }

  try {
    // v9.6.5：初回セットアップは軽い処理だけに絞る。
    // 集計更新・表示幅調整・SQL作成・移行前チェックは必要時に個別実行。
    createCompanySheets();
    setupSettingsSheet();
    setupLedgerOutputSettingsSheet();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    applySettingsToInputSheetsOnly();
    fixDailyReportPdfLinkColumn();
    rebuildCheckGroups();
    showDoneToast_("初回セットアップ（軽量）が完了しました。次に全体更新（軽量）を実行してください。");
  } finally {
    lock.releaseLock();
  }
}

function createProcedureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("手順シート");
  if (!sheet) sheet = ss.insertSheet("手順シート");

  sheet.clear();
  sheet.clearFormats();

  const headers = ["区分", "手順", "実行メニュー", "タイミング", "注意"];
  const rows = [
    ["初回", "1", "社内管理 → 1. 初回セットアップ → 初回セットアップ（軽量）", "最初だけ", "サンプルデータは入りません"],
    ["初回", "2", "社内管理 → 2. 日常運用 → 全体更新（軽量）", "初回セットアップ後", "一覧・要確認・集計を作成します"],
    ["日常", "1", "入力シートへ入力", "随時", "入力直後は一覧に即時反映しない設計です"],
    ["日常", "2", "社内管理 → 2. 日常運用 → 全体更新（軽量）", "入力後・確認前", "まとめて反映することで重くなりにくくしています"],
    ["日常", "3", "社内管理 → 2. 日常運用 → 一覧帳簿PDFを作成", "帳簿出力時", "帳簿出力設定でTRUEのものだけ出力します"],
    ["保守", "1", "社内管理 → 3. 保守・テスト → 個人確認グループを作り直す", "担当者や確認者を変更した時", "既読列は表示、個人確認列だけ折りたたみます"],
    ["保守", "2", "社内管理 → 3. 保守・テスト → 日報PDFリンク列を修復", "PDFリンク列にTRUE/FALSEが出た時", "PDFリンク列をURL用に戻します"],
    ["テスト", "1", "社内管理 → 3. 保守・テスト → テスト用サンプルデータを追加", "テスト環境だけ", "本番では押さないでください"],
    ["テスト", "2", "社内管理 → 3. 保守・テスト → 過去移動ルールをチェック", "移行前確認", "実データは移動せず判定だけ確認します"],
    ["月次", "1", "社内管理 → 3. 保守・テスト → 過去予定を過去一覧へ移動", "月次・確認後", "本番移行直後は押さないでください"],
    ["社外共有前確認", "1", "匿名化確認", "社外共有・資料化前", "実名・社名・地名・車番・取引先名が残っていないか確認"],
    ["AppSheet", "1", "Regenerate Structure", "列変更後", "ID列をKey、PDFリンクをURL、既読列をYes/Noに確認"]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 420);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 420);
  sheet.getRange(1, 1, rows.length + 1, headers.length).setWrap(true).setVerticalAlignment("middle");
  createFilterSafely_(sheet, headers.length);
}



/**
 * v9.6.6 追加修正（日時順優先・色分け再整理・空欄に戻す非表示）
 *
 * 目的：
 * - 各シートの並び順を「日付/日時優先」に統一する
 * - 要確認一覧も日付順を優先し、同日内で通知優先度順にする
 * - 返却待ち/修理中/未/済、電話対応などの色かぶりを減らす
 * - プルダウンに「空欄に戻す」を出さない。空欄にしたい時はセルをDeleteで消す
 * - 既に入ってしまった「空欄に戻す」は保守メニューから一括クリアできる
 */

function cleanDropdownList_(list) {
  const seen = {};
  return (list || [])
    .map(v => String(v || '').trim())
    .filter(v => v && v !== CLEAR_LABEL)
    .filter(v => {
      if (seen[v]) return false;
      seen[v] = true;
      return true;
    });
}

function setDropdown(sheet, col, list) {
  const rowCount = getApplyRowCount_(sheet);
  const cleaned = cleanDropdownList_(list);
  if (!cleaned.length) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cleaned, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, rowCount).setDataValidation(rule);
}

function setDropdownForRange_(sheet, startRow, col, rowCount, list) {
  const cleaned = cleanDropdownList_(list);
  if (!cleaned.length || rowCount <= 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cleaned, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(startRow, col, rowCount, 1).setDataValidation(rule);
}

function clearClearLabelValuesInAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sheet => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;

    try {
      const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
      const values = range.getValues();
      let changed = false;

      const next = values.map(row => row.map(v => {
        if (String(v || '').trim() === CLEAR_LABEL) {
          changed = true;
          return '';
        }
        return v;
      }));

      if (changed) range.setValues(next);
    } catch (e) {
      console.log(sheet.getName() + ' の空欄戻し表示クリアをスキップ: ' + e.message);
    }
  });

  showDoneToast_('「空欄に戻す」の表示をクリアしました');
}

function getDateSortValueForRow_(row, primaryCol, fallbackCol, direction) {
  let raw = '';
  if (primaryCol > 0) raw = row[primaryCol - 1];
  if (!raw && fallbackCol > 0) raw = row[fallbackCol - 1];

  const t = getSortableTimeSafe_(raw);
  if (t !== null) return t;

  return direction === 'desc' ? -8640000000000000 : 8640000000000000;
}

function getSortableTimeSafe_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.getTime();
  }
  const d = new Date(value);
  if (!isNaN(d.getTime())) return d.getTime();
  return null;
}

function getNoticeSortPriority_(notice) {
  const map = {
    '重要': 0,
    '期限切れ': 1,
    '今日': 2,
    '3日以内': 3,
    '7日以内': 4,
    '予約重複': 1
  };
  return map[String(notice || '').trim()] || 99;
}

function getStatusSortPriority_(status) {
  const map = {
    '未対応': 1,
    '要確認': 1,
    '期限切れ': 1,
    '失効': 1,
    '返却待ち': 2,
    '折返し': 2,
    '対応中': 3,
    '修理中': 3,
    '未着手': 3,
    '予定': 4,
    '予約済': 4,
    '着工前': 4,
    '施工前': 4,
    '施工中': 5,
    '進行中': 5,
    '使用中': 5,
    '下書き': 5,
    '提出済': 6,
    '実施済': 6,
    '完了': 8,
    '更新済': 8,
    '返却済': 8,
    '確認済': 8,
    '有効': 8,
    '中止': 9,
    '延期': 7
  };
  return map[String(status || '').trim()] || 50;
}

function sortSheetByConfig_(sheet, config) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 3 || lastCol < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dateCol = config.date ? headers.indexOf(config.date) + 1 : 0;
  const fallbackDateCol = config.fallbackDate ? headers.indexOf(config.fallbackDate) + 1 : 0;
  if (dateCol <= 0 && fallbackDateCol <= 0) return;

  const statusCol = config.status ? headers.indexOf(config.status) + 1 : 0;
  const noticeCol = config.notice ? headers.indexOf(config.notice) + 1 : 0;
  const secondaryDateCol = config.secondaryDate ? headers.indexOf(config.secondaryDate) + 1 : 0;
  const nameCol = config.name ? headers.indexOf(config.name) + 1 : 0;
  const typeCol = config.type ? headers.indexOf(config.type) + 1 : 0;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const nonEmpty = values.filter(row => row.some(v => v !== '' && v !== null));
  const empty = values.filter(row => !row.some(v => v !== '' && v !== null));

  nonEmpty.sort((a, b) => {
    // v9.6.6：最優先は日付/日時。シートごとの行が日付順不同に見えないようにする。
    const timeA = getDateSortValueForRow_(a, dateCol, fallbackDateCol, config.direction);
    const timeB = getDateSortValueForRow_(b, dateCol, fallbackDateCol, config.direction);
    if (timeA !== timeB) {
      return config.direction === 'desc' ? timeB - timeA : timeA - timeB;
    }

    // 同じ日付内だけ、通知・状態・種類・名称で整える。
    if (noticeCol > 0) {
      const pa = getNoticeSortPriority_(a[noticeCol - 1]);
      const pb = getNoticeSortPriority_(b[noticeCol - 1]);
      if (pa !== pb) return pa - pb;
    }

    if (statusCol > 0) {
      const sa = getStatusSortPriority_(a[statusCol - 1]);
      const sb = getStatusSortPriority_(b[statusCol - 1]);
      if (sa !== sb) return sa - sb;
    }

    if (secondaryDateCol > 0) {
      const subA = getSortableTimeSafe_(a[secondaryDateCol - 1]) || 8640000000000000;
      const subB = getSortableTimeSafe_(b[secondaryDateCol - 1]) || 8640000000000000;
      if (subA !== subB) return subA - subB;
    }

    if (typeCol > 0) {
      const ta = String(a[typeCol - 1] || '');
      const tb = String(b[typeCol - 1] || '');
      if (ta !== tb) return ta.localeCompare(tb, 'ja');
    }

    if (nameCol > 0) {
      return String(a[nameCol - 1] || '').localeCompare(String(b[nameCol - 1] || ''), 'ja');
    }

    return 0;
  });

  sheet.getRange(2, 1, values.length, lastCol).setValues(nonEmpty.concat(empty));
}

function createAlertList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName('一覧スケジュール');

  let sheet = ss.getSheetByName('要確認一覧');
  if (!sheet) sheet = ss.insertSheet('要確認一覧');

  const headers = getSheetHeaders_()['要確認一覧'];

  resetSheet(sheet, headers.length);
  setupSheet(sheet, headers);

  const rows = [];
  const priority = {
    '重要': 0,
    '期限切れ': 1,
    '今日': 2,
    '3日以内': 3,
    '7日以内': 4,
    '予約重複': 1
  };

  if (source && source.getLastRow() >= 2) {
    const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
    const values = source.getRange(2, 1, source.getLastRow() - 1, sourceHeaders.length).getValues();
    values
      .map(valuesRow => objectFromRow(sourceHeaders, valuesRow))
      .filter(row => shouldIncludeInAlertList_(row, priority))
      .forEach(row => rows.push(row));
  }

  const noticeSheet = ss.getSheetByName('お知らせ');
  if (noticeSheet && noticeSheet.getLastRow() >= 2) {
    const noticeHeaders = noticeSheet.getRange(1, 1, 1, noticeSheet.getLastColumn()).getValues()[0];
    const noticeValues = noticeSheet.getRange(2, 1, noticeSheet.getLastRow() - 1, noticeHeaders.length).getValues();

    noticeValues.forEach(valuesRow => {
      const row = objectFromRow(noticeHeaders, valuesRow);
      if (row['重要度'] === '高' && !isNoticeReadByAll_(row)) {
        rows.push({
          '日付': row['投稿日'] || new Date(),
          '種類': 'お知らせ',
          '内容': joinText(row['タイトル'], row['内容']),
          '担当': row['投稿者'],
          '状態': '要確認',
          '通知': '重要',
          '電話対応': '',
          '備考': row['備考'] || '',
          '元シート': 'お知らせ'
        });
      }
    });
  }

  rows.sort((a, b) => {
    // v9.6.6：要確認一覧は日付順を最優先にする。
    const da = getSortableTimeSafe_(a['日付']) || 8640000000000000;
    const db = getSortableTimeSafe_(b['日付']) || 8640000000000000;
    if (da !== db) return da - db;

    const pa = priority[String(a['通知'] || '').trim()] || 99;
    const pb = priority[String(b['通知'] || '').trim()] || 99;
    if (pa !== pb) return pa - pb;

    const sa = getStatusSortPriority_(a['状態']);
    const sb = getStatusSortPriority_(b['状態']);
    if (sa !== sb) return sa - sb;

    return String(a['種類'] || '').localeCompare(String(b['種類'] || ''), 'ja');
  });

  writeObjectsToSheet(sheet, rows.map(row => ({
    '日付': row['日付'],
    '種類': row['種類'],
    '内容': row['内容'],
    '担当': row['担当'],
    '状態': row['状態'],
    '通知': row['通知'],
    '電話対応': row['電話対応'],
    '備考': row['備考'] || '',
    '元シート': row['元シート']
  })));

  applyColorRules(sheet);
}

function applyColorRules(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const statusCol = headers.indexOf('状態') + 1;
  const phoneCol = headers.indexOf('電話対応') + 1;
  const readCol = headers.indexOf(READ_HEADER) + 1;
  const noticeCol = headers.indexOf('通知') + 1;
  const typeCol = headers.indexOf('種類') + 1;
  const importanceCol = headers.indexOf('重要度') + 1;
  const returnedCol = headers.indexOf('返却済') + 1;
  const copyCol = headers.indexOf('コピー有無') + 1;

  const rules = [];
  const ruleRowCount = Math.max(getApplyRowCount_(sheet), 1);

  function addTextRule(col, text, bg) {
    if (col <= 0) return;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(text)
        .setBackground(bg)
        .setRanges([sheet.getRange(2, col, ruleRowCount)])
        .build()
    );
  }

  // 状態列：似た色を減らし、意味で分ける
  ['要確認', '未対応', '期限切れ', '失効', '差戻し', '未依頼'].forEach(v => addTextRule(statusCol, v, '#f4cccc'));
  ['返却待ち'].forEach(v => addTextRule(statusCol, v, '#fce5cd'));
  ['修理中', '対応中', '折返し', '更新予定', '依頼済'].forEach(v => addTextRule(statusCol, v, '#fff2cc'));
  ['着工前', '施工前', '未着手'].forEach(v => addTextRule(statusCol, v, '#d9d2e9'));
  ['施工中', '進行中', '使用中'].forEach(v => addTextRule(statusCol, v, '#cfe2f3'));
  ['予定', '予約済', '実施済', '下書き'].forEach(v => addTextRule(statusCol, v, '#eeeeee'));
  ['提出済', '確認済', '完了', '更新済', '返却済', '有効'].forEach(v => addTextRule(statusCol, v, '#d9ead3'));
  ['延期'].forEach(v => addTextRule(statusCol, v, '#ffe599'));
  ['中止'].forEach(v => addTextRule(statusCol, v, '#ea9999'));

  // 電話対応：未対応/折返し/対応中/完了を区別
  addTextRule(phoneCol, '未対応', '#f4cccc');
  addTextRule(phoneCol, '折返し', '#fce5cd');
  addTextRule(phoneCol, '対応中', '#fff2cc');
  addTextRule(phoneCol, '完了', '#d9ead3');

  // 返却済：状態列とは別に、未/済をはっきり分ける
  addTextRule(returnedCol, '未', '#fff2cc');
  addTextRule(returnedCol, '済', '#d9ead3');

  addTextRule(copyCol, '有', '#d9ead3');
  addTextRule(copyCol, '無', '#f4cccc');
  addTextRule(copyCol, '未確認', '#fff2cc');

  if (readCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$' + columnLetter(readCol) + '2=FALSE')
        .setBackground('#fff2cc')
        .setRanges([sheet.getRange(2, readCol, ruleRowCount)])
        .build()
    );
  }

  // 通知列：緊急度を最も目立たせる
  addTextRule(noticeCol, '重要', '#e06666');
  addTextRule(noticeCol, '期限切れ', '#ea9999');
  addTextRule(noticeCol, '今日', '#f9cb9c');
  addTextRule(noticeCol, '3日以内', '#fff2cc');
  addTextRule(noticeCol, '7日以内', '#d9ead3');
  addTextRule(noticeCol, '予約重複', '#e06666');

  // 種類列：かなり淡くする。通知/状態の色を主役にする。
  addTextRule(typeCol, 'お知らせ', '#fffaf0');
  addTextRule(typeCol, '車検', '#fdecef');
  addTextRule(typeCol, '備品修理', '#f7eef8');
  addTextRule(typeCol, '免許資格', '#f0f2fb');
  addTextRule(typeCol, '会議', '#f4effb');
  addTextRule(typeCol, '行事', '#fff7ed');
  addTextRule(typeCol, '工事予定', '#edf7ff');
  addTextRule(typeCol, '出先予定', '#eef8ee');
  addTextRule(typeCol, '電話履歴', '#fffbea');
  addTextRule(typeCol, '作業状況', '#f7f7f7');
  addTextRule(typeCol, 'ToDo', '#eef8ee');
  addTextRule(typeCol, '社用車予約', '#eef9f7');
  addTextRule(typeCol, '日報', '#f4faef');

  addTextRule(importanceCol, '高', '#ea9999');
  addTextRule(importanceCol, '中', '#ffe599');
  addTextRule(importanceCol, '低', '#d9ead3');

  sheet.setConditionalFormatRules(rules);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('社内管理')
    .addSubMenu(
      ui.createMenu('1. 初回セットアップ')
        .addItem('初回セットアップ（軽量）', 'initialSetupForProduction')
        .addItem('シート構成を作成', 'createCompanySheets')
        .addItem('設定シートを作成/確認', 'setupSettingsSheet')
        .addItem('設定シートの備考を整える', 'normalizeSettingsSheetNotesMenu')
        .addItem('入力シートだけ設定反映', 'applySettingsToInputSheetsOnly')
        .addItem('帳簿出力設定を作成/確認', 'setupLedgerOutputSettingsSheet')
        .addItem('手順シートを作成', 'createProcedureSheet')
    )
    .addSubMenu(
      ui.createMenu('2. 日常運用')
        .addItem('全体更新（軽量）', 'refreshAll')
        .addItem('主要シートを日時順に並び替え', 'sortAllOperationalSheets')
        .addItem('一覧帳簿PDFを作成', 'createLedgerPdf')
        .addSeparator()
        .addItem('選択行の日報 他現場状況を作成', 'fillDailyReportOtherSituationForSelectedRow')
        .addItem('選択行の日報文章をGPT作成', 'createDailyReportText')
        .addItem('選択行の日報PDFを作成', 'createSelectedDailyReportPdf')
        .addItem('車検更新完了処理', 'processCompletedVehicleInspections')
    )
    .addSubMenu(
      ui.createMenu('3. 保守・テスト')
        .addItem('個人確認グループを作り直す', 'rebuildCheckGroups')
        .addItem('日報PDFリンク列を修復', 'fixDailyReportPdfLinkColumn')
        .addItem('空欄に戻す表示をクリア', 'clearClearLabelValuesInAllSheets')
        .addItem('表示幅を調整', 'formatInputSheetsForLongText')
        .addItem('裏方シートを非表示', 'hideSupportSheets')
        .addItem('スプレッドシートをバックアップ', 'backupSpreadsheet')
        .addSeparator()
        .addItem('テスト用サンプルデータを追加', 'addSampleDataForTest')
        .addItem('過去移動ルールをチェック', 'checkArchiveRulesForTest')
        .addItem('移行前チェックを実行', 'runPreMigrationCheck')
        .addSeparator()
        .addItem('SQL関連シートを作成', 'createSqlSupportSheets')
        .addItem('全体更新（重い）', 'refreshAllHeavy')
        .addItem('現在のシートを月ごとに折りたたみ', 'groupRowsByMonth')
        .addItem('全シートを月ごとに折りたたみ', 'groupAllRowsByMonth')
        .addItem('過去予定を過去一覧へ移動', 'movePastItemsToArchive')
    )
    .addToUi();
}


/**
 * v9.6.9 追加修正（設定シート備考整理・個人確認列縦書き維持・空欄に戻す選択維持）
 *
 * 目的:
 * - 社用時に担当者/既読確認者をフルネームにしても横スクロールを増やしすぎないよう、
 *   個人確認列だけ縦書き・細幅に戻す。
 * - 「空欄に戻す」をプルダウンに再表示し、選択したら onEdit で即クリアする。
 * - 既読列、PDFリンク列、ID列は縦書き対象外にする。
 *
 * 補足:
 * - Google Sheets のプルダウンに「完全な空欄項目」を安定して入れて Enter 確定するのは扱いづらいため、
 *   表示上は「空欄に戻す」を選び、onEdit で空欄へ戻す方式にする。
 */

function cleanDropdownList_(list) {
  const seen = {};
  return (list || [])
    .map(v => String(v || '').trim())
    .filter(v => v)
    .filter(v => {
      if (seen[v]) return false;
      seen[v] = true;
      return true;
    });
}

function setDropdown(sheet, col, list) {
  const rowCount = getApplyRowCount_(sheet);
  const cleaned = cleanDropdownList_(list);
  if (!cleaned.length) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cleaned, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(2, col, rowCount).setDataValidation(rule);
}

function setDropdownForRange_(sheet, startRow, col, rowCount, list) {
  const cleaned = cleanDropdownList_(list);
  if (!cleaned.length || rowCount <= 0) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(cleaned, true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange(startRow, col, rowCount, 1).setDataValidation(rule);
}

function formatInputSheetsForLongText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const targets = [
    'ダッシュボード',
    '一覧スケジュール',
    '要確認一覧',
    '出先予定',
    '工事予定',
    '会議予定',
    '行事予定',
    '作業状況',
    '電話履歴',
    '車検管理',
    '車検履歴',
    '社用車予約',
    '日報',
    '日報レシート管理',
    '免許資格管理',
    '備品修理管理',
    'お知らせ',
    '個人ToDo',
    '担当別未読',
    '既読率集計',
    '過去一覧',
    '帳簿PDF履歴',
    '帳簿出力設定',
    '設定',
    '手順シート',
    '社内管理説明',
    'SQL設計',
    '移行対応表',
    'SQLサンプル集'
  ];

  targets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    formatSheetBase_(sheet);
    formatSheetByName_(sheet);

    if (name === '設定') {
      formatSettingsSheet_(sheet);
    }

    if (name === '既読率集計') {
      formatReadRateSummarySheet_(sheet);
    }

    // 一覧スケジュールも含め、個人確認列だけ縦書きにする。
    formatPersonalCheckColumnsVertical_(sheet);
    hideTechnicalColumns_(sheet);
  });

  SpreadsheetApp.getActiveSpreadsheet().toast('表示幅を調整しました（一覧スケジュールの個人確認列も縦書き）');
}

function formatDailyReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('日報');
  if (!sheet) return;

  formatSheetBase_(sheet);
  formatSheetByName_(sheet);
  formatPersonalCheckColumnsVertical_(sheet);
  hideTechnicalColumns_(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast('日報シートの表示幅を調整しました（個人確認列は縦書き）');
}

function formatPersonalCheckColumnsVerticalAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getPersonalCheckVerticalTargetSheets_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    formatPersonalCheckColumnsVertical_(sheet);
    hideTechnicalColumns_(sheet);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('個人確認列を縦書きにしました');
}

function formatPersonalCheckColumnsVertical_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const personalMembers = getPersonalMembers_();
  if (!personalMembers || personalMembers.length === 0) return;

  try {
    sheet.setRowHeight(1, 120);
  } catch (e) {}

  // 既読列は横書きで残す。
  const readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol > 0) {
    try {
      sheet.setColumnWidth(readCol, 56);
      sheet.getRange(1, readCol, sheet.getMaxRows(), 1)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);
      sheet.getRange(1, readCol).setVerticalText(false).setTextRotation(0);
    } catch (e) {}
  }

  // 個人確認列は、ヘッダーだけでなく列全体を中央寄せにする。
  personalMembers.forEach(member => {
    const col = headers.indexOf(member) + 1;
    if (col <= 0) return;

    try {
      sheet.setColumnWidth(col, 36);

      const headerCell = sheet.getRange(1, col);
      headerCell
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);

      try {
        headerCell.setVerticalText(true);
      } catch (e1) {
        try {
          headerCell.setTextRotation(90);
        } catch (e2) {}
      }

      if (sheet.getMaxRows() > 1) {
        sheet.getRange(2, col, sheet.getMaxRows() - 1, 1)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle')
          .setWrap(false);
      }
    } catch (e) {
      console.log(sheet.getName() + ' / ' + member + ' の縦書き設定をスキップ: ' + e.message);
    }
  });
}

function hideTechnicalColumns_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hideNames = [CALENDAR_ID_HEADER];

  const idHeader = getIdHeaderForSheet_(sheet.getName());
  if (idHeader) hideNames.push(idHeader);

  hideNames.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;
    try {
      sheet.hideColumns(col);
    } catch (e) {}
  });
}

function rebuildCheckGroups() {
  resetColumnGroups();
  createCheckGroup();
  formatPersonalCheckColumnsVerticalAllSheets();
}

function initialSetupForProduction() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    showSkipToastOnce_(
      '別の処理中のため、初回セットアップをスキップしました。少し待ってから再実行してください。',
      'INITIAL_SETUP_SKIP_TOAST',
      10
    );
    return;
  }

  try {
    createCompanySheets();
    setupSettingsSheet();
    setupLedgerOutputSettingsSheet();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    applySettingsToInputSheetsOnly();
    fixDailyReportPdfLinkColumn();
    rebuildCheckGroups();

    try {
      formatInputSheetsForLongText();
    } catch (e) {
      console.log('初回セットアップ時の表示幅調整をスキップ: ' + e.message);
    }

    try {
      reorderSheetsForOperation();
    } catch (e) {
      console.log('初回セットアップ時のシート順調整をスキップ: ' + e.message);
    }

    showDoneToast_('初回セットアップ（軽量）が完了しました。次に全体更新（軽量）を実行してください。');
  } finally {
    lock.releaseLock();
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('社内管理')
    .addSubMenu(
      ui.createMenu('1. 初回セットアップ')
        .addItem('初回セットアップ（軽量）', 'initialSetupForProduction')
        .addItem('シート構成を作成', 'createCompanySheets')
        .addItem('設定シートを作成/確認', 'setupSettingsSheet')
        .addItem('設定シートの備考を整える', 'normalizeSettingsSheetNotesMenu')
        .addItem('入力シートだけ設定反映', 'applySettingsToInputSheetsOnly')
        .addItem('帳簿出力設定を作成/確認', 'setupLedgerOutputSettingsSheet')
        .addItem('手順シートを作成', 'createProcedureSheet')
        .addItem('社内管理説明シートを作成', 'createCompanyManagementExplanationSheet')
    )
    .addSubMenu(
      ui.createMenu('2. 日常運用')
        .addItem('全体更新（軽量）', 'refreshAll')
        .addItem('主要シートを日時順に並び替え', 'sortAllOperationalSheets')
        .addItem('一覧帳簿PDFを作成', 'createLedgerPdf')
        .addSeparator()
        .addItem('選択行の日報 他現場状況を作成', 'fillDailyReportOtherSituationForSelectedRow')
        .addItem('選択行の日報文章をGPT作成', 'createDailyReportText')
        .addItem('選択行の日報PDFを作成', 'createSelectedDailyReportPdf')
        .addItem('車検更新完了処理', 'processCompletedVehicleInspections')
    )
    .addSubMenu(
      ui.createMenu('3. 保守・テスト')
        .addItem('個人確認グループを作り直す', 'rebuildCheckGroups')
        .addItem('個人確認列を縦書きにする', 'formatPersonalCheckColumnsVerticalAllSheets')
        .addItem('日報PDFリンク列を修復', 'fixDailyReportPdfLinkColumn')
        .addItem('空欄に戻す表示をクリア', 'clearClearLabelValuesInAllSheets')
        .addItem('表示幅を調整', 'formatInputSheetsForLongText')
        .addItem('シート順を整える', 'reorderSheetsForOperation')
        .addItem('裏方シートを非表示', 'hideSupportSheets')
        .addItem('スプレッドシートをバックアップ', 'backupSpreadsheet')
        .addSeparator()
        .addItem('テスト用サンプルデータを追加', 'addSampleDataForTest')
        .addItem('過去移動ルールをチェック', 'checkArchiveRulesForTest')
        .addItem('移行前チェックを実行', 'runPreMigrationCheck')
        .addSeparator()
        .addItem('SQL関連シートを作成', 'createSqlSupportSheets')
        .addItem('全体更新（重い）', 'refreshAllHeavy')
        .addItem('現在のシートを月ごとに折りたたみ', 'groupRowsByMonth')
        .addItem('全シートを月ごとに折りたたみ', 'groupAllRowsByMonth')
        .addItem('過去予定を過去一覧へ移動', 'movePastItemsToArchive')
    )
    .addToUi();
}






























/**
 * 社内管理説明シートを作成する。
 *
 * 手順シートが「どのメニューをいつ押すか」を説明するのに対して、
 * 社内管理説明は「このシステムが何をするものか」「どのシートをどう使うか」
 * 「運用ルール・トラブル対応・社外共有時の注意」を説明するための概要シートです。
 */
function createCompanyManagementExplanationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "社内管理説明";
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {}

  try {
    sheet.showColumns(1, sheet.getMaxColumns());
  } catch (e) {}

  try {
    sheet.showRows(1, sheet.getMaxRows());
  } catch (e) {}

  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.setConditionalFormatRules([]);

  const rows = [
    ["区分", "項目", "説明", "利用者向けメモ"],

    ["概要", "システム名", "社内共有業務管理システム", "予定・電話・車検・日報・備品修理・お知らせなどを一元管理するための社内向け管理システムです。"],
    ["概要", "目的", "紙・口頭・個別ファイルで分かれていた業務情報を、Google Sheets / Apps Script / AppSheetで共有しやすくすることです。", "予定漏れ、電話折返し忘れ、車検期限見落とし、日報共有漏れなどを減らすことを目的にしています。"],
    ["概要", "利用環境", "Google Sheets、Google Apps Script、AppSheetを利用します。", "PCではスプレッドシート、スマホ・タブレットではAppSheetから入力できます。"],
    ["概要", "対象業務", "出先予定、工事予定、会議予定、行事予定、電話履歴、車検管理、社用車予約、日報、免許資格管理、備品修理管理、お知らせ、個人ToDoなどです。", "社内の予定確認・対応状況確認・帳簿PDF出力に使います。"],

    ["主な機能", "一覧スケジュール", "各入力シートの情報を1つの一覧へ集約します。", "普段はここを見ると、全体の予定や対応状況を確認できます。"],
    ["主な機能", "要確認一覧", "期限切れ、今日、3日以内、7日以内など、対応が必要なものを抽出します。", "未対応電話、返却待ち備品、車検期限、免許資格期限などを確認します。"],
    ["主な機能", "既読・個人確認", "全体の既読列と、個人ごとの確認列を分けて管理します。", "誰が確認したかを簡単に見られます。"],
    ["主な機能", "担当別未読", "担当者ごとの未読件数や未対応件数を集計します。", "確認漏れがある担当者を把握できます。"],
    ["主な機能", "既読率集計", "各予定・お知らせなどの確認状況を集計します。", "共有情報がどの程度確認されているかを見るためのシートです。"],
    ["主な機能", "日報", "日々の作業内容、進捗、問題点、明日の予定、写真、PDFリンクを管理します。", "必要に応じて日報文章作成やPDF出力もできます。"],
    ["主な機能", "日報レシート管理", "紙レシート、支払先、金額、経理確認、精算状態を管理します。", "通常の日報とは分け、未確認・差戻し・未精算だけ要確認一覧へ出します。"],
    ["主な機能", "帳簿PDF", "一覧スケジュールや要確認一覧などをPDF帳簿として出力します。", "他部門・経理・管理者への共有用として使えます。"],
    ["主な機能", "過去一覧", "完了済み・処理済みの情報を過去一覧へ移動できます。", "未対応電話や未返却備品などは安全のため自動移動しない設計です。"],

    ["入力シート", "出先予定", "外出予定、行き先、用件、担当、社用車、状態などを管理します。", "外出や訪問予定を入力します。"],
    ["入力シート", "工事予定", "工事名、現場、開始日、終了日、状態、担当などを管理します。", "工事の予定・進捗確認に使います。"],
    ["入力シート", "会議予定", "会議日、会議名、内容、担当、資料などを管理します。", "会議予定の共有に使います。"],
    ["入力シート", "行事予定", "社内行事や期間のある予定を管理します。", "開始日・終了日で管理できます。"],
    ["入力シート", "作業状況", "現場ごとの作業内容、状態、担当、写真、備考などを管理します。", "進行中作業や要確認作業の共有に使います。"],
    ["入力シート", "電話履歴", "電話相手、内容、電話対応、担当、対応メモなどを管理します。", "未対応・対応中・折返し・完了で管理します。"],
    ["入力シート", "車検管理", "車両名、車番、車検期限、保険期限、状態、写真などを管理します。", "車検・保険期限の見落とし防止に使います。"],
    ["入力シート", "社用車予約", "社用車、利用者、行き先、用途、利用時間などを管理します。", "社用車予約の重複確認にも使います。"],
    ["入力シート", "日報レシート管理", "紙レシートの支払先、内容、区分、金額、支払方法、経理確認、精算状態を管理します。", "紙の日報レシートを後から探せるようにするための管理表です。"],
    ["入力シート", "免許資格管理", "担当者ごとの資格名、取得日、更新期限、コピー有無、状態を管理します。", "資格・免許更新の期限管理に使います。"],
    ["入力シート", "備品修理管理", "備品名、場所、修理内容、返却予定日、返却済、状態を管理します。", "返却待ち・修理中・返却済などを管理します。"],
    ["入力シート", "お知らせ", "社内共有のお知らせ、重要度、投稿者、既読状況を管理します。", "重要なお知らせは要確認一覧にも表示できます。"],
    ["入力シート", "個人ToDo", "担当者ごとのToDo、期限、状態を管理します。", "未完了ToDoの確認に使います。"],

    ["自動作成シート", "一覧スケジュール", "各入力シートから予定・対応状況を集約するシートです。", "直接編集せず、メニューの全体更新で作成します。"],
    ["自動作成シート", "要確認一覧", "期限や未対応状態から、確認が必要なものだけを集めるシートです。", "直接編集せず、全体更新または要確認一覧更新で作成します。"],
    ["自動作成シート", "担当別未読", "担当者ごとの未読・未完了状況を集計します。", "自動集計用です。"],
    ["自動作成シート", "既読率集計", "既読や個人確認列の状況を集計します。", "共有状況の確認に使います。"],
    ["自動作成シート", "ダッシュボード", "今日の予定、期限切れ、未読件数、未対応電話などの件数を表示します。", "全体状況をひと目で確認できます。"],
    ["自動作成シート", "過去一覧", "完了済み・処理済みの過去データを保存します。", "月次処理で移動します。"],

    ["操作方法", "普段の入力", "各入力シートまたはAppSheetから必要事項を入力します。", "入力直後に一覧へ即時反映しない場合があります。"],
    ["操作方法", "普段の更新", "社内管理 → 日常運用 → 全体更新（軽量）を実行します。", "一覧スケジュール、要確認一覧、集計、ダッシュボードを更新します。"],
    ["操作方法", "設定変更後", "担当者・既読確認者を変更した場合は、設定シートを更新し、入力シートだけ設定反映を実行します。", "必要に応じて個人確認グループを作り直します。"],
    ["操作方法", "月次処理", "過去予定を過去一覧へ移動します。", "本番ではバックアップ後に実行することを推奨します。"],
    ["操作方法", "PDF出力", "一覧帳簿PDFや日報PDFをメニューから作成します。", "PDFリンク列に出力先URLが保存されます。"],
    ["操作方法", "AppSheet連携", "列追加や列名変更後は、AppSheet側でRegenerate Structureを実行します。", "ID列はKey、PDFリンクはURL、既読列はYes/Noにします。"],

    ["運用ルール", "直接編集しないシート", "一覧スケジュール、要確認一覧、担当別未読、既読率集計、ダッシュボード、過去一覧は基本的に直接編集しません。", "入力元シートを直してから全体更新します。"],
    ["運用ルール", "本番で押さないメニュー", "テスト用サンプルデータ追加は本番では押さないでください。", "テスト環境だけで使います。"],
    ["運用ルール", "過去移動の注意", "未対応電話、未返却備品、未完了ToDoなどは安全のため過去一覧へ移動しない仕様です。", "完了・返却済・更新済などになってから移動します。"],
    ["運用ルール", "空欄に戻す", "プルダウンで空欄に戻すを選ぶと、onEdit処理で空欄に戻します。", "Google Sheetsの空欄プルダウン対策です。"],
    ["運用ルール", "個人確認列", "既読列は表示し、個人確認列だけ折りたたみ・縦書きにします。", "横スクロールを減らすための表示設計です。"],
    ["運用ルール", "バックアップ", "大きな設定変更や月次処理の前にはバックアップを推奨します。", "社内管理メニューからバックアップできます。"],

    ["トラブル対応", "一覧に反映されない", "全体更新（軽量）を実行してください。", "onEditを軽くするため、入力直後に自動反映しない設計です。"],
    ["トラブル対応", "担当者が変わらない", "設定シートの担当者・既読確認者を確認し、入力シートだけ設定反映を実行してください。", "担当者と既読確認者は別管理です。"],
    ["トラブル対応", "PDFリンクがTRUE/FALSEになる", "日報PDFリンク列を修復を実行してください。", "PDFリンク列をURL用に戻します。"],
    ["トラブル対応", "フィルタ作成エラー", "既存フィルタを安全に削除して再作成する処理を入れています。", "それでも出る場合は該当シートのフィルタを手動解除します。"],
    ["トラブル対応", "AppSheetで列が違う", "AppSheet側でRegenerate Structureを実行してください。", "列追加・列名変更後に必要です。"],
    ["トラブル対応", "入力規則エラー", "プルダウン候補にない値が入っている可能性があります。", "設定シートや状態候補を確認します。"],

    ["共有・運用", "個人情報管理", "社外共有時は、実名、会社名、地名、車番、取引先名、メールアドレス、電話番号などを含めないでください。", "社外共有・説明用では架空名・架空データを使います。"],
    ["共有・運用", "システム概要", "このシステムは、社内業務をGoogle Workspaceで一元管理する業務改善システムです。", "単なる表ではなく、入力・集約・通知・既読・PDF・AppSheet連携まで含む構成です。"],
    ["共有・運用", "引き継ぎ", "手順シートと社内管理説明を残しておくことで、他の人でも運用手順を確認できます。", "本番では非表示にしても削除しない方がよいです。"]
  ];

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, rows.length, rows[0].length)
    .setWrap(true)
    .setVerticalAlignment("middle");

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 520);
  sheet.setColumnWidth(4, 520);

  const colorMap = {
    "概要": "#eaf4ff",
    "主な機能": "#fff2cc",
    "入力シート": "#fce4d6",
    "自動作成シート": "#e2f0d9",
    "操作方法": "#e4dfec",
    "運用ルール": "#f4cccc",
    "トラブル対応": "#d9ead3",
    "共有・運用": "#d9eaf7"
  };

  for (let r = 2; r <= rows.length; r++) {
    const category = rows[r - 1][0];
    sheet.getRange(r, 1, 1, rows[0].length).setBackground(colorMap[category] || "#ffffff");
  }

  sheet.getRange(1, 1, rows.length, rows[0].length)
    .setBorder(true, true, true, true, true, true);

  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    sheet.getRange(1, 1, rows.length, rows[0].length).createFilter();
  } catch (e) {
    console.log("社内管理説明シートのフィルタ作成をスキップ: " + e.message);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("社内管理説明シートを作成しました");
}











/**
 * v9.6.12 表示・シート順調整
 * - 一覧スケジュールの個人確認列も縦書き対象に含める
 * - 設定シート・既読率集計の列幅を固定する
 * - 運用で見やすい順番にシートタブを並べる
 */
function getPersonalCheckVerticalTargetSheets_() {
  return [
    '一覧スケジュール',
    '要確認一覧',
    '出先予定',
    '工事予定',
    '会議予定',
    '行事予定',
    '作業状況',
    '電話履歴',
    '車検管理',
    '社用車予約',
    '日報',
    '日報レシート管理',
    '免許資格管理',
    '備品修理管理',
    'お知らせ',
    '個人ToDo',
    '過去一覧'
  ];
}

function formatSettingsSheet_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('設定');
  if (!sheet) return;

  try {
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 120); // 設定種別
    sheet.setColumnWidth(2, 150); // 値
    sheet.setColumnWidth(3, 520); // 備考

    const lastRow = Math.max(sheet.getLastRow(), 20);
    sheet.getRange(1, 1, Math.min(lastRow, sheet.getMaxRows()), Math.min(3, sheet.getMaxColumns()))
      .setWrap(true)
      .setVerticalAlignment('middle');

    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#d9ead3')
      .setHorizontalAlignment('center');

    if (sheet.getMaxRows() >= 2) {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, 3)
        .setHorizontalAlignment('left')
        .setVerticalAlignment('middle');
    }
  } catch (e) {
    console.log('設定シートの表示幅調整をスキップ: ' + e.message);
  }
}

function formatReadRateSummarySheet_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('既読率集計');
  if (!sheet) return;

  try {
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150); // 担当
    sheet.setColumnWidth(2, 95);  // 対象件数
    sheet.setColumnWidth(3, 95);  // 既読件数
    sheet.setColumnWidth(4, 95);  // 未読件数
    sheet.setColumnWidth(5, 95);  // 既読率

    const lastRow = Math.max(sheet.getLastRow(), 10);
    sheet.getRange(1, 1, Math.min(lastRow, sheet.getMaxRows()), Math.min(5, sheet.getMaxColumns()))
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');

    sheet.getRange(1, 1, 1, 5)
      .setFontWeight('bold')
      .setBackground('#d9ead3');

    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).setHorizontalAlignment('left');
      sheet.getRange(2, 2, sheet.getLastRow() - 1, 4).setHorizontalAlignment('center');
      sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).setNumberFormat('0.0%');
    }
  } catch (e) {
    console.log('既読率集計の表示幅調整をスキップ: ' + e.message);
  }
}

function getOperationalSheetOrder_() {
  return [
    'ダッシュボード',
    '一覧スケジュール',
    '要確認一覧',
    '出先予定',
    '工事予定',
    '会議予定',
    '行事予定',
    '作業状況',
    '電話履歴',
    '車検管理',
    '車検履歴',
    '社用車予約',
    '日報',
    '日報レシート管理',
    '免許資格管理',
    '備品修理管理',
    'お知らせ',
    '個人ToDo',
    '担当別未読',
    '既読率集計',
    '過去一覧',
    '帳簿PDF履歴',
    '帳簿出力設定',
    '設定',
    '手順シート',
    '社内管理説明',
    'SQL設計',
    '移行対応表',
    'SQLサンプル集'
  ];
}

function reorderSheetsForOperation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const order = getOperationalSheetOrder_();

  // 末尾から順に先頭へ移動すると、最終的に order の順番になる。
  for (let i = order.length - 1; i >= 0; i--) {
    const sheet = ss.getSheetByName(order[i]);
    if (!sheet) continue;

    try {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(1);
    } catch (e) {
      console.log(order[i] + ' のシート順調整をスキップ: ' + e.message);
    }
  }

  const dashboard = ss.getSheetByName('ダッシュボード');
  if (dashboard) {
    try {
      ss.setActiveSheet(dashboard);
    } catch (e) {}
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('シート順を整えました');
}




/**
 * v9.6.17 追加修正（安定優先・高速更新版）
 *
 * 目的:
 * - 金額集計は追加しない。日報レシートは入力・確認・未精算管理までにする。
 * - タイムアウト対策として、初回セットアップと全体更新をさらに軽くする。
 * - onEditでは一覧再作成・表示幅調整・シート順調整を実行しない。
 * - 全体更新（軽量）は一覧/要確認/集計/ダッシュボード作成だけに寄せる。
 * - 表示幅、シート順、個人確認グループは必要時に保守メニューから個別実行する。
 */

function createFilterSafely_(sheet, colCount) {
  if (!sheet || !colCount || colCount < 1) return;

  try {
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
  } catch (e) {
    console.log(sheet.getName() + ' の既存フィルタ削除をスキップ: ' + e.message);
  }

  try {
    // v9.6.17: maxRows全体にフィルタをかけると重くなりやすいため、
    // 実データ行 + 余裕行だけに限定する。
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const rowCount = Math.min(sheet.getMaxRows(), Math.max(lastRow + 50, 100));
    sheet.getRange(1, 1, rowCount, colCount).createFilter();
  } catch (e) {
    console.log(sheet.getName() + ' のフィルタ作成をスキップ: ' + e.message);
  }
}

function resetSheet(sheet, headerCount) {
  if (!sheet) return;

  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (e) {}

  try {
    sheet.showColumns(1, sheet.getMaxColumns());
  } catch (e) {}

  try {
    for (let i = 0; i < 5; i++) {
      sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).shiftColumnGroupDepth(-1);
    }
  } catch (e) {}

  // v9.6.17: 全シート全行をclearしない。実データ範囲だけを初期化する。
  // 大量の空行までclearするとタイムアウトしやすいため。
  const rowsToClear = Math.max(sheet.getLastRow(), 2);
  const colsToClear = Math.max(sheet.getLastColumn(), headerCount || 1);
  try {
    sheet.getRange(1, 1, rowsToClear, colsToClear).clearContent().clearFormat().clearDataValidations();
  } catch (e) {
    console.log(sheet.getName() + ' の範囲初期化をスキップ: ' + e.message);
  }

  try {
    sheet.setConditionalFormatRules([]);
  } catch (e) {}

  const currentCols = sheet.getMaxColumns();
  if (headerCount && currentCols > headerCount) {
    try {
      sheet.deleteColumns(headerCount + 1, currentCols - headerCount);
    } catch (e) {}
  } else if (headerCount && currentCols < headerCount) {
    sheet.insertColumnsAfter(currentCols, headerCount - currentCols);
  }
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    showSkipToastOnce_(
      '別の処理中のため、全体更新をスキップしました。少し待ってから再実行してください。',
      'REFRESH_ALL_SKIP_TOAST_V9617',
      10
    );
    return;
  }

  try {
    // v9.6.17: 全体更新は集約中心。表示幅・シート順・グループ作成は別メニューへ分離。
    refreshInputSheetsLight_();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();

    clearSummaryDirty_();
    showDoneToast_('全体更新（軽量・高速）が完了しました。表示が崩れる場合だけ表示幅調整を実行してください。');
  } finally {
    lock.releaseLock();
  }
}

function refreshAllWithDisplayMaintenance() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    showSkipToastOnce_(
      '別の処理中のため、表示調整付き更新をスキップしました。少し待ってから再実行してください。',
      'REFRESH_DISPLAY_SKIP_TOAST_V9617',
      10
    );
    return;
  }

  try {
    refreshInputSheetsLight_();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();

    try {
      sortAllOperationalSheetsSilent_();
    } catch (e) {
      console.log('主要シート並び替えをスキップ: ' + e.message);
    }

    try {
      formatInputSheetsForLongText();
    } catch (e) {
      console.log('表示幅調整をスキップ: ' + e.message);
    }

    try {
      reorderSheetsForOperation();
    } catch (e) {
      console.log('シート順調整をスキップ: ' + e.message);
    }

    clearSummaryDirty_();
    showDoneToast_('表示調整付き更新が完了しました');
  } finally {
    lock.releaseLock();
  }
}

function initialSetupForProduction() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(3000)) {
    showSkipToastOnce_(
      '別の処理中のため、初回セットアップをスキップしました。少し待ってから再実行してください。',
      'INITIAL_SETUP_SKIP_TOAST_V9617',
      10
    );
    return;
  }

  try {
    // v9.6.17: 初回セットアップは最低限にする。
    // 表示幅・シート順・個人確認グループは保守メニューから必要時に個別実行。
    createCompanySheets();
    setupSettingsSheet();
    setupLedgerOutputSettingsSheet();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    applySettingsToInputSheetsOnly();
    fixDailyReportPdfLinkColumn();

    showDoneToast_('初回セットアップ（軽量・高速）が完了しました。次に全体更新（軽量）を実行してください。');
  } finally {
    lock.releaseLock();
  }
}

function createProcedureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('手順シート');
  if (!sheet) sheet = ss.insertSheet('手順シート');

  sheet.clear();
  sheet.clearFormats();

  const headers = ['区分', '手順', '実行メニュー', 'タイミング', '注意'];
  const rows = [
    ['初回', '1', '社内管理 → 1. 初回セットアップ → 初回セットアップ（軽量・高速）', '最初だけ', 'サンプルデータは入りません。処理を軽くするため、表示幅やグループ作成は別メニューに分けています。'],
    ['初回', '2', '社内管理 → 2. 日常運用 → 全体更新（軽量・高速）', '初回セットアップ後', '一覧・要確認・集計・ダッシュボードを作成します。'],
    ['初回', '3', '社内管理 → 3. 保守・テスト → 表示幅を調整', '見た目を整えたい時', 'タイムアウト対策のため、初回セットアップからは分離しています。'],
    ['初回', '4', '社内管理 → 3. 保守・テスト → シート順を整える', 'シート順が崩れた時', '必要な時だけ実行します。'],
    ['初回', '5', '社内管理 → 3. 保守・テスト → 個人確認グループを作り直す', '既読確認者を変更した時', '個人確認列だけ折りたたみます。'],
    ['日常', '1', '入力シートへ入力', '随時', '入力直後は一覧へ即時反映せず、軽量化のため変更あり状態だけ記録します。'],
    ['日常', '2', '社内管理 → 2. 日常運用 → 全体更新（軽量・高速）', '入力後・確認前', '普段はこれを使います。表示幅やシート順は触りません。'],
    ['日常', '3', '社内管理 → 2. 日常運用 → 表示調整付き更新', '更新と見た目調整をまとめたい時', '通常更新より重いので、毎回は使わなくてよいです。'],
    ['日常', '4', '社内管理 → 2. 日常運用 → 一覧帳簿PDFを作成', '帳簿出力時', '帳簿出力設定でTRUEのものだけ出力します。'],
    ['日常', '5', '社内管理 → 2. 日常運用 → 車検更新完了処理', '車検完了後', '次回車検期限を車検期限へ反映し、車検履歴に残します。'],
    ['レシート', '1', '日報レシート管理へ入力', '紙レシートを受け取った時', '金額は記録用です。自動金額集計は安定優先で入れていません。'],
    ['レシート', '2', '全体更新（軽量・高速）', '確認前', '未確認・差戻し・未精算だけ要確認一覧に出します。'],
    ['保守', '1', '社内管理 → 3. 保守・テスト → 日報PDFリンク列を修復', 'PDFリンク列にTRUE/FALSEが出た時', 'PDFリンク列をURL用に戻します。'],
    ['保守', '2', '社内管理 → 3. 保守・テスト → 空欄に戻す表示をクリア', 'プルダウンの空欄戻し文字が残った時', '既に入ってしまった表示を空欄に戻します。'],
    ['テスト', '1', '社内管理 → 3. 保守・テスト → テスト用サンプルデータを追加', 'テスト環境だけ', '本番では押さないでください。'],
    ['テスト', '2', '社内管理 → 3. 保守・テスト → 過去移動ルールをチェック', '移行前確認', '実データは移動せず判定だけ確認します。'],
    ['月次', '1', '社内管理 → 3. 保守・テスト → 過去予定を過去一覧へ移動', '月次・確認後', '本番移行直後は押さないでください。'],
    ['社外共有前確認', '1', '匿名化確認', '社外共有・資料化前', '実名・社名・地名・車番・取引先名が残っていないか確認します。'],
    ['AppSheet', '1', 'Regenerate Structure', '列変更後', 'ID列をKey、PDFリンクをURL、既読列をYes/Noに確認します。']
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center');

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 480);
  sheet.getRange(1, 1, rows.length + 1, headers.length).setWrap(true).setVerticalAlignment('middle');
  createFilterSafely_(sheet, headers.length);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('社内管理')
    .addSubMenu(
      ui.createMenu('1. 初回セットアップ')
        .addItem('初回セットアップ（軽量・高速）', 'initialSetupForProduction')
        .addItem('シート構成を作成', 'createCompanySheets')
        .addItem('設定シートを作成/確認', 'setupSettingsSheet')
        .addItem('設定シートの備考を整える', 'normalizeSettingsSheetNotesMenu')
        .addItem('入力シートだけ設定反映', 'applySettingsToInputSheetsOnly')
        .addItem('帳簿出力設定を作成/確認', 'setupLedgerOutputSettingsSheet')
        .addItem('手順シートを作成', 'createProcedureSheet')
        .addItem('社内管理説明シートを作成', 'createCompanyManagementExplanationSheet')
    )
    .addSubMenu(
      ui.createMenu('2. 日常運用')
        .addItem('全体更新（軽量・高速）', 'refreshAll')
        .addItem('表示調整付き更新', 'refreshAllWithDisplayMaintenance')
        .addItem('主要シートを日時順に並び替え', 'sortAllOperationalSheets')
        .addItem('一覧帳簿PDFを作成', 'createLedgerPdf')
        .addSeparator()
        .addItem('選択行の日報 他現場状況を作成', 'fillDailyReportOtherSituationForSelectedRow')
        .addItem('選択行の日報文章をGPT作成', 'createDailyReportText')
        .addItem('選択行の日報PDFを作成', 'createSelectedDailyReportPdf')
        .addItem('車検更新完了処理', 'processCompletedVehicleInspections')
    )
    .addSubMenu(
      ui.createMenu('3. 保守・テスト')
        .addItem('個人確認グループを作り直す', 'rebuildCheckGroups')
        .addItem('個人確認列を縦書きにする', 'formatPersonalCheckColumnsVertical')
        .addItem('日報PDFリンク列を修復', 'fixDailyReportPdfLinkColumn')
        .addItem('空欄に戻す表示をクリア', 'clearClearLabelValuesInAllSheets')
        .addItem('表示幅を調整', 'formatInputSheetsForLongText')
        .addItem('シート順を整える', 'reorderSheetsForOperation')
        .addItem('裏方シートを非表示', 'hideSupportSheets')
        .addItem('スプレッドシートをバックアップ', 'backupSpreadsheet')
        .addSeparator()
        .addItem('テスト用サンプルデータを追加', 'addSampleDataForTest')
        .addItem('過去移動ルールをチェック', 'checkArchiveRulesForTest')
        .addItem('移行前チェックを実行', 'runPreMigrationCheck')
        .addSeparator()
        .addItem('SQL関連シートを作成', 'createSqlSupportSheets')
        .addItem('全体更新（重い）', 'refreshAllHeavy')
        .addItem('現在のシートを月ごとに折りたたみ', 'groupRowsByMonth')
        .addItem('全シートを月ごとに折りたたみ', 'groupAllRowsByMonth')
        .addItem('過去予定を過去一覧へ移動', 'movePastItemsToArchive')
    )
    .addToUi();
}







/**
 * ============================================================
 * v9.6.30 追記・上書きブロック
 * - ユーザー提示のベースコードを元に、前シート前提ではなく空の新規ベータを作れるようにする。
 * - v9.6.24〜v9.6.29相当のフィードバック反映をこのベースへ戻す。
 * - 工事予定/行事予定/電話履歴/車検管理/運転免許管理/資格管理/備品修理管理/フィードバックを新構成へ整理。
 * - 空欄行の大量チェックボックスを避け、入力行だけチェックボックスを表示。
 * - SQL関連シートを作成・表示できるようにする。
 * ============================================================
 */

function getIdHeaderForSheet_(sheetName) {
  const map = {
    "出先予定": "出先ID",
    "工事予定": "工事ID",
    "会議予定": "会議ID",
    "行事予定": "行事ID",
    "作業状況": "作業ID",
    "電話履歴": "電話ID",
    "車検管理": "車両ID",
    "車検履歴": "履歴ID",
    "社用車予約": "予約ID",
    "日報": "日報ID",
    "日報レシート管理": "レシートID",
    "運転免許管理": "運転免許ID",
    "資格管理": "資格ID",
    "免許資格管理": "資格ID",
    "備品修理管理": "修理ID",
    "お知らせ": "お知らせID",
    "個人ToDo": "ToDo_ID",
    "フィードバック": "フィードバックID",
    "帳簿PDF履歴": "帳簿ID"
  };
  return map[sheetName] || "";
}

function buildRecordId_(sheetName) {
  const prefixMap = {
    "出先予定": "OUT",
    "工事予定": "CON",
    "会議予定": "MTG",
    "行事予定": "EVT",
    "作業状況": "WRK",
    "電話履歴": "TEL",
    "車検管理": "CAR",
    "車検履歴": "CARH",
    "社用車予約": "RES",
    "日報": "DR",
    "日報レシート管理": "RCT",
    "運転免許管理": "DRV",
    "資格管理": "LIC",
    "免許資格管理": "LIC",
    "備品修理管理": "REP",
    "お知らせ": "NEWS",
    "個人ToDo": "TODO",
    "フィードバック": "FB",
    "帳簿PDF履歴": "LEDGER"
  };
  const prefix = prefixMap[sheetName] || "ID";
  const nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  return prefix + "-" + nowText + "-" + Utilities.getUuid().slice(0, 8);
}

function getSheetHeaders_() {
  const personalMembers = getPersonalMembers_();

  return {
    "ダッシュボード": ["項目", "件数", "備考"],
    "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "出先ID"],
    "工事予定": ["工事名", "現場", "依頼主", "連絡先", "契約金額", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "工事ID"],
    "会議予定": ["日付", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "会議ID"],
    "行事予定": ["開始日", "終了日", "開始時刻", "終了時刻", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "行事ID"],
    "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...personalMembers, "備考", "作業ID"],
    "電話履歴": ["日付", "時間", "相手", "内容", "電話対応", "担当", "対応メモ", "通知", READ_HEADER, ...personalMembers, "備考", "電話ID"],
    "車検管理": ["車両名", "車番", "車検期限", "車検予定日", "通知", "保険期限", "車検状態", "写真", READ_HEADER, ...personalMembers, "備考", "車両ID"],
    "車検履歴": ["更新日", "車両名", "車番", "旧車検期限", "新車検期限", "旧保険期限", "新保険期限", "担当", "備考", "履歴ID"],
    "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...personalMembers, "備考", CALENDAR_ID_HEADER, "予約ID"],
    "日報": ["日付", "入力者", "担当", "現場", "作業内容", "進捗", "問題点", "明日の予定", "他現場状況", "写真", "日報文章", "状態", "備考", READ_HEADER, ...personalMembers, "PDFリンク", "日報ID"],
    "日報レシート管理": ["日付", "担当", "現場", "支払先", "内容", "区分", "金額", "支払方法", "レシート写真", "経理確認", "精算状態", "通知", "備考", "レシートID"],
    "運転免許管理": ["所有者", "普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "運転免許ID"],
    "資格管理": ["所有者", "資格名", "区分", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "資格ID"],
    "備品修理管理": ["購入日", "備品名", "修理業者", "内容", "修理依頼者", "修理依頼日", "返却予定日", "返却済", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "修理ID"],
    "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...personalMembers, "備考", "お知らせID"],
    "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...personalMembers, "備考", "ToDo_ID"],
    "フィードバック": ["日付", "記入者", "確認方法", "対象シート", "気になった内容", "区分", "困り度", "対応状況", "対応方針", "対応メモ", "対応日", "フィードバックID"],
    "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...personalMembers, "備考", "元シート"],
    "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "備考", "元シート"],
    "担当別未読": ["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"],
    "既読率集計": ["対象", "全体件数", "既読件数", "既読率"],
    "過去一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート", "備考", "移動日"],
    "帳簿PDF履歴": ["作成日時", "対象", "期間", "PDFリンク", "備考", "帳簿ID"],
    "帳簿出力設定": ["出力する", "帳簿名", "シート名", "出力列", "最大行数", "備考"]
  };
}

function getStatusListForSheet_(sheetName) {
  const map = {
    "出先予定": [CLEAR_LABEL, "予定", "移動中", "完了", "延期", "中止"],
    "工事予定": [CLEAR_LABEL, "予定", "着工前", "施工中", "完了", "延期", "中止"],
    "会議予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "行事予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "作業状況": [CLEAR_LABEL, "着工前", "施工中", "完了", "修理不可", "要確認"],
    "電話履歴": [CLEAR_LABEL],
    "社用車予約": [CLEAR_LABEL, "予定", "使用中", "返却済", "中止", "予約重複"],
    "日報": [CLEAR_LABEL, "下書き", "提出済", "確認済", "差戻し", "完了"],
    "日報レシート管理": [CLEAR_LABEL, "未確認", "確認中", "確認済", "差戻し", "未精算", "精算済"],
    "運転免許管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "資格管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "免許資格管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "備品修理管理": [CLEAR_LABEL, "未依頼", "依頼済", "修理中", "返却済", "完了", "修理不可"],
    "車検管理": [CLEAR_LABEL, "未対応", "予約済", "実施済", "更新済"],
    "個人ToDo": [CLEAR_LABEL, "未着手", "対応中", "完了", "中止"],
    "お知らせ": [CLEAR_LABEL, "要確認", "確認済", "完了"]
  };
  return map[sheetName] || [CLEAR_LABEL, "予定", "進行中", "完了", "延期", "中止", "未対応", "対応中", "要確認"];
}

function getVehicleInspectionStatusList_() {
  return [CLEAR_LABEL, "未対応", "予約済", "実施済", "更新済"];
}

function getFeedbackMethodList_() {
  return [CLEAR_LABEL, "PC", "タブレット", "スマホ", "AppSheet", "紙運用との比較", "口頭確認", "その他"];
}

function getFeedbackCategoryList_() {
  return [CLEAR_LABEL, "不具合", "入力しづらい", "表示が見づらい", "項目追加", "項目削除", "運用確認", "便利そう", "その他"];
}

function getFeedbackSeverityList_() {
  return [CLEAR_LABEL, "高", "中", "低", "メモ"];
}

function getFeedbackStatusList_() {
  return [CLEAR_LABEL, "未確認", "確認中", "対応予定", "対応済", "見送り"];
}

function isDateHeaderV9630_(header) {
  return [
    "日付",
    "日時",
    "作成日時",
    "更新日",
    "移動日",
    "開始日",
    "終了日",
    "車検期限",
    "車検予定日",
    "保険期限",
    "投稿日",
    "登録日",
    "購入日",
    "期限",
    "取得日",
    "更新期限",
    "修理依頼日",
    "返却予定日",
    "対応日"
  ].includes(header);
}

function isTimeHeaderV9630_(header) {
  return ["時間", "開始時刻", "終了時刻"].includes(header);
}

function isMoneyHeaderV9630_(header) {
  return ["金額", "契約金額"].includes(header);
}

function getApplyRowCount_(sheet) {
  if (!sheet) return 80;
  const maxRows = sheet.getMaxRows();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const targetRows = Math.min(maxRows, Math.max(81, lastRow + 40));
  return Math.max(targetRows - 1, 1);
}

function ensureSheetMinimumRows_(sheet) {
  if (!sheet) return;
  const targetRows = 101;
  const maxRows = sheet.getMaxRows();
  if (maxRows < targetRows) {
    sheet.insertRowsAfter(maxRows, targetRows - maxRows);
  }
}

function createFilterSafely_(sheet, colCount) {
  if (!sheet || !colCount || colCount < 1) return;
  try {
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
  } catch (e) {}
  try {
    const lastRow = Math.max(2, Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow() + 20, 30)));
    sheet.getRange(1, 1, lastRow, colCount).createFilter();
  } catch (e) {
    console.log(sheet.getName() + " のフィルタ作成をスキップ: " + e.message);
  }
}

function setTimeColumnV9630_(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount).setNumberFormat("hh:mm");
}

function setMoneyColumnV9630_(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount).setNumberFormat("#,##0");
}

function setupSheetWithoutClearing_(sheet, headers) {
  if (!sheet || !headers) return;

  ensureSheetMinimumRows_(sheet);

  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < headers.length) {
    sheet.insertColumnsAfter(currentMaxCols, headers.length - currentMaxCols);
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  createFilterSafely_(sheet, headers.length);

  const sheetName = sheet.getName();
  const rowCount = getApplyRowCount_(sheet);

  try {
    sheet.getRange(1, 1, rowCount + 1, headers.length).clearDataValidations();
  } catch (e) {}

  const statusList = getStatusListForSheet_(sheetName);
  const staffList = [CLEAR_LABEL, ...getStaffMembers_()];
  const carList = [CLEAR_LABEL, "4t", "軽トラ②", "8tユニック", "軽ワゴン③", "ローダー", "P.BOX①", "P.BOX②", "3t", "キャブ③", "軽トラ③", "2t①", "軽ワゴン⑥", "軽バン⑤", "軽ダンプ②"];
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
  const importanceList = [CLEAR_LABEL, "高", "中", "低"];

  headers.forEach((header, i) => {
    const col = i + 1;

    if (!isAutoOutputSheet_(sheetName)) {
      if (isDateHeaderV9630_(header)) setDateColumn(sheet, col);
      if (isTimeHeaderV9630_(header)) setTimeColumnV9630_(sheet, col);
      if (isMoneyHeaderV9630_(header)) setMoneyColumnV9630_(sheet, col);

      if (header === "状態") setDropdown(sheet, col, statusList);
      if (header === "車検状態") setDropdown(sheet, col, getVehicleInspectionStatusList_());
      if (header === "担当" || header === "入力者" || header === "利用者" || header === "投稿者" || header === "修理依頼者" || header === "所有者" || header === "記入者") {
        setDropdown(sheet, col, staffList);
      }
      if (header === "社用車") setDropdown(sheet, col, carList);
      if (header === "電話対応") setDropdown(sheet, col, phoneList);
      if (header === "重要度") setDropdown(sheet, col, importanceList);
      if (header === "コピー有無") setDropdown(sheet, col, getCopyStatusList_());
      if (header === "返却済") setDropdown(sheet, col, getReturnedStatusList_());
      if (header === "区分" && sheetName === "日報レシート管理") setDropdown(sheet, col, getReceiptCategoryList_());
      if (header === "支払方法") setDropdown(sheet, col, getReceiptPaymentMethodList_());
      if (header === "経理確認") setDropdown(sheet, col, getReceiptAccountingCheckList_());
      if (header === "精算状態") setDropdown(sheet, col, getReceiptSettlementStatusList_());

      if (sheetName === "フィードバック") {
        if (header === "確認方法") setDropdown(sheet, col, getFeedbackMethodList_());
        if (header === "対象シート") setDropdown(sheet, col, [CLEAR_LABEL, ...Object.keys(getSheetHeaders_()).filter(n => !["ダッシュボード"].includes(n))]);
        if (header === "区分") setDropdown(sheet, col, getFeedbackCategoryList_());
        if (header === "困り度") setDropdown(sheet, col, getFeedbackSeverityList_());
        if (header === "対応状況") setDropdown(sheet, col, getFeedbackStatusList_());
      }
    }
  });

  setCheckboxesForActiveDataRowsOnly_(sheet);
  ensureIdsForSheet_(sheet);
  applyColorRules(sheet);
  hideTechnicalColumns_(sheet);

  const calendarIdCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
  if (calendarIdCol > 0) {
    try { sheet.hideColumns(calendarIdCol); } catch (e) {}
  }
}

function setupSheet(sheet, headers) {
  setupSheetWithoutClearing_(sheet, headers);
}

function getCheckboxHeadersForSheetV9630_(sheetName) {
  if (sheetName === "運転免許管理") {
    return ["普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種", READ_HEADER, ...getPersonalMembers_()];
  }
  return [READ_HEADER, ...getPersonalMembers_()];
}

function isRowDataActiveForCheckboxV9630_(sheetName, headers, rowValues) {
  const idHeader = getIdHeaderForSheet_(sheetName);
  const checkboxHeaders = getCheckboxHeadersForSheetV9630_(sheetName);

  return rowValues.some((v, idx) => {
    const h = headers[idx];
    if (!h) return false;
    if (h === idHeader) return false;
    if (h === CALENDAR_ID_HEADER) return false;
    if (checkboxHeaders.includes(h)) return false;
    if (h === "通知") return false;
    return v !== "" && v !== null;
  });
}

function setCheckboxesForActiveDataRowsOnly_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const sheetName = sheet.getName();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const checkboxHeaders = getCheckboxHeadersForSheetV9630_(sheetName);
  const targetCols = checkboxHeaders
    .map(h => headers.indexOf(h) + 1)
    .filter(col => col > 0);

  if (targetCols.length === 0) return;

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const maxCleanRow = Math.min(sheet.getMaxRows(), Math.max(lastRow + 20, 40));
  const values = sheet.getRange(2, 1, maxCleanRow - 1, headers.length).getValues();

  targetCols.forEach(col => {
    const range = sheet.getRange(2, col, maxCleanRow - 1, 1);
    const colValues = range.getValues();

    for (let r = 0; r < values.length; r++) {
      const active = isRowDataActiveForCheckboxV9630_(sheetName, headers, values[r]);
      const cell = sheet.getRange(r + 2, col);
      if (active) {
        cell.insertCheckboxes();
        const current = colValues[r][0];
        if (current === "" || current === null) cell.setValue(false);
      } else {
        cell.clearContent();
        cell.clearDataValidations();
      }
    }
  });
}

function setCheckboxesForEditedRowSmart_(sheet, row) {
  if (!sheet || row <= 1 || sheet.getLastColumn() < 1) return;

  const sheetName = sheet.getName();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const active = isRowDataActiveForCheckboxV9630_(sheetName, headers, values);
  const checkboxHeaders = getCheckboxHeadersForSheetV9630_(sheetName);

  checkboxHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;
    const cell = sheet.getRange(row, col);
    if (active) {
      cell.insertCheckboxes();
      const current = cell.getValue();
      if (current === "" || current === null) cell.setValue(false);
    } else {
      cell.clearContent();
      cell.clearDataValidations();
    }
  });
}

function cleanupBlankRowCheckboxesCurrentSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  setCheckboxesForActiveDataRowsOnly_(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast("現在のシートの空欄行チェックボックスを整理しました");
}

function createCompanySheetsNewBlankNoLegacy() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、新規作成をスキップしました");
    return;
  }

  try {
    clearSettingsCache_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const response = ui.alert(
      "新規ベータ作成",
      "現在のスプレッドシート内の既存シートを初期化して、最新版構成で作り直します。空の新規スプレッドシート用です。実行しますか？",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ss.toast("新規ベータ作成をキャンセルしました");
      return;
    }

    let temp = ss.getSheetByName("_TEMP_SETUP_");
    if (!temp) temp = ss.insertSheet("_TEMP_SETUP_");

    ss.getSheets().forEach(sheet => {
      if (sheet.getName() !== "_TEMP_SETUP_") {
        ss.deleteSheet(sheet);
      }
    });

    ensureSettingsSheet_();

    const headersMap = getSheetHeaders_();
    getOperationalSheetOrder_().forEach(name => {
      if (!headersMap[name]) return;
      let sheet = ss.getSheetByName(name);
      if (!sheet) sheet = ss.insertSheet(name);
      setupHeaderOnly_(sheet, headersMap[name]);
    });

    Object.keys(headersMap).forEach(name => {
      if (ss.getSheetByName(name)) return;
      const sheet = ss.insertSheet(name);
      setupHeaderOnly_(sheet, headersMap[name]);
    });

    const tempSheet = ss.getSheetByName("_TEMP_SETUP_");
    if (tempSheet && ss.getSheets().length > 1) ss.deleteSheet(tempSheet);

    createFeedbackSheet();
    createSqlSupportSheetsVisible();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    setupInitialLedgerOutputSettings_();

    applySettingsToSheetsByNames_(getPrimaryInputSheetNamesForSetup_(), "入力シートを整備しました");
    reorderSheetsForOperation();
    formatInputSheetsForLongText();
    refreshAll();

    ss.toast("新規ベータ作成が完了しました");
  } finally {
    lock.releaseLock();
  }
}

function createFeedbackSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["フィードバック"];
  let sheet = ss.getSheetByName("フィードバック");
  if (!sheet) sheet = ss.insertSheet("フィードバック");

  setupSheetWithoutClearing_(sheet, headers);

  if (sheet.getLastRow() < 2) {
    const today = new Date();
    const sampleRows = [
      {
        "日付": today,
        "記入者": "",
        "確認方法": "PC",
        "対象シート": "一覧スケジュール",
        "気になった内容": "ベータサンプル：一覧で必要な情報が見やすいか確認してください。",
        "区分": "運用確認",
        "困り度": "メモ",
        "対応状況": "未確認",
        "対応方針": "",
        "対応メモ": "不要ならこの行は削除してください。"
      }
    ];
    writeObjectsToSheetNoCheckboxAllRows_(sheet, sampleRows);
  }

  formatSheetByName_(sheet, "フィードバック");
  return sheet;
}

function addFeedbackExampleRows() {
  const sheet = createFeedbackSheet();
  const today = new Date();
  writeObjectsToSheetNoCheckboxAllRows_(sheet, [
    {
      "日付": today,
      "確認方法": "AppSheet",
      "対象シート": "日報",
      "気になった内容": "入力欄が多い、または現場で入力しづらい項目があればここに記入します。",
      "区分": "入力しづらい",
      "困り度": "中",
      "対応状況": "未確認"
    },
    {
      "日付": today,
      "確認方法": "PC",
      "対象シート": "車検管理",
      "気になった内容": "車検状態や車検予定日の意味が分かりにくい場合はここに記入します。",
      "区分": "表示が見づらい",
      "困り度": "低",
      "対応状況": "未確認"
    }
  ]);
  setCheckboxesForActiveDataRowsOnly_(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast("フィードバック記入例を追加しました");
}

function writeObjectsToSheetNoCheckboxAllRows_(sheet, objects) {
  if (!sheet || !objects || objects.length === 0) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  const values = objects.map(obj => headers.map(header => obj[header] !== undefined ? obj[header] : ""));
  sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
  ensureIdsForSheet_(sheet);
  setCheckboxesForActiveDataRowsOnly_(sheet);
  applyColorRules(sheet);
}

function createSqlSupportSheetsVisible() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    createSqlSupportSheets();
  } catch (e) {
    console.log("既存SQL作成処理をスキップ: " + e.message);
  }

  const sqlNames = ["SQL設計", "移行対応表", "SQLサンプル集"];
  sqlNames.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.showSheet();
    if (sheet.getLastRow() === 0 || (sheet.getLastRow() === 1 && !sheet.getRange(1, 1).getValue())) {
      sheet.getRange(1, 1).setValue(name);
    }
  });

  reorderSheetsForOperation();
  SpreadsheetApp.getActiveSpreadsheet().toast("SQL関連シートを作成/表示しました");
}

function getPrimaryInputSheetNamesForSetup_() {
  return [
    "出先予定",
    "工事予定",
    "電話履歴",
    "車検管理",
    "社用車予約",
    "日報",
    "日報レシート管理",
    "運転免許管理",
    "資格管理",
    "備品修理管理",
    "お知らせ",
    "個人ToDo",
    "会議予定",
    "行事予定",
    "作業状況",
    "フィードバック"
  ];
}

function getOutputAndSupportSheetNamesForSetup_() {
  return [
    "一覧スケジュール",
    "要確認一覧",
    "車検履歴",
    "帳簿PDF履歴",
    "帳簿出力設定",
    "担当別未読",
    "既読率集計",
    "過去一覧",
    "ダッシュボード",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集",
    "手順シート",
    "社内管理説明"
  ];
}

function getOperationalSheetOrder_() {
  return [
    "ダッシュボード",
    "一覧スケジュール",
    "要確認一覧",
    "フィードバック",
    "出先予定",
    "工事予定",
    "会議予定",
    "行事予定",
    "作業状況",
    "電話履歴",
    "車検管理",
    "社用車予約",
    "日報",
    "日報レシート管理",
    "運転免許管理",
    "資格管理",
    "免許資格管理",
    "備品修理管理",
    "お知らせ",
    "個人ToDo",
    "担当別未読",
    "既読率集計",
    "過去一覧",
    "車検履歴",
    "帳簿PDF履歴",
    "帳簿出力設定",
    "設定",
    "手順シート",
    "社内管理説明",
    "SQL設計",
    "移行対応表",
    "SQLサンプル集"
  ];
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
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "出先予定"
  }), rows);

  collectRowsAsObjects(ss, "工事予定", row => row["開始日"] || row["工事名"], row => ({
    "日付": row["開始日"] || new Date(),
    "種類": "工事予定",
    "内容": joinText(row["工事名"], joinText(row["現場"], row["依頼主"])),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": row["状態"] === "完了" ? "" : getNoticeText(row["開始日"]),
    "電話対応": row["電話対応"],
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || joinText(row["連絡先"], row["契約金額"]),
    "元シート": "工事予定"
  }), rows);

  collectRowsAsObjects(ss, "会議予定", row => row["日付"], row => ({
    "日付": row["日付"],
    "種類": "会議",
    "内容": joinText(row["会議名"], row["内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "会議予定"
  }), rows);

  collectRowsAsObjects(ss, "行事予定", row => row["開始日"] || row["行事名"], row => ({
    "日付": row["開始日"] || new Date(),
    "種類": "行事",
    "内容": joinText(row["行事名"], joinText(row["開始時刻"], row["内容"])),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["開始日"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "行事予定"
  }), rows);

  collectRowsAsObjects(ss, "電話履歴", row => row["日付"] || row["相手"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "電話履歴",
    "内容": joinText(row["相手"], row["内容"]),
    "担当": row["担当"],
    "状態": row["電話対応"] || "",
    "通知": row["電話対応"] === "完了" ? "" : getNoticeText(row["日付"]),
    "電話対応": row["電話対応"],
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || row["対応メモ"] || "",
    "元シート": "電話履歴"
  }), rows);

  collectRowsAsObjects(ss, "車検管理", row => row["車検期限"] || row["車両名"], row => ({
    "日付": row["車検期限"] || new Date(),
    "種類": "車検",
    "内容": joinText(row["車両名"], row["車番"]),
    "担当": "",
    "状態": row["車検状態"] || row["状態"] || "",
    "通知": (row["車検状態"] === "更新済") ? "" : getNoticeText(row["車検期限"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || joinText("車検予定日", row["車検予定日"]),
    "元シート": "車検管理"
  }), rows);

  collectRowsAsObjects(ss, "社用車予約", row => row["日付"] || row["行き先"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "社用車予約",
    "内容": joinText(row["社用車"], joinText(row["行き先"], row["用途"])),
    "担当": row["利用者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "社用車予約"
  }), rows);

  collectRowsAsObjects(ss, "日報", row => row["日付"] || row["現場"] || row["作業内容"], row => ({
    "日付": row["日付"] || new Date(),
    "種類": "日報",
    "内容": joinText(row["現場"], row["作業内容"]),
    "担当": row["担当"] || row["入力者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["日付"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || row["他現場状況"] || "",
    "元シート": "日報"
  }), rows);

  collectRowsAsObjects(ss, "運転免許管理", row => row["更新期限"] || row["所有者"], row => ({
    "日付": row["更新期限"] || new Date(),
    "種類": "運転免許",
    "内容": joinText(row["所有者"], getCheckedLicenseTypesText_(row)),
    "担当": row["所有者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["更新期限"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "運転免許管理"
  }), rows);

  collectRowsAsObjects(ss, "資格管理", row => row["更新期限"] || row["資格名"], row => ({
    "日付": row["更新期限"] || new Date(),
    "種類": "資格",
    "内容": joinText(row["所有者"], joinText(row["資格名"], row["区分"])),
    "担当": row["所有者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["更新期限"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "資格管理"
  }), rows);

  collectRowsAsObjects(ss, "備品修理管理", row => row["返却予定日"] || row["備品名"], row => ({
    "日付": row["返却予定日"] || row["修理依頼日"] || row["購入日"] || new Date(),
    "種類": "備品修理",
    "内容": joinText(row["備品名"], joinText(row["修理業者"], row["内容"])),
    "担当": row["修理依頼者"],
    "状態": row["状態"],
    "通知": getNoticeText(row["返却予定日"] || row["修理依頼日"] || row["購入日"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "備品修理管理"
  }), rows);

  collectRowsAsObjects(ss, "個人ToDo", row => row["期限"] || row["内容"], row => ({
    "日付": row["期限"] || row["登録日"] || new Date(),
    "種類": "ToDo",
    "内容": row["内容"],
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": getNoticeText(row["期限"]),
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "個人ToDo"
  }), rows);

  collectRowsAsObjects(ss, "作業状況", row => row["現場"] || row["作業内容"], row => ({
    "日付": new Date(),
    "種類": "作業状況",
    "内容": joinText(row["現場"], row["作業内容"]),
    "担当": row["担当"],
    "状態": row["状態"],
    "通知": row["通知"] || "",
    "電話対応": "",
    "既読": row["既読"] === true || row["既読"] === "TRUE" || row["既読"] === "Y",
    ...copyPersonalChecks(row),
    "備考": row["備考"] || "",
    "元シート": "作業状況"
  }), rows);

  rows.sort((a, b) => new Date(a["日付"]) - new Date(b["日付"]));

  writeObjectsToSheet(sheet, rows);
  applyColorRules(sheet);
}

function getCheckedLicenseTypesText_(row) {
  return ["普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種"]
    .filter(name => row[name] === true || row[name] === "TRUE" || row[name] === "Y")
    .join(" / ");
}

function shouldIncludeInAlertList_(row, priority) {
  const sourceName = row["元シート"] || row["種類"] || "";
  const status = String(row["状態"] || "").trim();
  const notice = String(row["通知"] || "").trim();
  const phone = String(row["電話対応"] || "").trim();

  if (sourceName === "お知らせ" || row["種類"] === "お知らせ") return false;
  if (!priority[notice]) return false;

  if (sourceName === "電話履歴" || row["種類"] === "電話履歴") return phone !== "完了";
  if (sourceName === "備品修理管理" || row["種類"] === "備品修理") return !["返却済", "完了", "修理不可"].includes(status);
  if (sourceName === "運転免許管理" || sourceName === "資格管理" || row["種類"] === "運転免許" || row["種類"] === "資格") return !["有効", "更新済"].includes(status);
  if (sourceName === "車検管理" || row["種類"] === "車検") return status !== "更新済";
  if (sourceName === "社用車予約" || row["種類"] === "社用車予約") return !["返却済", "完了", "中止"].includes(status);
  if (sourceName === "個人ToDo" || row["種類"] === "ToDo") return !["完了", "中止"].includes(status);

  return !["完了", "確認済", "更新済", "返却済", "修理不可", "中止"].includes(status);
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ダッシュボード");
  if (!sheet) sheet = ss.insertSheet("ダッシュボード");

  if (sheet.getMaxColumns() < 4) sheet.insertColumnsAfter(sheet.getMaxColumns(), 4 - sheet.getMaxColumns());

  try { const filter = sheet.getFilter(); if (filter) filter.remove(); } catch (e) {}
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart(); } catch (e) {}

  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.setConditionalFormatRules([]);

  sheet.getRange("A1").setValue("社内共有ダッシュボード");
  sheet.getRange("A1:D1").setBackground("#d9ead3").setFontWeight("bold");
  sheet.getRange("A1").setFontSize(22).setFontWeight("bold");

  sheet.getRange("A3:C3").setValues([["項目", "件数", "備考"]]);
  sheet.getRange("A3:C3").setBackground("#d9ead3").setFontWeight("bold").setHorizontalAlignment("center");

  const data = [
    ["今日の予定", '=COUNTIF(一覧スケジュール!A2:A,TODAY())', ""],
    ["期限切れ", '=COUNTIF(一覧スケジュール!F2:F,"期限切れ")', ""],
    ["今日対応", '=COUNTIF(一覧スケジュール!F2:F,"今日")', ""],
    ["3日以内", '=COUNTIF(一覧スケジュール!F2:F,"3日以内")', ""],
    ["未読件数", '=COUNTIF(一覧スケジュール!H2:H,FALSE)', ""],
    ["未対応電話", '=COUNTIF(一覧スケジュール!G2:G,"未対応")', ""],
    ["折返し電話", '=COUNTIF(一覧スケジュール!G2:G,"折返し")', ""],
    ["高重要度お知らせ", '=COUNTIF(お知らせ!E2:E,"高")', ""],
    ["未完了ToDo", '=COUNTIFS(個人ToDo!C2:C,"<>",個人ToDo!E2:E,"<>完了")', ""],
    ["日報件数", '=COUNTA(日報!A2:A)', ""],
    ["日報レシート件数", '=COUNTA(日報レシート管理!A2:A)', ""],
    ["運転免許件数", '=COUNTA(運転免許管理!A2:A)', ""],
    ["資格件数", '=COUNTA(資格管理!A2:A)', ""],
    ["備品修理件数", '=COUNTA(備品修理管理!A2:A)', ""],
    ["フィードバック未確認", '=COUNTIF(フィードバック!H2:H,"未確認")', ""]
  ];

  sheet.getRange(4, 1, data.length, 3).setValues(data);
  sheet.getRange(3, 1, data.length + 1, 3).setBorder(true, true, true, true, true, true);
  sheet.getRange(4, 2, data.length, 1).setHorizontalAlignment("center");
  sheet.setFrozenRows(3);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 260);
}

function setupInitialLedgerOutputSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("帳簿出力設定");
  if (!sheet) sheet = ss.insertSheet("帳簿出力設定");
  setupSheetWithoutClearing_(sheet, getSheetHeaders_()["帳簿出力設定"]);

  if (sheet.getLastRow() >= 2) return;

  writeObjectsToSheetNoCheckboxAllRows_(sheet, [
    {"出力する": true, "帳簿名": "一覧スケジュール", "シート名": "一覧スケジュール", "出力列": "日付,種類,内容,担当,状態,通知,備考", "最大行数": 200, "備考": "通常確認用"},
    {"出力する": true, "帳簿名": "要確認一覧", "シート名": "要確認一覧", "出力列": "日付,種類,内容,担当,状態,通知,備考", "最大行数": 200, "備考": "要対応確認用"},
    {"出力する": false, "帳簿名": "フィードバック", "シート名": "フィードバック", "出力列": "日付,記入者,対象シート,気になった内容,区分,困り度,対応状況", "最大行数": 200, "備考": "ベータ確認用"}
  ]);
}

function addBetaSampleDataAllSheetsSafe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const staff = getStaffMembers_();
  const a = staff[0] || "山田";
  const b = staff[1] || "鈴木";
  const c = staff[2] || "田中";

  const sampleMap = {
    "出先予定": [
      {"日付": today, "行き先": "中央町サンプル現場", "用件": "ベータサンプル 現地確認", "担当": a, "社用車": "4t", "状態": "予定", "通知": getNoticeText(today), "電話対応": "未対応", "備考": "ベータサンプル"}
    ],
    "工事予定": [
      {"工事名": "ベータサンプル 舗装修繕工事", "現場": "東町サンプル現場", "依頼主": "サンプル建設株式会社", "連絡先": "000-0000-0000", "契約金額": 1200000, "開始日": today, "終了日": addDays_(today, 7), "状態": "着工前", "担当": b, "通知": getNoticeText(today), "電話対応": "未対応", "備考": "ベータサンプル"}
    ],
    "会議予定": [
      {"日付": addDays_(today, 1), "会議名": "ベータサンプル 工程会議", "内容": "工程確認", "担当": a, "状態": "予定", "通知": getNoticeText(addDays_(today, 1)), "備考": "ベータサンプル"}
    ],
    "行事予定": [
      {"開始日": addDays_(today, 2), "終了日": addDays_(today, 2), "開始時刻": "09:00", "終了時刻": "10:00", "行事名": "ベータサンプル 安全大会", "内容": "安全確認", "担当": c, "状態": "予定", "通知": getNoticeText(addDays_(today, 2)), "備考": "ベータサンプル"}
    ],
    "電話履歴": [
      {"日付": today, "時間": "10:30", "相手": "サンプル取引先", "内容": "見積確認の折返し依頼", "電話対応": "折返し", "担当": a, "対応メモ": "午後に折返し", "通知": getNoticeText(today), "備考": "ベータサンプル"}
    ],
    "車検管理": [
      {"車両名": "サンプル車両A", "車番": "00-00", "車検期限": addDays_(today, 3), "車検予定日": "", "通知": getNoticeText(addDays_(today, 3)), "保険期限": addDays_(today, 30), "車検状態": "未対応", "備考": "ベータサンプル"},
      {"車両名": "サンプル車両B", "車番": "11-11", "車検期限": addDays_(today, -1), "車検予定日": addDays_(today, 5), "通知": getNoticeText(addDays_(today, -1)), "保険期限": addDays_(today, 40), "車検状態": "予約済", "備考": "ベータサンプル"}
    ],
    "社用車予約": [
      {"日付": addDays_(today, 1), "開始時刻": "08:30", "終了時刻": "12:00", "社用車": "軽トラ②", "利用者": b, "行き先": "西町サンプル現場", "用途": "資材運搬", "状態": "予定", "通知": getNoticeText(addDays_(today, 1)), "備考": "ベータサンプル"}
    ],
    "日報": [
      {"日付": today, "入力者": a, "担当": a, "現場": "中央町サンプル現場", "作業内容": "ベータサンプル 掘削作業", "進捗": "予定通り", "問題点": "特になし", "明日の予定": "砕石敷均し", "状態": "下書き", "備考": "ベータサンプル"}
    ],
    "日報レシート管理": [
      {"日付": today, "担当": a, "現場": "中央町サンプル現場", "支払先": "サンプル給油所", "内容": "軽油代", "区分": "燃料", "金額": 5200, "支払方法": "立替", "経理確認": "未確認", "精算状態": "未精算", "通知": "未精算", "備考": "ベータサンプル"}
    ],
    "運転免許管理": [
      {"所有者": a, "普通": true, "準中型": true, "取得日": addDays_(today, -1000), "更新期限": addDays_(today, 60), "コピー有無": "有", "状態": "有効", "通知": getNoticeText(addDays_(today, 60)), "備考": "ベータサンプル"}
    ],
    "資格管理": [
      {"所有者": b, "資格名": "サンプル資格", "区分": "技能講習", "取得日": addDays_(today, -500), "更新期限": addDays_(today, 20), "コピー有無": "未確認", "状態": "更新予定", "通知": getNoticeText(addDays_(today, 20)), "備考": "ベータサンプル"}
    ],
    "備品修理管理": [
      {"購入日": addDays_(today, -300), "備品名": "サンプル発電機", "修理業者": "サンプル修理店", "内容": "始動不良", "修理依頼者": c, "修理依頼日": today, "返却予定日": addDays_(today, 5), "返却済": "未", "状態": "依頼済", "通知": getNoticeText(addDays_(today, 5)), "備考": "ベータサンプル"}
    ],
    "お知らせ": [
      {"投稿日": today, "タイトル": "ベータサンプル お知らせ", "内容": "システム確認用のお知らせです。", "投稿者": a, "重要度": "中", "通知": "今日", "備考": "ベータサンプル"}
    ],
    "個人ToDo": [
      {"登録日": today, "担当": b, "内容": "ベータサンプル 見積確認", "期限": addDays_(today, 1), "状態": "未着手", "通知": getNoticeText(addDays_(today, 1)), "備考": "ベータサンプル"}
    ],
    "フィードバック": [
      {"日付": today, "記入者": a, "確認方法": "PC", "対象シート": "工事予定", "気になった内容": "ベータサンプル：依頼主・連絡先・契約金額が必要か確認", "区分": "運用確認", "困り度": "メモ", "対応状況": "未確認", "対応メモ": "不要なら削除"}
    ]
  };

  Object.keys(sampleMap).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    if (sheetContainsText_(sheet, "ベータサンプル")) return;
    writeObjectsToSheetNoCheckboxAllRows_(sheet, sampleMap[sheetName]);
  });

  refreshAll();
  formatInputSheetsForLongText();
  SpreadsheetApp.getActiveSpreadsheet().toast("ベータ用サンプルを各シートへ追加しました");
}

function sheetContainsText_(sheet, text) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return false;
  const finder = sheet.createTextFinder(text).matchCase(false);
  return !!finder.findNext();
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function formatSheetByName_(sheet, name) {
  if (!sheet) return;
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];

  const widthMap = {
    "工事予定": {"工事名": 180, "現場": 160, "依頼主": 160, "連絡先": 130, "契約金額": 110, "開始日": 95, "終了日": 95, "状態": 95, "担当": 110, "通知": 90, "電話対応": 100, "備考": 220},
    "電話履歴": {"日付": 95, "時間": 70, "相手": 160, "内容": 260, "電話対応": 100, "担当": 110, "対応メモ": 240, "通知": 90, "備考": 220},
    "車検管理": {"車両名": 150, "車番": 95, "車検期限": 105, "車検予定日": 105, "通知": 90, "保険期限": 105, "車検状態": 100, "写真": 120, "備考": 220},
    "運転免許管理": {"所有者": 130, "取得日": 95, "更新期限": 105, "コピー有無": 100, "状態": 90, "通知": 90, "備考": 220},
    "資格管理": {"所有者": 130, "資格名": 180, "区分": 110, "取得日": 95, "更新期限": 105, "コピー有無": 100, "状態": 90, "通知": 90, "備考": 220},
    "備品修理管理": {"購入日": 95, "備品名": 170, "修理業者": 170, "内容": 240, "修理依頼者": 120, "修理依頼日": 105, "返却予定日": 105, "返却済": 80, "状態": 100, "通知": 90, "備考": 220},
    "フィードバック": {"日付": 95, "記入者": 110, "確認方法": 130, "対象シート": 140, "気になった内容": 420, "区分": 120, "困り度": 80, "対応状況": 100, "対応方針": 260, "対応メモ": 260, "対応日": 95}
  };

  headers.forEach((header, idx) => {
    const col = idx + 1;
    let width = 100;
    if (widthMap[name] && widthMap[name][header]) width = widthMap[name][header];
    else if (getPersonalMembers_().includes(header)) width = 42;
    else if (["普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種"].includes(header)) width = 70;
    else if (header === READ_HEADER) width = 70;
    else if (header.indexOf("ID") >= 0 || header === CALENDAR_ID_HEADER) width = 120;
    try { sheet.setColumnWidth(col, width); } catch (e) {}
  });

  try {
    sheet.getRange(1, 1, 1, headers.length).setWrap(true).setVerticalAlignment("middle");
    sheet.setRowHeight(1, ["運転免許管理", "資格管理"].includes(name) ? 46 : 42);
    const lastRow = Math.max(sheet.getLastRow(), 2);
    sheet.getRange(2, 1, lastRow - 1, headers.length).setWrap(true).setVerticalAlignment("middle");
  } catch (e) {}
}

function formatInputSheetsForLongText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    formatSheetByName_(sheet, name);
    applyColorRules(sheet);
    hideTechnicalColumns_(sheet);
  });

  formatPersonalCheckColumnsVerticalAllSheets();
  SpreadsheetApp.getActiveSpreadsheet().toast("表示幅・カラーリングを調整しました");
}

function hideTechnicalColumns_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  headers.forEach((header, idx) => {
    if (!header) return;
    const col = idx + 1;
    if (header === CALENDAR_ID_HEADER || /ID$/.test(header) || header === "履歴ID" || header === "ToDo_ID") {
      try { sheet.hideColumns(col); } catch (e) {}
    }
  });
}

function applyColorRules(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const rules = [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = Math.max(sheet.getMaxRows(), 2);
  const sheetName = sheet.getName();

  function addRule(header, text, bg) {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;
    const range = sheet.getRange(2, col, Math.min(lastRow - 1, 300), 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(text)
      .setBackground(bg)
      .setRanges([range])
      .build());
  }

  ["状態", "車検状態", "電話対応", "通知", "経理確認", "精算状態", "困り度", "対応状況"].forEach(header => {
    addRule(header, "期限切れ", "#ea9999");
    addRule(header, "今日", "#f9cb9c");
    addRule(header, "3日以内", "#fff2cc");
    addRule(header, "未対応", "#f4cccc");
    addRule(header, "未確認", "#f4cccc");
    addRule(header, "差戻し", "#e6b8af");
    addRule(header, "未精算", "#f9cb9c");
    addRule(header, "対応中", "#fff2cc");
    addRule(header, "折返し", "#fce5cd");
    addRule(header, "施工中", "#fff2cc");
    addRule(header, "修理中", "#fff2cc");
    addRule(header, "返却済", "#d9ead3");
    addRule(header, "確認済", "#d9ead3");
    addRule(header, "更新済", "#d9ead3");
    addRule(header, "完了", "#d9ead3");
    addRule(header, "修理不可", "#d9d2e9");
    addRule(header, "高", "#f4cccc");
  });

  sheet.setConditionalFormatRules(rules);
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の更新処理中のため、全体更新をスキップしました");
    return;
  }

  try {
    refreshInputSheetsLightV9630_();
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    SpreadsheetApp.getActiveSpreadsheet().toast("全体更新（軽量・高速）が完了しました");
  } finally {
    lock.releaseLock();
  }
}

function refreshInputSheetsLightV9630_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getPrimaryInputSheetNamesForSetup_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    ensureIdsForSheet_(sheet);
    refreshNoticeColumnForSheetV9630_(sheet);
  });
}

function refreshNoticeColumnForSheetV9630_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const noticeCol = headers.indexOf("通知") + 1;
  if (noticeCol <= 0) return;

  let dateHeader = "";
  const name = sheet.getName();
  if (name === "工事予定") dateHeader = "開始日";
  else if (name === "電話履歴") dateHeader = "日付";
  else if (name === "車検管理") dateHeader = "車検期限";
  else if (name === "備品修理管理") dateHeader = "返却予定日";
  else if (name === "運転免許管理" || name === "資格管理") dateHeader = "更新期限";
  else if (name === "個人ToDo") dateHeader = "期限";
  else if (name === "行事予定") dateHeader = "開始日";
  else if (headers.indexOf("日付") >= 0) dateHeader = "日付";

  const dateCol = headers.indexOf(dateHeader) + 1;
  if (dateCol <= 0) return;

  const statusCol = headers.indexOf("状態") + 1;
  const carStatusCol = headers.indexOf("車検状態") + 1;
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const noticeValues = values.map(row => {
    const status = statusCol > 0 ? String(row[statusCol - 1] || "") : "";
    const carStatus = carStatusCol > 0 ? String(row[carStatusCol - 1] || "") : "";

    if (["完了", "確認済", "更新済", "返却済", "修理不可", "中止"].includes(status)) return [""];
    if (carStatus === "更新済") return [""];

    return [getNoticeText(row[dateCol - 1])];
  });

  sheet.getRange(2, noticeCol, noticeValues.length, 1).setValues(noticeValues);
}

function onEdit(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const row = e.range.getRow();
  if (!sheet || row === 1) return;

  try { ensureRowsAfterEdit_(sheet, row); } catch (err) {}
  try { ensureIdForEditedRow_(sheet, row); } catch (err) {}
  try { setCheckboxesForEditedRowSmart_(sheet, row); } catch (err) {}
  try { updateNoticeForEditedRow_(sheet, row); } catch (err) {}

  const sheetName = sheet.getName();
  const autoRefreshTargets = getPrimaryInputSheetNamesForSetup_();

  if (autoRefreshTargets.includes(sheetName)) {
    try {
      createScheduleList();
      createAlertList();
      createDashboard();
    } catch (err) {
      console.log("onEdit軽量更新をスキップ: " + err.message);
    }
  }
}

function initialSetupForProduction() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、初回セットアップをスキップしました");
    return;
  }

  try {
    clearSettingsCache_();
    ensureSettingsSheet_();
    createCompanySheets();
    createFeedbackSheet();
    createSqlSupportSheetsVisible();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    setupInitialLedgerOutputSettings_();
    applySettingsToSheetsByNames_(getPrimaryInputSheetNamesForSetup_(), "入力シートを整備しました");
    reorderSheetsForOperation();
    formatInputSheetsForLongText();
    refreshAll();
    SpreadsheetApp.getActiveSpreadsheet().toast("初回セットアップが完了しました");
  } finally {
    lock.releaseLock();
  }
}

function feedbackPreparationForBetaV9630() {
  createFeedbackSheet();
  addBetaSampleDataAllSheetsSafe();
  createSqlSupportSheetsVisible();
  refreshAll();
  reorderSheetsForOperation();
  formatInputSheetsForLongText();
  SpreadsheetApp.getActiveSpreadsheet().toast("ベータ確認準備が完了しました");
}

function createCompanySheets() {
  clearSettingsCache_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet_();

  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    let sheet = ss.getSheetByName(name);
    const headers = headersMap[name];
    if (!sheet) sheet = ss.insertSheet(name);
    setupHeaderOnly_(sheet, headers);
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast("シート構成を作成しました");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("社内管理")
    .addSubMenu(
      ui.createMenu("1. 初回セットアップ")
        .addItem("新規ベータ作成（前シート前提なし）", "createCompanySheetsNewBlankNoLegacy")
        .addItem("初回セットアップ（軽量・高速）", "initialSetupForProduction")
        .addItem("初回セットアップ仕上げ（表示・順番・既読）", "refreshAllWithDisplayMaintenance")
        .addItem("フィードバックシートを作成/整備", "createFeedbackSheet")
        .addItem("SQL関連シートを作成/表示", "createSqlSupportSheetsVisible")
    )
    .addSubMenu(
      ui.createMenu("2. 日常運用")
        .addItem("全体更新（軽量・高速）", "refreshAll")
        .addItem("表示調整付き更新", "refreshAllWithDisplayMaintenance")
        .addItem("Googleカレンダーへ反映", "syncCalendarEvents")
        .addItem("車検更新完了処理", "completeVehicleInspectionUpdates")
        .addItem("一覧帳簿PDFを作成", "createSelectedLedgerPdfs")
        .addItem("選択行の日報文章を作成", "createDailyReportTextForSelectedRow")
        .addItem("選択行の日報PDFを作成", "createDailyReportPdfForSelectedRow")
        .addItem("選択行の日報文章＋PDFを作成", "createDailyReportTextAndPdfForSelectedRow")
    )
    .addSubMenu(
      ui.createMenu("3. 保守・テスト")
        .addItem("ベータ確認準備（表示＋各シートサンプル＋FBシート）", "feedbackPreparationForBetaV9630")
        .addItem("ベータ用サンプルを各シートへ安全追加", "addBetaSampleDataAllSheetsSafe")
        .addItem("フィードバック記入例を追加", "addFeedbackExampleRows")
        .addItem("フィードバックシートを作成/整備", "createFeedbackSheet")
        .addItem("SQL関連シートを作成/表示", "createSqlSupportSheetsVisible")
        .addSeparator()
        .addItem("今のシートの空欄チェックボックス整理", "cleanupBlankRowCheckboxesCurrentSheet")
        .addItem("データ整合性チェック・軽微修正", "runDataConsistencyCheckAndFix")
        .addItem("表示幅・カラーリングを調整", "formatInputSheetsForLongText")
        .addItem("シート順を整える", "reorderSheetsForOperation")
        .addItem("個人確認列を縦書きにする", "formatPersonalCheckColumnsVerticalAllSheets")
        .addItem("個人確認グループを作り直す", "rebuildCheckGroups")
        .addSeparator()
        .addItem("裏方シートを非表示", "hideSupportSheets")
        .addItem("スプレッドシートをバックアップ", "backupSpreadsheet")
        .addItem("空欄に戻す表示をクリア", "clearClearLabelValuesInAllSheets")
    )
    .addToUi();
}


/**
 * v9.6.31 表示幅・行高修正版
 * - v9.6.30で個人確認列の縦書き処理がヘッダー行を120pxに戻していたため修正。
 * - 入力シートはヘッダー46px前後、本文32〜44pxに統一。
 * - 出先予定など通常入力シートの列幅も明示指定。
 * - 表示幅・カラーリング調整を押した時に必ずこのルールが最後に効くように上書き。
 */

function getComfortLayoutConfigV9631_(sheetName) {
  const compact = { header: 46, body: 32 };
  const normal = { header: 46, body: 36 };
  const longText = { header: 48, body: 44 };

  if (["日報", "フィードバック"].includes(sheetName)) return { header: 50, body: 56 };
  if (["備品修理管理", "日報レシート管理"].includes(sheetName)) return longText;
  if (["一覧スケジュール", "要確認一覧", "過去一覧"].includes(sheetName)) return normal;
  if (["社内管理説明", "手順シート", "SQL設計", "移行対応表", "SQLサンプル集"].includes(sheetName)) return longText;
  return compact;
}

function getColumnWidthMapV9631_(sheetName) {
  const common = {
    "日付": 96, "日時": 120, "時間": 72, "開始時刻": 80, "終了時刻": 80,
    "開始日": 96, "終了日": 96, "期限": 96, "通知": 86, "状態": 92,
    "車検状態": 100, "担当": 112, "入力者": 112, "所有者": 126,
    "既読": 66, "備考": 220, "電話対応": 102,
    "カレンダーID": 120
  };

  const maps = {
    "出先予定": {"行き先": 170, "用件": 230, "社用車": 112},
    "工事予定": {"工事名": 190, "現場": 165, "依頼主": 165, "連絡先": 132, "契約金額": 112},
    "会議予定": {"会議名": 180, "内容": 260, "資料": 160},
    "行事予定": {"行事名": 190, "内容": 260},
    "作業状況": {"現場": 170, "作業内容": 260, "写真": 120},
    "電話履歴": {"相手": 165, "内容": 280, "対応メモ": 240},
    "車検管理": {"車両名": 150, "車番": 95, "車検期限": 105, "車検予定日": 105, "保険期限": 105, "写真": 120},
    "社用車予約": {"社用車": 112, "利用者": 112, "行き先": 170, "用途": 220},
    "日報": {"現場": 170, "作業内容": 260, "進捗": 150, "問題点": 220, "明日の予定": 220, "他現場状況": 220, "写真": 120, "日報文章": 360, "PDFリンク": 180},
    "日報レシート管理": {"現場": 155, "支払先": 170, "内容": 235, "区分": 96, "金額": 88, "支払方法": 102, "レシート写真": 145, "経理確認": 100, "精算状態": 100},
    "運転免許管理": {"普通": 64, "準中型": 70, "中型": 64, "大型": 64, "大型特殊": 78, "けん引": 70, "二種": 64, "取得日": 96, "更新期限": 105, "コピー有無": 100},
    "資格管理": {"資格名": 190, "区分": 120, "取得日": 96, "更新期限": 105, "コピー有無": 100},
    "備品修理管理": {"購入日": 96, "備品名": 170, "修理業者": 175, "内容": 250, "修理依頼者": 120, "修理依頼日": 105, "返却予定日": 105, "返却済": 78},
    "お知らせ": {"投稿日": 96, "タイトル": 220, "内容": 360, "投稿者": 112, "重要度": 86},
    "個人ToDo": {"登録日": 96, "内容": 300, "期限": 96},
    "一覧スケジュール": {"種類": 110, "内容": 330, "元シート": 130},
    "要確認一覧": {"種類": 110, "内容": 360, "元シート": 130},
    "フィードバック": {"記入者": 112, "確認方法": 130, "対象シート": 140, "気になった内容": 430, "区分": 126, "困り度": 82, "対応状況": 105, "対応方針": 260, "対応メモ": 280, "対応日": 96}
  };

  return Object.assign({}, common, maps[sheetName] || {});
}

function formatSheetByName_(sheet, name) {
  if (!sheet) return;
  name = name || sheet.getName();
  if (sheet.getLastColumn() < 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const widthMap = getColumnWidthMapV9631_(name);
  const personalMembers = getPersonalMembers_();
  const licenseHeaders = ["普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種"];

  headers.forEach((header, idx) => {
    if (!header) return;
    const col = idx + 1;
    let width = widthMap[header] || 104;

    if (personalMembers.includes(header)) width = 42;
    if (licenseHeaders.includes(header)) width = widthMap[header] || 66;
    if (/ID$/.test(header) || header === "ToDo_ID" || header === CALENDAR_ID_HEADER) width = 120;

    try { sheet.setColumnWidth(col, width); } catch (e) {}
  });

  const config = getComfortLayoutConfigV9631_(name);
  try {
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, config.header);
    sheet.getRange(1, 1, 1, headers.length)
      .setWrap(true)
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setFontWeight("bold")
      .setBackground("#d9ead3");

    const rowCount = Math.min(sheet.getMaxRows() - 1, Math.max(sheet.getLastRow() + 20, 60) - 1);
    if (rowCount > 0) {
      sheet.setRowHeights(2, rowCount, config.body);
      sheet.getRange(2, 1, rowCount, headers.length)
        .setVerticalAlignment("middle")
        .setWrap(["日報", "フィードバック", "備品修理管理", "日報レシート管理"].includes(name));
    }

    const moneyCol = headers.indexOf("契約金額") + 1;
    if (moneyCol > 0) sheet.getRange(2, moneyCol, Math.max(rowCount, 1), 1).setNumberFormat("#,##0");
    const amountCol = headers.indexOf("金額") + 1;
    if (amountCol > 0) sheet.getRange(2, amountCol, Math.max(rowCount, 1), 1).setNumberFormat("#,##0");
  } catch (e) {
    console.log(name + " の表示調整をスキップ: " + e.message);
  }
}

function formatPersonalCheckColumnsVertical_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const personalMembers = getPersonalMembers_();
  const sheetName = sheet.getName();
  const config = getComfortLayoutConfigV9631_(sheetName);

  try {
    // 旧版はここで120pxにしていた。v9.6.31では広げすぎない。
    sheet.setRowHeight(1, config.header);
  } catch (e) {}

  const readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol > 0) {
    try {
      sheet.setColumnWidth(readCol, 66);
      sheet.getRange(1, readCol).setVerticalText(false).setTextRotation(0);
      sheet.getRange(1, readCol, Math.min(sheet.getMaxRows(), 80), 1)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    } catch (e) {}
  }

  personalMembers.forEach(name => {
    const col = headers.indexOf(name) + 1;
    if (col <= 0) return;
    try {
      sheet.setColumnWidth(col, 42);
      sheet.getRange(1, col).setVerticalText(true).setWrap(true).setHorizontalAlignment("center").setVerticalAlignment("middle");
      sheet.getRange(2, col, Math.min(sheet.getMaxRows() - 1, 79), 1)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    } catch (e) {}
  });
}

function formatPersonalCheckColumnsVerticalAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getPersonalCheckVerticalTargetSheets_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    formatPersonalCheckColumnsVertical_(sheet);
    hideTechnicalColumns_(sheet);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("個人確認列を縦書きにしました（行高控えめ）");
}

function formatInputSheetsForLongText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headersMap = getSheetHeaders_();

  Object.keys(headersMap).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    formatSheetByName_(sheet, name);
    applyColorRules(sheet);
    hideTechnicalColumns_(sheet);
    formatPersonalCheckColumnsVertical_(sheet);
    // 縦書き処理後にもう一度行高だけ戻す。
    try { sheet.setRowHeight(1, getComfortLayoutConfigV9631_(name).header); } catch (e) {}
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("表示幅・行高・カラーリングを調整しました（v9.6.31）");
}

function refreshAllWithDisplayMaintenance() {
  refreshAll();
  reorderSheetsForOperation();
  formatInputSheetsForLongText();
  SpreadsheetApp.getActiveSpreadsheet().toast("表示調整付き更新が完了しました（v9.6.31）");
}

function createCompanySheetsNewBlankNoLegacy() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getActiveSpreadsheet().toast("他の処理中のため、新規作成をスキップしました");
    return;
  }

  try {
    clearSettingsCache_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const response = ui.alert(
      "新規ベータ作成",
      "現在のスプレッドシート内の既存シートを初期化して、最新版構成で作り直します。空の新規スプレッドシート用です。実行しますか？",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ss.toast("新規ベータ作成をキャンセルしました");
      return;
    }

    let temp = ss.getSheetByName("_TEMP_SETUP_");
    if (!temp) temp = ss.insertSheet("_TEMP_SETUP_");

    ss.getSheets().forEach(sheet => {
      if (sheet.getName() !== "_TEMP_SETUP_") ss.deleteSheet(sheet);
    });

    ensureSettingsSheet_();

    const headersMap = getSheetHeaders_();
    getOperationalSheetOrder_().forEach(name => {
      if (!headersMap[name]) return;
      const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
      setupHeaderOnly_(sheet, headersMap[name]);
    });

    Object.keys(headersMap).forEach(name => {
      if (ss.getSheetByName(name)) return;
      const sheet = ss.insertSheet(name);
      setupHeaderOnly_(sheet, headersMap[name]);
    });

    const tempSheet = ss.getSheetByName("_TEMP_SETUP_");
    if (tempSheet && ss.getSheets().length > 1) ss.deleteSheet(tempSheet);

    createFeedbackSheet();
    createSqlSupportSheetsVisible();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    setupInitialLedgerOutputSettings_();

    applySettingsToSheetsByNames_(getPrimaryInputSheetNamesForSetup_(), "入力シートを整備しました");
    refreshAll();
    reorderSheetsForOperation();
    formatInputSheetsForLongText();

    ss.toast("新規ベータ作成が完了しました（表示幅・行高調整済み）");
  } finally {
    lock.releaseLock();
  }
}






