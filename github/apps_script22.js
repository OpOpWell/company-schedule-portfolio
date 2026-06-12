
/**
 * 社内共有業務管理システム v9.7.2 軽量更新分離版
 *
 * 方針:
 * - v9.6系で増えた重複定義・旧列名混在を整理した新規ベータ向けクリーン版。
 * - 新規ベータ作成は「このコードの列構成で最初から作る」前提。
 * - 既存シートへ使う場合は、必ずコピー/バックアップ後に「旧列名をv9.7.0へ移行」を実行してください。
 * - 日常運用では「日常更新（一覧＋要確認）」を使い、重い全体更新は保守時だけ使う想定。
 *
 * 主な修正:
 * - 車番列を文字列固定。サンプル車番も日付化しない値に変更。
 * - 車検管理は「車検期限 / 車検予定日 / 車検状態」に統一。
 * - 旧「免許資格管理」は新規構成では使わず、「運転免許管理」「資格管理」に分離。
 * - 備品修理管理は「購入日 / 修理業者 / 修理依頼者 / 修理不可」に統一。
 * - 時間列は hh:mm 表示。
 * - onEditは軽量化し、一覧・要確認の再作成は手動更新で行う。
 * - 月別グループメニューを復帰。
 * - 重い表示調整・条件付き書式は初期/保守時中心に実行。
 *
 * v9.7.2:
 * - 日常更新・集計更新・保守用全体更新を分離。
 * - ベータ確認準備中のサンプル追加で重複する全体再計算を抑制。
 * - 全体更新から毎回の時間書式再設定を外し、必要時だけ実行する設計に変更。
 */

const SYSTEM_VERSION = "v9.7.2";
const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";
const CALENDAR_ID_HEADER = "カレンダーID";
const SUMMARY_DIRTY_KEY = "SUMMARY_DIRTY";
let IS_BATCH_SETUP_ = false;

const DEFAULT_STAFF_MEMBERS = [
  "山田", "鈴木", "田中", "高橋", "伊藤", "中村", "小林", "吉田", "山本", "佐々木", "松本", "井上", "木村", "林", "清水"
];

const DEFAULT_PERSONAL_MEMBERS = ["山田", "高橋", "鈴木"];

const MIN_DATA_ROWS_FOR_VALIDATION = 300;
const AUTO_EXTEND_ROWS_BUFFER = 100;
const AUTO_EXTEND_TRIGGER_MARGIN = 30;

const DATE_HEADERS = [
  "日付", "投稿日", "登録日", "購入日", "開始日", "終了日", "車検期限", "車検予定日", "取得日", "更新期限", "修理依頼日", "返却予定日", "作成日時", "対応日", "移動日", "期限"
];

const TIME_HEADERS = ["時間", "開始時刻", "終了時刻"];
const VEHICLE_NUMBER_HEADERS = ["車番"];

const LICENSE_TYPE_HEADERS = ["普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種"];

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
  "日報レシート管理": "レシートID",
  "運転免許管理": "免許ID",
  "資格管理": "資格ID",
  "備品修理管理": "修理ID",
  "お知らせ": "お知らせID",
  "個人ToDo": "ToDo_ID",
  "フィードバック": "フィードバックID",
  "帳簿PDF履歴": "帳簿ID"
};

const SHEET_ORDER = [
  "機能チェック",
  "一覧スケジュール",
  "要確認一覧",
  "出先予定",
  "工事予定",
  "電話履歴",
  "車検管理",
  "社用車予約",
  "運転免許管理",
  "資格管理",
  "備品修理管理",
  "日報",
  "日報レシート管理",
  "会議予定",
  "行事予定",
  "作業状況",
  "お知らせ",
  "個人ToDo",
  "フィードバック",
  "過去一覧",
  "車検履歴",
  "担当別未読",
  "既読率集計",
  "ダッシュボード",
  "帳簿PDF履歴",
  "帳簿出力設定",
  "SQL設計",
  "移行対応表",
  "SQLサンプル集",
  "手順シート",
  "社内管理説明",
  "設定"
];

const SUPPORT_SHEETS = [
  "設定", "SQL設計", "移行対応表", "SQLサンプル集", "帳簿PDF履歴", "帳簿出力設定", "担当別未読", "既読率集計", "ダッシュボード", "手順シート", "社内管理説明", "機能チェック"
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("社内管理")
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("1. 初回セットアップ")
        .addItem("新規ベータ作成（前シート前提なし）", "setupNewBetaClean")
        .addItem("ベータ確認準備（軽量）", "prepareBetaCheckLight")
        .addItem("手順シートを作成/更新", "createProcedureSheet")
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("2. 日常運用")
        .addItem("日常更新（一覧＋要確認）", "refreshDailyOnly")
        .addItem("集計更新（未読・既読率・DB）", "refreshSummaryOnly")
        .addSeparator()
        .addItem("全体更新（保守用・重め）", "refreshAll")
        .addSeparator()
        .addItem("一覧スケジュールを更新", "createScheduleList")
        .addItem("要確認一覧を更新", "createAlertList")
        .addItem("今のシートだけ表示調整", "formatActiveSheetOnly")
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("3. 保守・テスト")
        .addItem("表示幅・カラーリングを調整", "formatInputSheetsForLongText")
        .addItem("書式込み全体更新（重い）", "refreshAllWithFormat")
        .addItem("旧列名をv9.7.0へ移行", "migrateLegacyColumnsV970")
        .addItem("データ整合性チェック・軽微修正", "runDataConsistencyCheckAndFix")
        .addSeparator()
        .addItem("サンプル追加：主要予定系", "addSampleMainSchedules")
        .addItem("サンプル追加：車検・免許・資格・備品", "addSampleVehicleLicenseEquipment")
        .addItem("サンプル追加：日報・レシート・FB", "addSampleDailyReceiptFeedback")
        .addItem("ベータ用サンプルを削除", "deleteBetaSamples")
        .addSeparator()
        .addItem("機能チェックシートを作成", "createFeatureCheckSheet")
        .addItem("裏方シートを作成/修復", "ensureSupportSheets")
        .addItem("車検管理を修復（車番文字列＋色）", "repairVehicleInspectionSheet")
        .addItem("裏方シートを非表示", "hideSupportSheets")
        .addItem("裏方シートを表示", "showSupportSheets")
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("4. 月次・整理")
        .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
        .addItem("一覧スケジュールを月別グループ化", "groupScheduleByMonth")
        .addItem("過去一覧を月別グループ化", "groupArchiveByMonth")
        .addItem("今のシートを月別グループ化", "groupActiveSheetByMonth")
        .addItem("月次整理（過去移動＋過去一覧月別）", "monthlyMaintenance")
    )
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("5. PDF・帳簿")
        .addItem("選択行の日報PDFを作成", "createDailyReportPdfForActiveRow")
        .addItem("一覧帳簿PDFを作成", "createLedgerPdf")
    )
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  if (row <= 1) return;
  const sheetName = sheet.getName();
  if (isAutoOutputSheet_(sheetName) || SUPPORT_SHEETS.includes(sheetName)) return;

  try {
    ensureRowsAfterEdit_(sheet, row);
    ensureIdForEditedRow_(sheet, row);
    normalizeClearLabelForEditedCell_(e);
    updateNoticeForEditedRow_(sheet, row);
    setCheckboxesForEditedRow_(sheet, row);
    markSummaryDirty_();
  } catch (err) {
    console.log("onEdit error: " + err.message);
  }
}

function setupNewBetaClean() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "新規ベータ作成",
    "この処理は現在のシートをこのコードのv9.7.0構成で作り直します。既存データ入りの本体では実行しないでください。続行しますか？",
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    clearSettingsCache_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSettingsSheet_();
    const headersMap = getSheetHeaders_();

    Object.keys(headersMap).forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) sheet = ss.insertSheet(name);
      resetSheetForSetup_(sheet, headersMap[name].length);
      setupSheet_(sheet, headersMap[name]);
    });

    createSqlSheets();
    createProcedureSheet();
    createCompanyManagementExplanationSheet();
    createLedgerOutputSettings();
    reorderSheets_();
    formatInputSheetsForLongText();
    markSummaryDirty_();
    toast_("新規ベータを作成しました。次に『ベータ確認準備（軽量）』を実行してください。");
  } finally {
    lock.releaseLock();
  }
}

function prepareBetaCheckLight() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    IS_BATCH_SETUP_ = true;
    ensureSettingsSheet_();
    ensureSupportSheetsCore_();
    createFeedbackSheet();
    addSampleMainSchedules();
    addSampleVehicleLicenseEquipment();
    addSampleDailyReceiptFeedback();
    IS_BATCH_SETUP_ = false;
    refreshAllCore_({applyTimeFormat: false, ensureSupport: true});
    createFeatureCheckSheet();
    toast_("ベータ確認準備が完了しました。機能チェックシートを確認してください。");
  } finally {
    IS_BATCH_SETUP_ = false;
    lock.releaseLock();
  }
}

function refreshDailyOnly() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の更新処理中のため、日常更新をスキップしました");
    return;
  }

  try {
    createScheduleList();
    createAlertList();
    clearSummaryDirty_();
    toast_("日常更新が完了しました（一覧＋要確認）");
  } finally {
    lock.releaseLock();
  }
}

function refreshSummaryOnly() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の更新処理中のため、集計更新をスキップしました");
    return;
  }

  try {
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    toast_("集計更新が完了しました（未読・既読率・ダッシュボード）");
  } finally {
    lock.releaseLock();
  }
}

function refreshAll() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の更新処理中のため、全体更新をスキップしました");
    return;
  }

  try {
    refreshAllCore_({applyTimeFormat: false, ensureSupport: true});
    toast_("全体更新が完了しました（保守用）");
  } finally {
    lock.releaseLock();
  }
}

function refreshAllCore_(options) {
  const opt = options || {};
  if (opt.ensureSupport !== false) ensureSupportSheetsCore_();
  refreshInputSheets();
  createScheduleList();
  createAlertList();
  createAssigneeUnreadSummary();
  createReadRateSummary();
  createDashboard();
  if (opt.applyTimeFormat === true) applyTimeFormatsToAllSheets();
  clearSummaryDirty_();
}

function refreshAllWithFormat() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の更新処理中のため、書式込み全体更新をスキップしました");
    return;
  }

  try {
    refreshAllCore_({applyTimeFormat: true, ensureSupport: true});
    toast_("書式込み全体更新が完了しました");
  } finally {
    lock.releaseLock();
  }
}

function refreshInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    updateNoticeColumnForSheet_(sheet);
    ensureIdsForSheet_(sheet);
    setCheckboxesForDataRows(sheet);
  });
}

function getSheetHeaders_() {
  const personal = getPersonalMembers_();
  return {
    "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "出先ID"],
    "工事予定": ["工事名", "現場", "依頼主", "連絡先", "契約金額", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "工事ID"],
    "会議予定": ["日付", "開始時刻", "終了時刻", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "会議ID"],
    "行事予定": ["日付", "開始時刻", "終了時刻", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "行事ID"],
    "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...personal, "備考", "作業ID"],
    "電話履歴": ["日付", "時間", "相手", "内容", "電話対応", "担当", "対応メモ", "通知", READ_HEADER, ...personal, "備考", "電話ID"],
    "車検管理": ["車両名", "車番", "車検期限", "車検予定日", "通知", "保険期限", "車検状態", "写真", READ_HEADER, ...personal, "備考", "車両ID"],
    "車検履歴": ["更新日", "車両名", "車番", "旧車検期限", "新車検期限", "担当", "備考", "履歴ID"],
    "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "予約ID"],
    "日報": ["日付", "入力者", "担当", "現場", "作業内容", "進捗", "問題点", "明日の予定", "他現場状況", "写真", "日報文章", "状態", "備考", READ_HEADER, ...personal, "PDFリンク", "日報ID"],
    "日報レシート管理": ["日付", "担当", "現場", "支払先", "内容", "区分", "金額", "支払方法", "レシート写真", "経理確認", "精算状態", "通知", "備考", "レシートID"],
    "運転免許管理": ["所有者", ...LICENSE_TYPE_HEADERS, "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personal, "備考", "免許ID"],
    "資格管理": ["所有者", "資格名", "区分", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personal, "備考", "資格ID"],
    "備品修理管理": ["購入日", "備品名", "修理業者", "内容", "修理依頼者", "修理依頼日", "返却予定日", "返却済", "状態", "通知", READ_HEADER, ...personal, "備考", "修理ID"],
    "お知らせ": ["投稿日", "タイトル", "内容", "投稿者", "重要度", "通知", READ_HEADER, ...personal, "備考", "お知らせID"],
    "個人ToDo": ["登録日", "担当", "内容", "期限", "状態", "通知", READ_HEADER, ...personal, "備考", "ToDo_ID"],
    "フィードバック": ["日付", "記入者", "確認方法", "対象シート", "気になった内容", "区分", "困り度", "対応状況", "対応方針", "対応メモ", "対応日", "フィードバックID"],
    "一覧スケジュール": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", READ_HEADER, ...personal, "備考", "元シート"],
    "要確認一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "備考", "元シート"],
    "過去一覧": ["日付", "種類", "内容", "担当", "状態", "通知", "電話対応", "元シート", "備考", "移動日"],
    "担当別未読": ["担当", "未読件数", "未完了ToDo", "未対応電話", "今日の予定"],
    "既読率集計": ["対象", "全体件数", "既読件数", "未読件数", "既読率"],
    "帳簿PDF履歴": ["作成日時", "対象", "期間", "PDFリンク", "備考", "帳簿ID"],
    "帳簿出力設定": ["出力する", "帳簿名", "シート名", "出力列", "最大行数", "備考"],
    "ダッシュボード": ["項目", "件数", "備考"],
    "手順シート": ["区分", "順番", "実行メニュー", "使うタイミング", "注意"],
    "社内管理説明": ["項目", "説明"],
    "SQL設計": ["テーブル名", "日本語名", "カラム名", "データ型", "制約", "元シート", "元シート列", "備考"],
    "移行対応表": ["シート名", "列名", "SQLテーブル", "SQLカラム", "備考"],
    "SQLサンプル集": ["区分", "用途", "SQL", "説明"]
  };
}

function getInputSheetNames_() {
  return [
    "出先予定", "工事予定", "会議予定", "行事予定", "作業状況", "電話履歴", "車検管理", "社用車予約", "日報", "日報レシート管理", "運転免許管理", "資格管理", "備品修理管理", "お知らせ", "個人ToDo", "フィードバック"
  ];
}

function isAutoOutputSheet_(sheetName) {
  return ["一覧スケジュール", "要確認一覧", "担当別未読", "既読率集計", "ダッシュボード", "過去一覧", "帳簿PDF履歴", "帳簿出力設定", "SQL設計", "移行対応表", "SQLサンプル集", "手順シート", "社内管理説明", "機能チェック"].includes(sheetName);
}

function getStatusListForSheet_(sheetName) {
  const map = {
    "出先予定": [CLEAR_LABEL, "予定", "移動中", "完了", "延期", "中止"],
    "工事予定": [CLEAR_LABEL, "予定", "着工前", "施工中", "完了", "延期", "中止"],
    "会議予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "行事予定": [CLEAR_LABEL, "予定", "完了", "延期", "中止"],
    "作業状況": [CLEAR_LABEL, "着工前", "施工中", "完了", "要確認", "中止"],
    "電話履歴": [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"],
    "社用車予約": [CLEAR_LABEL, "予定", "使用中", "返却済", "予約重複", "中止"],
    "日報": [CLEAR_LABEL, "下書き", "提出済", "確認済", "差戻し", "完了"],
    "日報レシート管理": [CLEAR_LABEL, "未確認", "確認中", "確認済", "差戻し", "未精算", "精算済"],
    "運転免許管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "資格管理": [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"],
    "備品修理管理": [CLEAR_LABEL, "未依頼", "依頼済", "修理中", "修理不可", "返却済", "完了"],
    "車検管理": [CLEAR_LABEL, "未対応", "予約済", "実施済", "更新済"],
    "個人ToDo": [CLEAR_LABEL, "未着手", "対応中", "完了", "中止"],
    "お知らせ": [CLEAR_LABEL, "要確認", "確認済", "完了"],
    "フィードバック": [CLEAR_LABEL, "未対応", "対応中", "反映済", "保留", "対象外"]
  };
  return map[sheetName] || [CLEAR_LABEL, "予定", "対応中", "完了", "中止"];
}

function getDropdownListForHeader_(sheetName, header) {
  const staffList = [CLEAR_LABEL, ...getStaffMembers_()];
  const carList = [CLEAR_LABEL, "4t", "軽トラ②", "8tユニック", "軽ワゴン③", "ローダー", "P.BOX①", "P.BOX②", "3t", "キャブ③", "軽トラ③", "2t①", "軽ワゴン⑥", "軽バン⑤", "軽ダンプ②"];
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
  const importanceList = [CLEAR_LABEL, "高", "中", "低"];
  const copyList = [CLEAR_LABEL, "有", "無", "未確認"];
  const returnedList = [CLEAR_LABEL, "未", "済"];
  const receiptCategoryList = [CLEAR_LABEL, "燃料", "駐車場", "高速代", "消耗品", "資材", "工具", "修理", "その他"];
  const paymentList = [CLEAR_LABEL, "現金", "会社カード", "立替", "請求書", "その他"];
  const accountingList = [CLEAR_LABEL, "未確認", "確認中", "確認済", "差戻し"];
  const settlementList = [CLEAR_LABEL, "未精算", "精算中", "精算済", "対象外"];
  const feedbackCategoryList = [CLEAR_LABEL, "入力しづらい", "項目不足", "表示不備", "動作不備", "要望", "その他"];
  const difficultyList = [CLEAR_LABEL, "低", "中", "高", "至急"];
  const checkMethodList = [CLEAR_LABEL, "PC", "AppSheet", "紙", "口頭", "その他"];

  if (["担当", "入力者", "利用者", "投稿者", "記入者", "所有者", "修理依頼者"].includes(header)) return staffList;
  if (header === "社用車") return carList;
  if (header === "電話対応") return phoneList;
  if (header === "重要度") return importanceList;
  if (header === "コピー有無") return copyList;
  if (header === "返却済") return returnedList;
  if (header === "区分" && sheetName === "日報レシート管理") return receiptCategoryList;
  if (header === "支払方法") return paymentList;
  if (header === "経理確認") return accountingList;
  if (header === "精算状態") return settlementList;
  if (header === "区分" && sheetName === "フィードバック") return feedbackCategoryList;
  if (header === "困り度") return difficultyList;
  if (header === "確認方法") return checkMethodList;
  if (header === "対象シート") return [CLEAR_LABEL, ...getInputSheetNames_(), "一覧スケジュール", "要確認一覧", "AppSheet"];
  if (header === "対応状況") return getStatusListForSheet_(sheetName);
  if (header === "状態") return getStatusListForSheet_(sheetName);
  if (header === "車検状態") return getStatusListForSheet_("車検管理");
  return null;
}

function ensureSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("設定");
  if (!sheet) sheet = ss.insertSheet("設定");

  const headers = ["設定種別", "値", "備考"];
  const current = sheet.getRange(1, 1, 1, 3).getValues()[0];
  if (current[0] !== "設定種別" || current[1] !== "値") {
    sheet.clear();
    sheet.getRange(1, 1, 1, 3).setValues([headers]);
    const rows = [];
    DEFAULT_STAFF_MEMBERS.forEach(name => rows.push(["担当者", name, "担当プルダウン用"]));
    DEFAULT_PERSONAL_MEMBERS.forEach(name => rows.push(["既読確認者", name, "個人確認列用"]));
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  formatSheetBase_(sheet, headers.length);
  return sheet;
}

let SETTINGS_CACHE_ = null;
function clearSettingsCache_() { SETTINGS_CACHE_ = null; }
function getSettingsCache_() {
  if (SETTINGS_CACHE_) return SETTINGS_CACHE_;
  const cache = { "担当者": DEFAULT_STAFF_MEMBERS.slice(), "既読確認者": DEFAULT_PERSONAL_MEMBERS.slice() };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  if (!sheet || sheet.getLastRow() < 2) {
    SETTINGS_CACHE_ = cache;
    return cache;
  }
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const staff = [];
  const personal = [];
  values.forEach(r => {
    const type = String(r[0] || "").trim();
    const value = String(r[1] || "").trim();
    if (!value) return;
    if (type === "担当者") staff.push(value);
    if (type === "既読確認者") personal.push(value);
  });
  if (staff.length) cache["担当者"] = staff;
  if (personal.length) cache["既読確認者"] = personal;
  SETTINGS_CACHE_ = cache;
  return cache;
}
function getStaffMembers_() { return getSettingsCache_()["担当者"].slice(); }
function getPersonalMembers_() { return getSettingsCache_()["既読確認者"].slice(); }

function setupSheet_(sheet, headers) {
  ensureSheetMinimumRows_(sheet);
  ensureColumns_(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  formatSheetBase_(sheet, headers.length);
  createFilterSafely_(sheet, headers.length);
  clearSetupRangeValidations_(sheet, headers.length);
  applyDataValidationByHeaders_(sheet, headers);
  ensureIdsForSheet_(sheet);
  applyColorRules(sheet);
}

function resetSheetForSetup_(sheet, headerCount) {
  removeFilter_(sheet);
  showAllColumns_(sheet);
  try { sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), Math.max(sheet.getMaxColumns(), 1)).breakApart(); } catch (e) {}
  sheet.clear();
  sheet.clearFormats();
  sheet.setConditionalFormatRules([]);
  ensureColumns_(sheet, headerCount);
  if (sheet.getMaxColumns() > headerCount) sheet.deleteColumns(headerCount + 1, sheet.getMaxColumns() - headerCount);
  ensureSheetMinimumRows_(sheet);
}

function resetSheetLight_(sheet, headerCount) {
  removeFilter_(sheet);
  showAllColumns_(sheet);
  ensureColumns_(sheet, headerCount);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = Math.max(sheet.getLastColumn(), headerCount);
  sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  sheet.getRange(1, 1, Math.min(sheet.getMaxRows(), Math.max(lastRow, MIN_DATA_ROWS_FOR_VALIDATION + 1)), Math.max(lastCol, headerCount)).clearDataValidations();
}

function ensureColumns_(sheet, count) {
  const maxCols = sheet.getMaxColumns();
  if (maxCols < count) sheet.insertColumnsAfter(maxCols, count - maxCols);
}

function ensureSheetMinimumRows_(sheet) {
  const targetRows = MIN_DATA_ROWS_FOR_VALIDATION + 1;
  if (sheet.getMaxRows() < targetRows) sheet.insertRowsAfter(sheet.getMaxRows(), targetRows - sheet.getMaxRows());
}

function ensureRowsAfterEdit_(sheet, row) {
  if (!sheet || row <= 1) return;
  const remaining = sheet.getMaxRows() - row;
  if (remaining > AUTO_EXTEND_TRIGGER_MARGIN) return;
  sheet.insertRowsAfter(sheet.getMaxRows(), AUTO_EXTEND_ROWS_BUFFER);
  const headers = getSheetHeaders_()[sheet.getName()];
  if (headers) applyDataValidationByHeaders_(sheet, headers);
}

function getApplyRowCount_(sheet) {
  const maxRows = sheet.getMaxRows();
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const targetRows = Math.min(maxRows, Math.max(MIN_DATA_ROWS_FOR_VALIDATION + 1, lastRow + AUTO_EXTEND_ROWS_BUFFER));
  return Math.max(targetRows - 1, 1);
}

function clearSetupRangeValidations_(sheet, colCount) {
  const rows = Math.min(sheet.getMaxRows(), getApplyRowCount_(sheet) + 1);
  sheet.getRange(1, 1, rows, colCount).clearDataValidations();
}

function applyDataValidationByHeaders_(sheet, headers) {
  const sheetName = sheet.getName();
  if (isAutoOutputSheet_(sheetName)) return;
  headers.forEach((header, index) => {
    const col = index + 1;
    if (DATE_HEADERS.includes(header)) setDateColumn(sheet, col);
    if (TIME_HEADERS.includes(header)) setTimeColumn(sheet, col);
    if (VEHICLE_NUMBER_HEADERS.includes(header)) setTextColumn(sheet, col);
    const list = getDropdownListForHeader_(sheetName, header);
    if (list) setDropdown(sheet, col, list);
  });
  setCheckboxesForDataRows(sheet);
}

function setDateColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sheet.getRange(2, col, rowCount).setNumberFormat("yyyy/mm/dd").setDataValidation(rule);
}

function setTimeColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount).setNumberFormat("hh:mm");
}

function setTextColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount).setNumberFormat("@").clearDataValidations();
}

function setDropdown(sheet, col, list) {
  const rowCount = getApplyRowCount_(sheet);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(2, col, rowCount).setDataValidation(rule);
}

function setCheckboxesForDataRows(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  const headers = getHeaders_(sheet);
  const checkboxHeaders = [READ_HEADER, ...getPersonalMembers_()];
  if (sheet.getName() === "運転免許管理") checkboxHeaders.push(...LICENSE_TYPE_HEADERS);
  const rowCount = getApplyRowCount_(sheet);
  checkboxHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;
    const range = sheet.getRange(2, col, rowCount, 1);
    range.insertCheckboxes();
  });
}

function setCheckboxesForEditedRow_(sheet, row) {
  const headers = getHeaders_(sheet);
  const checkboxHeaders = [READ_HEADER, ...getPersonalMembers_()];
  if (sheet.getName() === "運転免許管理") checkboxHeaders.push(...LICENSE_TYPE_HEADERS);
  checkboxHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;
    sheet.getRange(row, col).insertCheckboxes();
  });
}

function getHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
}

function objectFromRow(headers, valuesRow) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = valuesRow[i]);
  return obj;
}

function writeObjectsToSheet(sheet, objects) {
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  if (!headers.length) return;
  if (!objects || objects.length === 0) return;
  const personal = getPersonalMembers_();
  const values = objects.map(obj => headers.map(header => {
    if (header === READ_HEADER || personal.includes(header) || LICENSE_TYPE_HEADERS.includes(header)) return obj[header] === true;
    return obj[header] !== undefined ? obj[header] : "";
  }));
  sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  ensureIdsForSheet_(sheet);
  setCheckboxesForDataRows(sheet);
}

function collectRowsAsObjects(ss, sheetName, condition, mapper, rows) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach(v => {
    const row = objectFromRow(headers, v);
    if (condition(row)) rows.push(mapper(row));
  });
}

function createScheduleList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["一覧スケジュール"];
  let sheet = ss.getSheetByName("一覧スケジュール");
  if (!sheet) sheet = ss.insertSheet("一覧スケジュール");
  resetSheetLight_(sheet, headers.length);
  setupSheet_(sheet, headers);

  const rows = [];
  collectRowsAsObjects(ss, "出先予定", r => r["日付"], r => scheduleRow_(r["日付"], "出先予定", joinText(r["行き先"], r["用件"]), r["担当"], r["状態"], getNoticeText(r["日付"], r["状態"]), r["電話対応"], r, "出先予定"), rows);
  collectRowsAsObjects(ss, "工事予定", r => r["開始日"], r => scheduleRow_(r["開始日"], "工事予定", joinText(r["工事名"], joinText(r["現場"], r["依頼主"])), r["担当"], r["状態"], getNoticeText(r["開始日"], r["状態"]), r["電話対応"], r, "工事予定"), rows);
  collectRowsAsObjects(ss, "会議予定", r => r["日付"], r => scheduleRow_(r["日付"], "会議", joinText(r["会議名"], r["内容"]), r["担当"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "会議予定"), rows);
  collectRowsAsObjects(ss, "行事予定", r => r["日付"], r => scheduleRow_(r["日付"], "行事", joinText(r["行事名"], r["内容"]), r["担当"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "行事予定"), rows);
  collectRowsAsObjects(ss, "電話履歴", r => r["日付"] || r["相手"], r => scheduleRow_(r["日付"] || new Date(), "電話履歴", joinText(r["相手"], r["内容"]), r["担当"], r["電話対応"], getNoticeText(r["日付"], r["電話対応"]), r["電話対応"], r, "電話履歴"), rows);
  collectRowsAsObjects(ss, "車検管理", r => r["車検期限"] || r["車検予定日"], r => scheduleRow_(r["車検予定日"] || r["車検期限"], "車検", joinText(r["車両名"], r["車番"]), "", r["車検状態"], getNoticeText(r["車検予定日"] || r["車検期限"], r["車検状態"]), "", r, "車検管理"), rows);
  collectRowsAsObjects(ss, "社用車予約", r => r["日付"] || r["社用車"], r => scheduleRow_(r["日付"] || new Date(), "社用車予約", joinText(r["社用車"], joinText(r["行き先"], r["用途"])), r["利用者"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "社用車予約"), rows);
  collectRowsAsObjects(ss, "日報", r => r["日付"] || r["現場"], r => scheduleRow_(r["日付"] || new Date(), "日報", joinText(r["現場"], r["作業内容"]), r["担当"] || r["入力者"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "日報"), rows);
  collectRowsAsObjects(ss, "運転免許管理", r => r["更新期限"] || r["所有者"], r => scheduleRow_(r["更新期限"] || new Date(), "運転免許", joinText(r["所有者"], buildLicenseTypeText_(r)), r["所有者"], r["状態"], getNoticeText(r["更新期限"], r["状態"]), "", r, "運転免許管理"), rows);
  collectRowsAsObjects(ss, "資格管理", r => r["更新期限"] || r["資格名"], r => scheduleRow_(r["更新期限"] || new Date(), "資格", joinText(r["所有者"], joinText(r["資格名"], r["区分"])), r["所有者"], r["状態"], getNoticeText(r["更新期限"], r["状態"]), "", r, "資格管理"), rows);
  collectRowsAsObjects(ss, "備品修理管理", r => r["返却予定日"] || r["備品名"], r => scheduleRow_(r["返却予定日"] || r["修理依頼日"] || r["購入日"] || new Date(), "備品修理", joinText(r["備品名"], joinText(r["修理業者"], r["内容"])), r["修理依頼者"], r["状態"], getNoticeText(r["返却予定日"] || r["修理依頼日"], r["状態"]), "", r, "備品修理管理"), rows);
  collectRowsAsObjects(ss, "個人ToDo", r => r["期限"] || r["内容"], r => scheduleRow_(r["期限"] || r["登録日"] || new Date(), "ToDo", r["内容"], r["担当"], r["状態"], getNoticeText(r["期限"], r["状態"]), "", r, "個人ToDo"), rows);

  rows.sort((a, b) => toTime_(a["日付"]) - toTime_(b["日付"]));
  writeObjectsToSheet(sheet, rows);
  formatSheetByName_("一覧スケジュール");
}

function scheduleRow_(date, type, content, assignee, status, notice, phone, sourceRow, sourceName) {
  const row = {
    "日付": date,
    "種類": type,
    "内容": content,
    "担当": assignee,
    "状態": status,
    "通知": notice,
    "電話対応": phone,
    "既読": isTruthy_(sourceRow[READ_HEADER]),
    "備考": sourceRow["備考"] || sourceRow["対応メモ"] || "",
    "元シート": sourceName
  };
  getPersonalMembers_().forEach(name => row[name] = isTruthy_(sourceRow[name]));
  return row;
}

function createAlertList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("一覧スケジュール");
  const headers = getSheetHeaders_()["要確認一覧"];
  let sheet = ss.getSheetByName("要確認一覧");
  if (!sheet) sheet = ss.insertSheet("要確認一覧");
  resetSheetLight_(sheet, headers.length);
  setupSheet_(sheet, headers);

  const priority = { "期限切れ": 1, "今日": 2, "3日以内": 3, "7日以内": 4, "未確認": 5, "差戻し": 1, "未精算": 4, "重要": 2 };
  let rows = [];
  if (source && source.getLastRow() >= 2) {
    const srcHeaders = getHeaders_(source);
    const values = source.getRange(2, 1, source.getLastRow() - 1, srcHeaders.length).getValues();
    rows = values.map(v => objectFromRow(srcHeaders, v)).filter(r => shouldIncludeInAlertList_(r, priority));
  }

  appendNoticeAlerts_(ss, rows);
  appendDailyReceiptAlerts_(ss, rows);

  rows.sort((a, b) => {
    const pa = priority[a["通知"]] || 99;
    const pb = priority[b["通知"]] || 99;
    if (pa !== pb) return pa - pb;
    return toTime_(a["日付"]) - toTime_(b["日付"]);
  });

  writeObjectsToSheet(sheet, rows.map(r => ({
    "日付": r["日付"],
    "種類": r["種類"],
    "内容": r["内容"],
    "担当": r["担当"],
    "状態": r["状態"],
    "通知": r["通知"],
    "電話対応": r["電話対応"],
    "備考": r["備考"] || "",
    "元シート": r["元シート"]
  })));
  formatSheetByName_("要確認一覧");
}

function shouldIncludeInAlertList_(row, priority) {
  const notice = String(row["通知"] || "").trim();
  const status = String(row["状態"] || "").trim();
  const source = String(row["元シート"] || row["種類"] || "").trim();
  const phone = String(row["電話対応"] || "").trim();
  if (!priority[notice]) return false;

  if (source === "電話履歴" || row["種類"] === "電話履歴") return phone !== "完了";
  if (source === "車検管理" || row["種類"] === "車検") return !["更新済"].includes(status);
  if (source === "運転免許管理" || row["種類"] === "運転免許") return !["有効", "更新済"].includes(status);
  if (source === "資格管理" || row["種類"] === "資格") return !["有効", "更新済"].includes(status);
  if (source === "備品修理管理" || row["種類"] === "備品修理") return !["返却済", "完了", "修理不可"].includes(status);
  if (source === "社用車予約" || row["種類"] === "社用車予約") return !["返却済", "中止"].includes(status);
  if (source === "個人ToDo" || row["種類"] === "ToDo") return !["完了", "中止"].includes(status);
  return !["完了", "確認済", "更新済", "返却済", "中止"].includes(status);
}

function appendNoticeAlerts_(ss, rows) {
  const sheet = ss.getSheetByName("お知らせ");
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach(v => {
    const r = objectFromRow(headers, v);
    if (r["重要度"] === "高" && !isReadByAll_(r)) {
      rows.push({ "日付": r["投稿日"] || new Date(), "種類": "お知らせ", "内容": joinText(r["タイトル"], r["内容"]), "担当": r["投稿者"], "状態": "要確認", "通知": "重要", "電話対応": "", "備考": r["備考"] || "", "元シート": "お知らせ" });
    }
  });
}

function appendDailyReceiptAlerts_(ss, rows) {
  const sheet = ss.getSheetByName("日報レシート管理");
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach(v => {
    const r = objectFromRow(headers, v);
    const accounting = String(r["経理確認"] || "").trim();
    const settlement = String(r["精算状態"] || "").trim();
    if (!["差戻し", "未確認"].includes(accounting) && settlement !== "未精算") return;
    rows.push({
      "日付": r["日付"] || new Date(),
      "種類": "日報レシート",
      "内容": joinText(r["支払先"], joinText(r["内容"], r["金額"] ? String(r["金額"]) + "円" : "")),
      "担当": r["担当"],
      "状態": joinText(accounting, settlement) || "要確認",
      "通知": accounting === "差戻し" ? "差戻し" : settlement === "未精算" ? "未精算" : "未確認",
      "電話対応": "",
      "備考": r["備考"] || joinText(r["現場"], r["区分"]),
      "元シート": "日報レシート管理"
    });
  });
}

function updateNoticeColumnForSheet_(sheet) {
  const headers = getHeaders_(sheet);
  const noticeCol = headers.indexOf("通知") + 1;
  if (noticeCol <= 0 || sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const notices = values.map(v => {
    const r = objectFromRow(headers, v);
    return [getNoticeForRow_(sheet.getName(), r)];
  });
  sheet.getRange(2, noticeCol, notices.length, 1).setValues(notices);
}

function updateNoticeForEditedRow_(sheet, row) {
  const headers = getHeaders_(sheet);
  const noticeCol = headers.indexOf("通知") + 1;
  if (noticeCol <= 0) return;
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = objectFromRow(headers, values);
  sheet.getRange(row, noticeCol).setValue(getNoticeForRow_(sheet.getName(), obj));
}

function getNoticeForRow_(sheetName, row) {
  if (sheetName === "日報レシート管理") {
    const accounting = String(row["経理確認"] || "").trim();
    const settlement = String(row["精算状態"] || "").trim();
    if (accounting === "差戻し") return "差戻し";
    if (settlement === "未精算") return "未精算";
    if (accounting === "未確認") return "未確認";
    return "";
  }

  if (sheetName === "車検管理") return getNoticeText(row["車検予定日"] || row["車検期限"], row["車検状態"]);
  if (sheetName === "運転免許管理") return getNoticeText(row["更新期限"], row["状態"]);
  if (sheetName === "資格管理") return getNoticeText(row["更新期限"], row["状態"]);
  if (sheetName === "備品修理管理") return getNoticeText(row["返却予定日"] || row["修理依頼日"], row["状態"]);
  if (sheetName === "個人ToDo") return getNoticeText(row["期限"], row["状態"]);
  if (sheetName === "工事予定") return getNoticeText(row["開始日"], row["状態"]);
  if (sheetName === "行事予定" || sheetName === "会議予定" || sheetName === "出先予定" || sheetName === "社用車予約" || sheetName === "電話履歴" || sheetName === "日報") return getNoticeText(row["日付"], row["状態"] || row["電話対応"]);
  return "";
}

function getNoticeText(dateValue, status) {
  if (!dateValue) return "";
  const s = String(status || "").trim();
  if (["完了", "更新済", "確認済", "返却済", "精算済", "中止", "修理不可"].includes(s)) return "";

  const d = toDateOnly_(dateValue);
  if (!d) return "";
  const today = toDateOnly_(new Date());
  const diff = Math.floor((d.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));
  if (diff < 0) return "期限切れ";
  if (diff === 0) return "今日";
  if (diff <= 3) return "3日以内";
  if (diff <= 7) return "7日以内";
  if (diff <= 30) return "30日以内";
  return "";
}

function toDateOnly_(value) {
  if (!value) return null;
  let d = value instanceof Date ? new Date(value) : new Date(value);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function toTime_(value) {
  const d = toDateOnly_(value);
  return d ? d.getTime() : 0;
}

function ensureIdsForSheet_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const idHeader = SHEET_ID_HEADERS[sheet.getName()];
  if (!idHeader) return;
  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf(idHeader) + 1;
  if (idCol <= 0) return;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const ids = values.map(row => {
    const current = row[idCol - 1];
    const hasData = row.some((v, i) => i !== idCol - 1 && v !== "" && v !== null);
    return [hasData && !current ? buildRecordId_(sheet.getName()) : (current || "")];
  });
  sheet.getRange(2, idCol, ids.length, 1).setValues(ids);
}

function ensureIdForEditedRow_(sheet, row) {
  const idHeader = SHEET_ID_HEADERS[sheet.getName()];
  if (!idHeader) return;
  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf(idHeader) + 1;
  if (idCol <= 0) return;
  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  if (values[idCol - 1]) return;
  const hasData = values.some((v, i) => i !== idCol - 1 && v !== "" && v !== null);
  if (hasData) sheet.getRange(row, idCol).setValue(buildRecordId_(sheet.getName()));
}

function buildRecordId_(sheetName) {
  const prefixMap = { "出先予定":"OUT", "工事予定":"CON", "会議予定":"MTG", "行事予定":"EVT", "作業状況":"WRK", "電話履歴":"TEL", "車検管理":"CAR", "車検履歴":"CARH", "社用車予約":"RES", "日報":"DR", "日報レシート管理":"RCT", "運転免許管理":"LIC", "資格管理":"QAL", "備品修理管理":"REP", "お知らせ":"NEWS", "個人ToDo":"TODO", "フィードバック":"FB", "帳簿PDF履歴":"LEDGER" };
  const prefix = prefixMap[sheetName] || "ID";
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  return prefix + "-" + now + "-" + Utilities.getUuid().slice(0, 8);
}

function addSampleMainSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheetsIfMissing_(["出先予定", "工事予定", "会議予定", "行事予定", "電話履歴", "社用車予約", "お知らせ", "個人ToDo"]);
  const today = new Date();

  writeObjectsToSheet(ss.getSheetByName("出先予定"), [
    {"日付": addDays_(today, 0), "行き先": "サンプル市役所", "用件": "書類提出", "担当": "山田", "社用車": "軽トラ②", "状態": "予定", "電話対応": "", "備考": "SAMPLE"},
    {"日付": addDays_(today, 2), "行き先": "サンプル現場A", "用件": "進捗確認", "担当": "鈴木", "社用車": "4t", "状態": "予定", "電話対応": "", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("工事予定"), [
    {"工事名": "サンプル舗装工事", "現場": "サンプル町1丁目", "依頼主": "サンプル建設", "連絡先": "018-000-0000", "契約金額": 1200000, "開始日": addDays_(today, 1), "終了日": addDays_(today, 10), "状態": "着工前", "担当": "高橋", "電話対応": "未対応", "備考": "SAMPLE"},
    {"工事名": "サンプル排水工事", "現場": "サンプル町2丁目", "依頼主": "サンプル商事", "連絡先": "018-111-1111", "契約金額": 850000, "開始日": addDays_(today, -2), "終了日": addDays_(today, 5), "状態": "施工中", "担当": "田中", "電話対応": "対応中", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("会議予定"), [
    {"日付": addDays_(today, 0), "開始時刻": "10:00", "終了時刻": "11:00", "会議名": "サンプル工程会議", "内容": "今週の工程確認", "担当": "山田", "状態": "予定", "資料": "", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("行事予定"), [
    {"日付": addDays_(today, 3), "開始時刻": "09:00", "終了時刻": "12:00", "行事名": "サンプル安全講習", "内容": "安全教育", "担当": "鈴木", "状態": "予定", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("電話履歴"), [
    {"日付": addDays_(today, 0), "時間": "09:30", "相手": "サンプル取引先", "内容": "折返し依頼", "電話対応": "折返し", "担当": "高橋", "対応メモ": "午後に折返し", "備考": "SAMPLE"},
    {"日付": addDays_(today, -1), "時間": "15:00", "相手": "サンプル顧客", "内容": "日程確認", "電話対応": "完了", "担当": "田中", "対応メモ": "確認済", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("社用車予約"), [
    {"日付": addDays_(today, 1), "開始時刻": "08:30", "終了時刻": "17:00", "社用車": "軽トラ②", "利用者": "山田", "行き先": "サンプル現場", "用途": "資材運搬", "状態": "予定", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("お知らせ"), [
    {"投稿日": today, "タイトル": "サンプルお知らせ", "内容": "ベータ確認用のお知らせです", "投稿者": "山田", "重要度": "高", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("個人ToDo"), [
    {"登録日": today, "担当": "山田", "内容": "サンプルToDo確認", "期限": addDays_(today, 2), "状態": "未着手", "備考": "SAMPLE"}
  ]);

  if (!IS_BATCH_SETUP_) {
    refreshInputSheets();
    toast_("主要予定系サンプルを追加しました");
  }
}

function addSampleVehicleLicenseEquipment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheetsIfMissing_(["車検管理", "運転免許管理", "資格管理", "備品修理管理"]);
  const today = new Date();

  writeObjectsToSheet(ss.getSheetByName("車検管理"), [
    {"車両名": "サンプル車両A", "車番": "秋田500あ0001", "車検期限": addDays_(today, 3), "車検予定日": addDays_(today, 5), "保険期限": addDays_(today, 30), "車検状態": "未対応", "写真": "", "備考": "SAMPLE"},
    {"車両名": "サンプル車両B", "車番": "秋田500あ0002", "車検期限": addDays_(today, -1), "車検予定日": addDays_(today, 2), "保険期限": addDays_(today, 40), "車検状態": "予約済", "写真": "", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("運転免許管理"), [
    {"所有者": "山田", "普通": true, "準中型": true, "中型": false, "大型": false, "大型特殊": false, "けん引": false, "二種": false, "取得日": addDays_(today, -1000), "更新期限": addDays_(today, 7), "コピー有無": "有", "状態": "更新予定", "備考": "SAMPLE"},
    {"所有者": "鈴木", "普通": true, "準中型": false, "中型": true, "大型": false, "大型特殊": false, "けん引": false, "二種": false, "取得日": addDays_(today, -1200), "更新期限": addDays_(today, 120), "コピー有無": "有", "状態": "有効", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("資格管理"), [
    {"所有者": "高橋", "資格名": "玉掛け技能講習", "区分": "技能講習", "取得日": addDays_(today, -600), "更新期限": addDays_(today, 5), "コピー有無": "未確認", "状態": "更新予定", "備考": "SAMPLE"},
    {"所有者": "田中", "資格名": "小型移動式クレーン", "区分": "技能講習", "取得日": addDays_(today, -900), "更新期限": addDays_(today, 200), "コピー有無": "有", "状態": "有効", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("備品修理管理"), [
    {"購入日": addDays_(today, -300), "備品名": "サンプル発電機", "修理業者": "サンプル修理店", "内容": "エンジン始動不良", "修理依頼者": "山田", "修理依頼日": addDays_(today, -2), "返却予定日": addDays_(today, 3), "返却済": "未", "状態": "修理中", "備考": "SAMPLE"},
    {"購入日": addDays_(today, -500), "備品名": "サンプル工具", "修理業者": "サンプル工具店", "内容": "部品破損", "修理依頼者": "鈴木", "修理依頼日": addDays_(today, -5), "返却予定日": addDays_(today, -1), "返却済": "未", "状態": "依頼済", "備考": "SAMPLE"}
  ]);

  if (!IS_BATCH_SETUP_) {
    repairVehicleInspectionSheet();
    refreshInputSheets();
    toast_("車検・免許・資格・備品サンプルを追加しました");
  }
}

function addSampleDailyReceiptFeedback() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheetsIfMissing_(["日報", "日報レシート管理", "フィードバック"]);
  const today = new Date();

  writeObjectsToSheet(ss.getSheetByName("日報"), [
    {"日付": today, "入力者": "山田", "担当": "山田", "現場": "サンプル現場A", "作業内容": "掘削・整地", "進捗": "予定通り", "問題点": "特になし", "明日の予定": "砕石敷均し", "他現場状況": "", "写真": "", "日報文章": "", "状態": "下書き", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("日報レシート管理"), [
    {"日付": today, "担当": "山田", "現場": "サンプル現場A", "支払先": "サンプルGS", "内容": "燃料代", "区分": "燃料", "金額": 5000, "支払方法": "立替", "レシート写真": "", "経理確認": "未確認", "精算状態": "未精算", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("フィードバック"), [
    {"日付": today, "記入者": "山田", "確認方法": "PC", "対象シート": "車検管理", "気になった内容": "車番が日付にならないか確認", "区分": "動作不備", "困り度": "高", "対応状況": "未対応", "対応方針": "車番列を文字列固定", "対応メモ": "SAMPLE", "対応日": ""}
  ]);

  if (!IS_BATCH_SETUP_) {
    refreshInputSheets();
    toast_("日報・レシート・FBサンプルを追加しました");
  }
}

function deleteBetaSamples() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert("サンプル削除", "備考にSAMPLEが入っている行を削除します。続行しますか？", ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(getSheetHeaders_()).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const headers = getHeaders_(sheet);
    const memoCol = headers.indexOf("備考") + 1;
    if (memoCol <= 0) return;
    for (let r = sheet.getLastRow(); r >= 2; r--) {
      const memo = String(sheet.getRange(r, memoCol).getValue() || "");
      if (memo.indexOf("SAMPLE") >= 0) sheet.deleteRow(r);
    }
  });
  refreshAll();
  toast_("ベータ用サンプルを削除しました");
}

function setupSheetsIfMissing_(names) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = getSheetHeaders_();
  names.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (map[name]) setupSheet_(sheet, map[name]);
  });
}

function repairVehicleInspectionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("車検管理");
  if (!sheet) return;
  const headers = getSheetHeaders_()["車検管理"];
  setupSheet_(sheet, headers);
  const h = getHeaders_(sheet);
  const carNoCol = h.indexOf("車番") + 1;
  if (carNoCol > 0) {
    sheet.getRange(2, carNoCol, getApplyRowCount_(sheet), 1).setNumberFormat("@").clearDataValidations();
  }
  updateNoticeColumnForSheet_(sheet);
  applyColorRules(sheet);
  formatSheetByName_("車検管理");
  toast_("車検管理を修復しました");
}

function migrateLegacyColumnsV970() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert("旧列名をv9.7.0へ移行", "既存シートの旧列名をv9.7.0仕様へ可能な範囲で変更します。バックアップ後の実行を推奨します。続行しますか？", ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const renameMap = {
    "車検管理": {"次回車検期限": "車検予定日", "状態": "車検状態"},
    "備品修理管理": {"登録日": "購入日", "場所": "修理業者", "担当": "修理依頼者"},
    "電話履歴": {"日時": "日付", "メモ": "対応メモ"},
    "免許資格管理": {}
  };

  Object.keys(renameMap).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const headers = getHeaders_(sheet);
    Object.keys(renameMap[sheetName]).forEach(oldName => {
      const col = headers.indexOf(oldName) + 1;
      if (col > 0) sheet.getRange(1, col).setValue(renameMap[sheetName][oldName]);
    });
  });

  const old = ss.getSheetByName("免許資格管理");
  if (old) old.setName("免許資格管理_旧");

  setupSheetsIfMissing_(Object.keys(getSheetHeaders_()));
  repairVehicleInspectionSheet();
  runDataConsistencyCheckAndFix();
  toast_("旧列名の移行を実行しました");
}

function runDataConsistencyCheckAndFix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(getSheetHeaders_()).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const headers = getSheetHeaders_()[name];
    setupSheet_(sheet, headers);
    updateNoticeColumnForSheet_(sheet);
    ensureIdsForSheet_(sheet);
  });
  applyTimeFormatsToAllSheets();
  repairVehicleInspectionSheet();
  toast_("データ整合性チェック・軽微修正が完了しました");
}

function createFeatureCheckSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("機能チェック");
  if (!sheet) sheet = ss.insertSheet("機能チェック");
  sheet.clear();
  sheet.clearFormats();

  const rows = [["区分", "確認項目", "結果", "詳細"]];
  const headersMap = getSheetHeaders_();
  Object.keys(headersMap).forEach(name => {
    const s = ss.getSheetByName(name);
    rows.push(["シート", name + " が存在する", s ? "OK" : "NG", s ? "" : "シートがありません"]);
    if (s) {
      const h = getHeaders_(s);
      const missing = headersMap[name].filter(x => !h.includes(x));
      rows.push(["ヘッダー", name + " 必須ヘッダー", missing.length === 0 ? "OK" : "NG", missing.join(", ")]);
    }
  });

  const vehicle = ss.getSheetByName("車検管理");
  if (vehicle) {
    const h = getHeaders_(vehicle);
    rows.push(["車検", "車番列がある", h.includes("車番") ? "OK" : "NG", ""]);
    rows.push(["車検", "旧列 次回車検期限 がない", h.includes("次回車検期限") ? "NG" : "OK", ""]);
    rows.push(["車検", "旧列 次回保険期限 がない", h.includes("次回保険期限") ? "NG" : "OK", ""]);
  }

  const schedule = ss.getSheetByName("一覧スケジュール");
  if (schedule && schedule.getLastRow() >= 2) {
    const vals = schedule.getRange(2, 2, schedule.getLastRow() - 1, 1).getValues().flat().map(String);
    ["車検", "運転免許", "資格", "備品修理", "工事予定", "電話履歴"].forEach(type => {
      rows.push(["一覧", type + " が一覧に出る", vals.includes(type) ? "OK" : "確認", vals.includes(type) ? "" : "サンプルがない/更新前の可能性"]);
    });
  }

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 420);
  sheet.getRange(1, 1, rows.length, rows[0].length).setWrap(true).setVerticalAlignment("middle");
  applyStatusBackground_(sheet, 3, 2, rows.length - 1);
  toast_("機能チェックシートを作成しました");
}

function createAssigneeUnreadSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["担当別未読"];
  let sheet = ss.getSheetByName("担当別未読");
  if (!sheet) sheet = ss.insertSheet("担当別未読");
  resetSheetLight_(sheet, headers.length);
  setupSheet_(sheet, headers);

  const schedule = ss.getSheetByName("一覧スケジュール");
  const staff = getStaffMembers_();
  const rows = staff.map(name => ({"担当": name, "未読件数": 0, "未完了ToDo": 0, "未対応電話": 0, "今日の予定": 0}));
  const map = {};
  rows.forEach(r => map[r["担当"]] = r);

  if (schedule && schedule.getLastRow() >= 2) {
    const h = getHeaders_(schedule);
    const values = schedule.getRange(2, 1, schedule.getLastRow() - 1, h.length).getValues();
    values.forEach(v => {
      const r = objectFromRow(h, v);
      const a = r["担当"];
      if (!map[a]) return;
      if (!isTruthy_(r[READ_HEADER])) map[a]["未読件数"]++;
      if (r["種類"] === "ToDo" && !["完了", "中止"].includes(String(r["状態"]))) map[a]["未完了ToDo"]++;
      if (r["種類"] === "電話履歴" && ["未対応", "折返し"].includes(String(r["電話対応"]))) map[a]["未対応電話"]++;
      if (getNoticeText(r["日付"], "") === "今日") map[a]["今日の予定"]++;
    });
  }

  writeObjectsToSheet(sheet, rows);
}

function createReadRateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["既読率集計"];
  let sheet = ss.getSheetByName("既読率集計");
  if (!sheet) sheet = ss.insertSheet("既読率集計");
  resetSheetLight_(sheet, headers.length);
  setupSheet_(sheet, headers);

  const schedule = ss.getSheetByName("一覧スケジュール");
  const map = {};
  if (schedule && schedule.getLastRow() >= 2) {
    const h = getHeaders_(schedule);
    const values = schedule.getRange(2, 1, schedule.getLastRow() - 1, h.length).getValues();
    values.forEach(v => {
      const r = objectFromRow(h, v);
      const type = r["種類"] || "未分類";
      if (!map[type]) map[type] = {total: 0, read: 0};
      map[type].total++;
      if (isTruthy_(r[READ_HEADER])) map[type].read++;
    });
  }

  const rows = Object.keys(map).sort().map(type => {
    const total = map[type].total;
    const read = map[type].read;
    const unread = Math.max(total - read, 0);
    return {
      "対象": type,
      "全体件数": total,
      "既読件数": read,
      "未読件数": unread,
      "既読率": total ? Math.round(read / total * 100) + "%" : ""
    };
  });

  writeObjectsToSheet(sheet, rows);
}

function createDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("ダッシュボード");
  if (!sheet) sheet = ss.insertSheet("ダッシュボード");
  sheet.clear();
  sheet.clearFormats();
  const rows = [
    ["項目", "件数", "備考"],
    ["システム版", SYSTEM_VERSION, ""],
    ["今日の予定", countNotice_("今日"), "一覧スケジュール"],
    ["期限切れ", countNotice_("期限切れ"), "一覧スケジュール"],
    ["3日以内", countNotice_("3日以内"), "一覧スケジュール"],
    ["7日以内", countNotice_("7日以内"), "一覧スケジュール"],
    ["要確認件数", getSheetDataCount_("要確認一覧"), ""],
    ["未確認レシート", countReceiptAlerts_(), ""],
    ["更新必要フラグ", isSummaryDirty_() ? "あり" : "なし", "onEdit後は日常更新推奨"]
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#d9ead3");
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 300);
}

function countNotice_(notice) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("一覧スケジュール");
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const h = getHeaders_(sheet);
  const col = h.indexOf("通知") + 1;
  if (col <= 0) return 0;
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().flat().filter(v => v === notice).length;
}

function countReceiptAlerts_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("日報レシート管理");
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const h = getHeaders_(sheet);
  const aCol = h.indexOf("経理確認") + 1;
  const sCol = h.indexOf("精算状態") + 1;
  if (aCol <= 0 || sCol <= 0) return 0;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, h.length).getValues();
  return values.filter(v => {
    const r = objectFromRow(h, v);
    return r["経理確認"] === "未確認" || r["経理確認"] === "差戻し" || r["精算状態"] === "未精算";
  }).length;
}

function getSheetDataCount_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 0;
  return Math.max(sheet.getLastRow() - 1, 0);
}

function formatInputSheetsForLongText() {
  Object.keys(getSheetHeaders_()).forEach(name => formatSheetByName_(name));
  toast_("表示幅・カラーリングを調整しました");
}

function formatActiveSheetOnly() {
  formatSheetBase_(SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getLastColumn());
  formatSheetByName_(SpreadsheetApp.getActiveSheet().getName());
  toast_("今のシートを表示調整しました");
}

function formatSheetByName_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  if (!headers.length) return;
  formatSheetBase_(sheet, headers.length);
  const widthMap = {
    "内容": 260, "備考": 260, "対応メモ": 260, "日報文章": 420, "問題点": 260, "明日の予定": 260,
    "現場": 180, "工事名": 180, "行き先": 160, "車両名": 160, "車番": 160, "写真": 120,
    "通知": 120, "状態": 120, "車検状態": 120, "電話対応": 120, "開始時刻": 90, "終了時刻": 90, "時間": 90,
    "契約金額": 110, "連絡先": 130, "PDFリンク": 260, "フィードバックID": 180
  };
  headers.forEach((h, i) => sheet.setColumnWidth(i + 1, widthMap[h] || 115));
  if (sheet.getLastRow() > 1) sheet.setRowHeights(2, Math.min(sheet.getLastRow() - 1, 200), 34);
  try { sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setWrap(true).setVerticalAlignment("middle"); } catch (e) {}
  applyColorRules(sheet);
}

function formatSheetBase_(sheet, colCount) {
  if (!sheet) return;
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, colCount).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setRowHeight(1, 42);
}

function applyColorRules(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  const headers = getHeaders_(sheet);
  if (!headers.length) return;
  const rules = [];
  const noticeCol = headers.indexOf("通知") + 1;
  const statusCol = Math.max(headers.indexOf("状態") + 1, headers.indexOf("車検状態") + 1, headers.indexOf("電話対応") + 1, headers.indexOf("対応状況") + 1);
  if (noticeCol > 0) addTextRule_(rules, sheet, noticeCol, "期限切れ", "#ea9999");
  if (noticeCol > 0) addTextRule_(rules, sheet, noticeCol, "今日", "#f4cccc");
  if (noticeCol > 0) addTextRule_(rules, sheet, noticeCol, "3日以内", "#fff2cc");
  if (noticeCol > 0) addTextRule_(rules, sheet, noticeCol, "7日以内", "#d9ead3");
  if (statusCol > 0) {
    addTextRule_(rules, sheet, statusCol, "未対応", "#f4cccc");
    addTextRule_(rules, sheet, statusCol, "折返し", "#fff2cc");
    addTextRule_(rules, sheet, statusCol, "予約済", "#fff2cc");
    addTextRule_(rules, sheet, statusCol, "施工中", "#d9ead3");
    addTextRule_(rules, sheet, statusCol, "修理中", "#d9ead3");
    addTextRule_(rules, sheet, statusCol, "完了", "#d9ead3");
    addTextRule_(rules, sheet, statusCol, "更新済", "#d9ead3");
    addTextRule_(rules, sheet, statusCol, "返却済", "#d9ead3");
  }
  sheet.setConditionalFormatRules(rules);
}

function addTextRule_(rules, sheet, col, text, color) {
  const range = sheet.getRange(2, col, Math.max(getApplyRowCount_(sheet), 1), 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build());
}

function applyStatusBackground_(sheet, col, startRow, numRows) {
  if (numRows <= 0) return;
  const range = sheet.getRange(startRow, col, numRows, 1);
  const values = range.getValues();
  const backgrounds = values.map(r => [String(r[0]) === "OK" ? "#d9ead3" : String(r[0]) === "NG" ? "#ea9999" : "#fff2cc"]);
  range.setBackgrounds(backgrounds);
}

function applyTimeFormatsToAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(getSheetHeaders_()).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const headers = getHeaders_(sheet);
    headers.forEach((h, i) => {
      if (TIME_HEADERS.includes(h)) setTimeColumn(sheet, i + 1);
      if (VEHICLE_NUMBER_HEADERS.includes(h)) setTextColumn(sheet, i + 1);
    });
  });
}

function groupScheduleByMonth() { groupSheetByMonth_("一覧スケジュール", "日付"); }
function groupArchiveByMonth() { groupSheetByMonth_("過去一覧", "日付"); }
function groupActiveSheetByMonth() { groupSheetByMonth_(SpreadsheetApp.getActiveSheet().getName(), "日付"); }

function groupSheetByMonth_(sheetName, dateHeader) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 3) return;
  const headers = getHeaders_(sheet);
  const dateCol = headers.indexOf(dateHeader) + 1;
  if (dateCol <= 0) {
    toast_("日付列が見つかりません: " + sheetName);
    return;
  }
  clearRowGroups_(sheet);
  const values = sheet.getRange(2, dateCol, sheet.getLastRow() - 1, 1).getValues();
  let start = 2;
  let currentKey = "";
  for (let i = 0; i < values.length; i++) {
    const key = formatMonthKey_(values[i][0]);
    if (i === 0) currentKey = key;
    if (key !== currentKey) {
      groupRangeRows_(sheet, start, i + 1);
      start = i + 2;
      currentKey = key;
    }
  }
  groupRangeRows_(sheet, start, sheet.getLastRow());
  toast_(sheetName + "を月別グループ化しました");
}

function groupRangeRows_(sheet, startRow, endRow) {
  if (endRow <= startRow) return;
  try { sheet.getRange(startRow, 1, endRow - startRow + 1, 1).shiftRowGroupDepth(1); } catch (e) {}
}

function clearRowGroups_(sheet) {
  try {
    for (let i = 0; i < 8; i++) {
      sheet.getRange(1, 1, sheet.getMaxRows(), 1).shiftRowGroupDepth(-1);
    }
  } catch (e) {}
}

function formatMonthKey_(value) {
  const d = toDateOnly_(value);
  if (!d) return "日付なし";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy年MM月");
}

function movePastItemsToArchive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("一覧スケジュール");
  let archive = ss.getSheetByName("過去一覧");
  if (!archive) archive = ss.insertSheet("過去一覧");
  setupSheet_(archive, getSheetHeaders_()["過去一覧"]);
  if (!source || source.getLastRow() < 2) return;
  const h = getHeaders_(source);
  const values = source.getRange(2, 1, source.getLastRow() - 1, h.length).getValues();
  const today = toDateOnly_(new Date());
  const rows = [];
  values.forEach(v => {
    const r = objectFromRow(h, v);
    const d = toDateOnly_(r["日付"]);
    const status = String(r["状態"] || "");
    if (d && d.getTime() < today.getTime() && ["完了", "確認済", "更新済", "返却済", "中止"].includes(status)) {
      rows.push({"日付": r["日付"], "種類": r["種類"], "内容": r["内容"], "担当": r["担当"], "状態": r["状態"], "通知": r["通知"], "電話対応": r["電話対応"], "元シート": r["元シート"], "備考": r["備考"], "移動日": new Date()});
    }
  });
  if (rows.length) writeObjectsToSheet(archive, rows);
  toast_("過去一覧へ移動候補を追加しました: " + rows.length + "件");
}

function monthlyMaintenance() {
  movePastItemsToArchive();
  groupArchiveByMonth();
}

function createProcedureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("手順シート");
  if (!sheet) sheet = ss.insertSheet("手順シート");
  sheet.clear();
  sheet.clearFormats();
  const rows = [
    ["区分", "順番", "実行メニュー", "使うタイミング", "注意"],
    ["最短", 1, "Apps Script全文貼り替え → 保存 → 再読み込み", "コード更新時", "貼り替え後、スプレッドシートを再読み込みします。"],
    ["最短", 2, "1. 初回セットアップ → 新規ベータ作成（前シート前提なし）", "新しい空シートで最初だけ", "既存データ入り本体では実行しません。"],
    ["最短", 3, "1. 初回セットアップ → ベータ確認準備（軽量）", "ベータ確認開始時", "サンプル追加・全体更新・機能チェックを軽量に実行します。"],
    ["日常", 1, "2. 日常運用 → 日常更新（一覧＋要確認）", "入力後・確認前", "普段はこれを使います。全体更新は保守用です。"],
    ["日常", 2, "フィードバック確認", "社内ベータ中", "気づいた点をフィードバックシートに集めます。"],
    ["仕様変更", 1, "3. 保守・テスト → 旧列名をv9.7.0へ移行", "列名変更後", "バックアップ後に実行します。"],
    ["仕様変更", 2, "AppSheet Regenerate Structure", "列追加・列名変更後", "AppSheet側で実行します。"],
    ["表示", 1, "3. 保守・テスト → 表示幅・カラーリングを調整", "表示崩れ時", "重めなので毎回は実行しません。"],
    ["月次", 1, "4. 月次・整理 → 月次整理", "月末・過去整理時", "過去一覧の整理用です。"],
    ["開発者用", 1, "3. 保守・テスト → 機能チェックシートを作成", "大きな修正後", "機能確認用です。普段は不要です。"]
  ];
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  formatSheetBase_(sheet, rows[0].length);
  sheet.setColumnWidth(1, 90);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 330);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 420);
  sheet.getRange(1, 1, rows.length, rows[0].length).setWrap(true).setVerticalAlignment("middle");
}

function createCompanyManagementExplanationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("社内管理説明");
  if (!sheet) sheet = ss.insertSheet("社内管理説明");
  sheet.clear();
  const rows = [
    ["項目", "説明"],
    ["概要", "予定・電話・車検・免許資格・備品修理・日報・レシート・ToDo・お知らせを一元管理するベータ版です。"],
    ["日常運用", "入力後に必要に応じて日常更新（一覧＋要確認）を実行します。全体更新は保守用です。"],
    ["車検管理", "車番は文字列固定です。00-00のような日付化しやすいサンプルは使いません。"],
    ["免許資格", "運転免許管理と資格管理を分離しています。旧免許資格管理は新規構成では使いません。"],
    ["注意", "新規ベータ作成は初期化系です。既存データ入りの本体では実行しないでください。"]
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  formatSheetBase_(sheet, 2);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 620);
  sheet.getRange(1, 1, rows.length, 2).setWrap(true).setVerticalAlignment("middle");
}

function createFeedbackSheet() { setupSheetsIfMissing_(["フィードバック"]); }

function ensureSupportSheets() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    ensureSupportSheetsCore_();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    reorderSheets_();
    toast_("裏方シートを作成/修復しました");
  } finally {
    lock.releaseLock();
  }
}

function ensureSupportSheetsCore_() {
  ensureSettingsSheet_();
  createSqlSheets();
  createLedgerOutputSettings();
  createProcedureSheet();
  createCompanyManagementExplanationSheet();
}

function createSqlSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {
    "SQL設計": [
      ["テーブル名", "日本語名", "カラム名", "データ型", "制約", "元シート", "元シート列", "備考"],
      ["vehicles", "車検管理", "vehicle_number", "TEXT", "NOT NULL", "車検管理", "車番", "車番は文字列で保持する"],
      ["vehicles", "車検管理", "inspection_due_date", "DATE", "", "車検管理", "車検期限", "通知判定に使用"],
      ["vehicles", "車検管理", "inspection_status", "TEXT", "CHECK", "車検管理", "車検状態", "未対応/予約済/実施済/更新済"],
      ["driver_licenses", "運転免許管理", "owner_name", "TEXT", "NOT NULL", "運転免許管理", "所有者", "担当者名に対応"],
      ["driver_licenses", "運転免許管理", "renewal_due_date", "DATE", "", "運転免許管理", "更新期限", "一覧・要確認に反映"],
      ["qualifications", "資格管理", "qualification_name", "TEXT", "NOT NULL", "資格管理", "資格名", "資格マスタ化候補"],
      ["repairs", "備品修理管理", "repair_vendor", "TEXT", "", "備品修理管理", "修理業者", "旧 場所 から変更"],
      ["daily_receipts", "日報レシート管理", "amount", "NUMERIC", "", "日報レシート管理", "金額", "帳簿PDF出力候補"]
    ],
    "移行対応表": [
      ["シート名", "列名", "SQLテーブル", "SQLカラム", "備考"],
      ["工事予定", "工事名", "construction_projects", "project_name", ""],
      ["工事予定", "依頼主", "construction_projects", "client_name", "フィードバックで追加"],
      ["工事予定", "契約金額", "construction_projects", "contract_amount", "数値/金額"],
      ["電話履歴", "日付", "phone_logs", "call_date", "日時から分離"],
      ["電話履歴", "時間", "phone_logs", "call_time", "hh:mm"],
      ["車検管理", "車番", "vehicles", "vehicle_number", "文字列"],
      ["車検管理", "車検予定日", "vehicles", "inspection_scheduled_date", "次回車検期限から変更"],
      ["運転免許管理", "所有者", "driver_licenses", "owner_name", "担当から変更"],
      ["資格管理", "資格名", "qualifications", "qualification_name", ""],
      ["備品修理管理", "購入日", "repairs", "purchase_date", "登録日から変更"],
      ["備品修理管理", "修理依頼者", "repairs", "requester_name", "担当から変更"]
    ],
    "SQLサンプル集": [
      ["区分", "用途", "SQL", "説明"],
      ["期限確認", "期限近い車検", "SELECT * FROM vehicles WHERE inspection_due_date <= CURRENT_DATE + INTERVAL '7 days';", "7日以内の車検確認"],
      ["電話", "未対応電話", "SELECT * FROM phone_logs WHERE phone_status <> '完了';", "折返し・未対応の抽出"],
      ["免許", "更新期限が近い免許", "SELECT * FROM driver_licenses WHERE renewal_due_date <= CURRENT_DATE + INTERVAL '30 days';", "30日以内の免許更新確認"],
      ["資格", "更新期限が近い資格", "SELECT * FROM qualifications WHERE renewal_due_date <= CURRENT_DATE + INTERVAL '30 days';", "30日以内の資格更新確認"],
      ["備品", "修理中の備品", "SELECT * FROM repairs WHERE repair_status IN ('依頼済','修理中');", "未完了修理の確認"]
    ]
  };

  Object.keys(data).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    sheet.clearFormats();
    sheet.getRange(1, 1, data[name].length, data[name][0].length).setValues(data[name]);
    formatSheetBase_(sheet, data[name][0].length);
    sheet.autoResizeColumns(1, data[name][0].length);
  });
}

function createLedgerOutputSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("帳簿出力設定");
  if (!sheet) sheet = ss.insertSheet("帳簿出力設定");
  const headers = getSheetHeaders_()["帳簿出力設定"];
  setupSheet_(sheet, headers);
  if (sheet.getLastRow() < 2) {
    writeObjectsToSheet(sheet, [
      {"出力する": true, "帳簿名": "要確認一覧", "シート名": "要確認一覧", "出力列": "全列", "最大行数": 200, "備考": "SAMPLE"},
      {"出力する": true, "帳簿名": "日報レシート", "シート名": "日報レシート管理", "出力列": "全列", "最大行数": 200, "備考": "SAMPLE"}
    ]);
  }
}

function createDailyReportPdfForActiveRow() {
  toast_("v9.7.1クリーン版ではPDF本体生成は簡易版です。必要なら次版でPDFテンプレートを再接続します。");
}

function createLedgerPdf() {
  toast_("v9.7.1クリーン版では帳簿PDF本体生成は簡易版です。必要なら次版でPDFテンプレートを再接続します。");
}

function hideSupportSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SUPPORT_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.hideSheet();
  });
  toast_("裏方シートを非表示にしました");
}

function showSupportSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SUPPORT_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.showSheet();
  });
  toast_("裏方シートを表示しました");
}

function reorderSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SHEET_ORDER.forEach((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1);
  });
}

function createFilterSafely_(sheet, colCount) {
  removeFilter_(sheet);
  try { sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), colCount).createFilter(); } catch (e) {}
}

function removeFilter_(sheet) {
  try { const f = sheet.getFilter(); if (f) f.remove(); } catch (e) {}
}

function showAllColumns_(sheet) {
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (e) {}
}

function normalizeClearLabelForEditedCell_(e) {
  if (!e || !e.range) return;
  if (String(e.value || "") === CLEAR_LABEL) e.range.clearContent();
}

function buildLicenseTypeText_(row) {
  const items = LICENSE_TYPE_HEADERS.filter(h => isTruthy_(row[h]));
  return items.length ? items.join("・") : "免許種類未入力";
}

function isReadByAll_(row) {
  const members = getPersonalMembers_();
  if (!members.length) return false;
  return members.every(m => isTruthy_(row[m]));
}

function isTruthy_(v) {
  return v === true || v === "TRUE" || v === "Y" || v === "済";
}

function joinText(a, b) {
  const x = String(a || "").trim();
  const y = String(b || "").trim();
  if (x && y) return x + " / " + y;
  return x || y || "";
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function markSummaryDirty_() {
  PropertiesService.getDocumentProperties().setProperty(SUMMARY_DIRTY_KEY, "1");
}

function clearSummaryDirty_() {
  PropertiesService.getDocumentProperties().deleteProperty(SUMMARY_DIRTY_KEY);
}

function isSummaryDirty_() {
  return PropertiesService.getDocumentProperties().getProperty(SUMMARY_DIRTY_KEY) === "1";
}

function toast_(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message);
}







































