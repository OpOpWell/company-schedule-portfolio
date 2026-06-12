
/**
 * 社内共有業務管理システム v9.7.24 安定版戻し（完全1ファイル貼り替え用）
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
 *
 * v9.7.3:
 * - 入力済み行だけ既読・個人確認チェックボックスを表示。
 * - 空行はチェックボックスを消して、見た目を軽くする。
 * - 個人確認列を縦書き・細幅・折りたたみグループに整理。
 * - v9.6.14のカラーリング思想を、v9.7系の列名に合わせて再整理。
 *
 * v9.7.4:
 * - GPT診断用シートを追加。
 * - 実データ全文ではなく、構成・件数・ヘッダー・要注意点だけをGPTに貼れる形で出力。
 * v9.7.7:
 * - 旧版に近いサクサク運用へ戻すため、onEditと日常メニューをさらに軽量化。
 * - onEditでは編集行のID・通知・既読チェックだけ処理し、全体の表示調整やグループ再作成は行わない。
 * - チェックボックス全体再設定は最終行までに限定し、保守時だけ実行する。
 * - 空行ID整理は「今のシートだけ」を基本にし、全シート一括整理は非常用へ隔離。
 *
 * v9.7.8:
 * - 新セットアップ中に未作成シートがあっても reorderSheets_() が止まらないよう安全化。
 * - 存在するシートだけを 1,2,3... の順で移動し、moveActiveSheet の無効引数エラーを回避。
 * - 並び替え失敗時はログに残して処理を継続。
 *
 * v9.7.12:
 * - GPT診断用シートに、カラーリング/条件付き書式/表示設定/重さ診断を追加。
 * v9.7.14:
 * - カラー被りを減らすため、通知・状態・種類・電話・経理系の色を再整理。
 * - 同じ赤/黄/緑の使い回しを減らし、一覧スケジュールで種類を見分けやすくした。
 * - 各シートの条件付き書式ルール数、色付け対象列、IDだけ空行、getLastRow乖離を診断できるようにした。
 * v9.7.10:
 * - 既読列の見出しだけ横書きで残る問題を修正。
 * - applyPersonalCheckColumnLayout_() で「既読」も個人確認列と同じ縦書き・細幅に統一。
 */

const SYSTEM_VERSION = "v9.7.14-menu-slim";
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
  "GPT診断用",
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
  "設定", "SQL設計", "移行対応表", "SQLサンプル集", "帳簿PDF履歴", "帳簿出力設定", "担当別未読", "既読率集計", "ダッシュボード", "手順シート", "社内管理説明", "機能チェック", "GPT診断用"
];

/**
 * v9.7.24 注意:
 * - Apps Script 側に古い .gs ファイルが残っていると、古い onOpen() が実行されて
 *   メニューが「日常更新・集計更新・GPT診断・PDF」だけに戻ることがあります。
 * - このコードは「1ファイルだけ残して全文貼り替え」する前提です。
 * - onOpen() はこの1個だけにしてください。
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // v9.7.14 安定版:
  // メニューが長くなりすぎると見づらいため、日常操作・診断・表示調整・構成修復・重い処理に整理。
  // 車検履歴自動追加は入れず、車検更新済の空欄戻しだけ手動メニューにする。

  ui.createMenu("社内管理")
    .addItem("バージョン確認", "showSystemVersion")
    .addSeparator()

    // 日常運用
    .addItem("日常更新（一覧＋要確認）", "refreshDailyOnly")
    .addItem("集計更新（未読・既読率・DB）", "refreshSummaryOnly")
    .addSeparator()

    // 診断・帳票
    .addItem("GPT診断用シートを作成", "createGptDiagnosticSheet")
    .addItem("機能チェックシートを作成", "createFeatureCheckSheet")
    .addItem("データ整合性チェック・軽微修正", "runDataConsistencyCheckAndFix")
    .addSeparator()
    .addItem("選択行の日報PDFを作成", "createDailyReportPdfForActiveRow")
    .addItem("一覧帳簿PDFを作成", "createLedgerPdf")
    .addSeparator()

    .addSubMenu(
      ui.createMenu("保守用：表示調整")
        .addItem("全入力シートの横スクロール軽減（統一）", "compactAllInputSheets")
        .addItem("今のシートだけ横スクロール軽減（統一）", "compactActiveSheet")
        .addItem("今のシートの非表示列をすべて表示", "showAllColumnsForActiveSheet")
        .addItem("今のシートだけ表示調整", "formatActiveSheetOnly")
        .addItem("既読表示・個人確認列を再調整（全体・重い）", "applyReadDisplayToAllSheets")
        .addItem("個人確認グループを作り直す（全体・重い）", "rebuildCheckGroups")
    )
    .addSubMenu(
      ui.createMenu("保守用：個別修復")
        .addItem("行事予定に終了日列を追加/修復", "ensureV9714MinorLayoutUpdates_")
        .addItem("車検更新済を空欄に戻す", "clearUpdatedVehicleStatuses")
        .addItem("運転免許シートのズレ調整", "alignLicenseSheetLayout")
        .addItem("免許・資格・備品の横スクロール軽減（統一）", "compactLicenseQualificationEquipmentSheets")
    )
    .addSubMenu(
      ui.createMenu("保守用：ID・既読")
        .addItem("今のシートだけID・既読整理", "cleanActiveSheetOrphanIdsAndCheckboxes")
        .addItem("今のシートだけ既読チェック再設定", "refreshActiveSheetCheckboxes")
        .addItem("全入力シートのID・既読チェックを整理（非常用・重い）", "cleanOrphanIdsAndCheckboxes")
    )
    .addSubMenu(
      ui.createMenu("保守用：初回・構成修復")
        .addItem("入力シートだけ設定反映", "refreshInputSheets")
        .addItem("裏方シートを作成/修復", "ensureSupportSheets")
        .addItem("手順シートを作成/更新", "createProcedureSheet")
        .addItem("ベータ確認準備（軽量）", "prepareBetaCheckLight")
        .addItem("新規ベータ作成（危険・既存データ注意）", "setupNewBetaClean")
    )
    .addSubMenu(
      ui.createMenu("保守用：表示整理")
        .addItem("裏方シートを非表示", "hideSupportSheets")
        .addItem("裏方シートを表示", "showSupportSheets")
        .addItem("一覧スケジュールを月別グループ化", "groupScheduleByMonth")
        .addItem("過去一覧を月別グループ化", "groupArchiveByMonth")
        .addItem("今のシートを月別グループ化", "groupActiveSheetByMonth")
    )
    .addSubMenu(
      ui.createMenu("保守用：サンプル")
        .addItem("サンプル追加：主要予定系", "addSampleMainSchedules")
        .addItem("サンプル追加：車検・免許・資格・備品", "addSampleVehicleLicenseEquipment")
        .addItem("サンプル追加：日報・レシート・FB", "addSampleDailyReceiptFeedback")
        .addItem("ベータ用サンプルを削除（注意）", "deleteBetaSamples")
    )
    .addSubMenu(
      ui.createMenu("保守用：重い処理・通常使用禁止")
        .addItem("全体更新（保守用・重め）", "refreshAll")
        .addItem("書式込み全体更新（かなり重い）", "refreshAllWithFormat")
        .addItem("旧列名をv9.7.0へ移行", "migrateLegacyColumnsV970")
        .addItem("車検管理を修復（車番文字列＋色）", "repairVehicleInspectionSheet")
        .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
        .addItem("月次整理（過去移動＋月別）", "monthlyMaintenance")
    )
    .addToUi();
}

function showSystemVersion() {
  toast_("社内共有業務管理システム " + SYSTEM_VERSION + " が動いています");
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
    normalizeVehicleInspectionUpdatedRow_(sheet, row);
    updateNoticeForEditedRow_(sheet, row);
    setCheckboxesForEditedRow_(sheet, row);
    markSummaryDirty_();
  } catch (err) {
    console.log("onEdit error: " + err.message);
  }
}

function setupNewBetaClean() {
  // v9.7.9:
  // Apps Scriptエディタの実行ボタンから実行した場合、
  // SpreadsheetApp.getUi() が使えず「Cannot call SpreadsheetApp.getUi() from this context」で止まるため、
  // 新セットアップ本体から確認ダイアログは外す。
  // 実行前確認は、メニュー名・手順シート・バックアップ運用で行う。
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
  ensureV9714MinorLayoutUpdates_();
  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    if (name === "車検管理") normalizeVehicleInspectionUpdatedRows_(sheet);
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
    "行事予定": ["日付", "終了日", "開始時刻", "終了時刻", "行事名", "内容", "担当", "状態", "通知", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "行事ID"],
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
  return ["一覧スケジュール", "要確認一覧", "担当別未読", "既読率集計", "ダッシュボード", "過去一覧", "帳簿PDF履歴", "帳簿出力設定", "SQL設計", "移行対応表", "SQLサンプル集", "手順シート", "社内管理説明", "機能チェック", "GPT診断用"].includes(sheetName);
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
  applyPersonalCheckColumnLayout_(sheet);
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

/**
 * 入力済み行だけに既読・個人確認チェックボックスを出す。
 * 空行にはチェックボックスを表示しない。
 */
function setCheckboxesForDataRows(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = getHeaders_(sheet);
  const checkboxHeaders = getCheckboxHeadersForSheet_(sheet);
  if (!checkboxHeaders.length) return;

  // 旧版に近い軽さへ戻すため、全入力規則対象行ではなく「実データがある最終行まで」だけを見る。
  // 空行を何百行も掃除しない。
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  const hasDataFlags = values.map(rowValues => isBusinessDataRow_(headers, rowValues, sheet.getName()));

  checkboxHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;

    const targetRange = sheet.getRange(2, col, rowCount, 1);
    const oldValues = targetRange.getValues();

    targetRange.clearDataValidations();
    targetRange.clearContent();

    getTrueRuns_(hasDataFlags).forEach(run => {
      const startRow = run.start + 2;
      const length = run.length;
      const range = sheet.getRange(startRow, col, length, 1);
      range.insertCheckboxes();

      const newValues = oldValues
        .slice(run.start, run.start + length)
        .map(row => [isTruthy_(row[0])]);
      range.setValues(newValues);
    });
  });

  // 個人確認列の縦書き・折りたたみは重いので、ここでは毎回実行しない。
  // 必要な時だけ保守メニューの表示調整を使う。
}

function getTrueRuns_(flags) {
  const runs = [];
  let start = -1;

  flags.forEach((flag, index) => {
    if (flag && start < 0) start = index;
    if ((!flag || index === flags.length - 1) && start >= 0) {
      const end = flag && index === flags.length - 1 ? index + 1 : index;
      runs.push({ start: start, length: end - start });
      start = -1;
    }
  });

  return runs;
}

/**
 * 編集された1行だけ、入力済みならチェックボックスを出す。
 * 行の内容が消えたらチェックボックスも消す。
 */
function setCheckboxesForEditedRow_(sheet, row) {
  if (!sheet || row <= 1 || sheet.getLastColumn() < 1) return;
  const headers = getHeaders_(sheet);
  const checkboxHeaders = getCheckboxHeadersForSheet_(sheet);
  if (!checkboxHeaders.length) return;

  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const hasBusinessData = isBusinessDataRow_(headers, values, sheet.getName());

  checkboxHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;

    const cell = sheet.getRange(row, col);
    const oldValue = cell.getValue();

    if (hasBusinessData) {
      cell.insertCheckboxes();
      cell.setValue(isTruthy_(oldValue));
    } else {
      cell.clearDataValidations();
      cell.clearContent();
    }
  });

  // 旧版サクサク運用に寄せるため、onEditのたびに列幅・縦書き・グループ再設定はしない。
}

function getCheckboxHeadersForSheet_(sheet) {
  const headers = [READ_HEADER, ...getPersonalMembers_()];
  if (sheet && sheet.getName() === "運転免許管理") headers.push(...LICENSE_TYPE_HEADERS);
  return headers;
}

/**
 * 業務データが入っている行か判定する。
 * 既読・個人確認・通知・IDだけでは入力済み扱いにしない。
 */
function isBusinessDataRow_(headers, rowValues, sheetName) {
  const ignoreHeaders = new Set([
    READ_HEADER,
    CALENDAR_ID_HEADER,
    "通知",
    "備考",
    "PDFリンク",
    SHEET_ID_HEADERS[sheetName] || ""
  ]);

  getPersonalMembers_().forEach(name => ignoreHeaders.add(name));

  return headers.some((header, index) => {
    if (!header || ignoreHeaders.has(header)) return false;
    const value = rowValues[index];
    if (value === true) return true;
    if (value === false) return false;
    if (value instanceof Date) return true;
    return value !== "" && value !== null && value !== undefined;
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
  collectRowsAsObjects(ss, "行事予定", r => r["日付"], r => scheduleRow_(r["日付"], "行事", buildMultiDayEventContent_(r["行事名"], r["内容"], r["日付"], r["終了日"]), r["担当"], r["状態"], getPeriodNoticeText_(r["日付"], r["終了日"], r["状態"]), "", r, "行事予定"), rows);
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
  if (sheetName === "行事予定") return getPeriodNoticeText_(row["日付"], row["終了日"], row["状態"]);
  if (sheetName === "会議予定" || sheetName === "出先予定" || sheetName === "社用車予約" || sheetName === "電話履歴" || sheetName === "日報") return getNoticeText(row["日付"], row["状態"] || row["電話対応"]);
  return "";
}


function getPeriodNoticeText_(startValue, endValue, status) {
  const s = String(status || "").trim();
  if (["完了", "更新済", "確認済", "返却済", "精算済", "中止", "修理不可"].includes(s)) return "";

  const start = toDateOnly_(startValue);
  const end = toDateOnly_(endValue) || start;
  if (!start) return "";

  const today = toDateOnly_(new Date());

  // 期間中の行事は「今日」として要確認に出す。
  if (end && start.getTime() <= today.getTime() && today.getTime() <= end.getTime()) return "今日";

  // 終了日を過ぎた未完了行事は期限切れ。
  if (end && end.getTime() < today.getTime()) return "期限切れ";

  return getNoticeText(start, s);
}

function buildMultiDayEventContent_(title, content, startValue, endValue) {
  const titleText = joinText(title, content);
  const start = toDateOnly_(startValue);
  const end = toDateOnly_(endValue);
  if (!start || !end || start.getTime() === end.getTime()) return titleText;

  const tz = Session.getScriptTimeZone();
  const startText = Utilities.formatDate(start, tz, "M/d");
  const endText = Utilities.formatDate(end, tz, "M/d");
  return joinText(titleText, startText + "〜" + endText);
}


function clearUpdatedVehicleStatuses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("車検管理");
  if (!sheet) {
    toast_("車検管理シートが見つかりません");
    return;
  }
  normalizeVehicleInspectionUpdatedRows_(sheet);
  toast_("車検管理の更新済を確認し、未来日の更新済は空欄に戻しました");
}

function normalizeVehicleInspectionUpdatedRow_(sheet, row) {
  if (!sheet || sheet.getName() !== "車検管理" || row <= 1) return;

  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf("車検状態") + 1;
  if (statusCol <= 0) return;

  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = objectFromRow(headers, values);
  const status = String(obj["車検状態"] || "").trim();
  if (status !== "更新済") return;

  const due = toDateOnly_(obj["車検予定日"] || obj["車検期限"]);
  const today = toDateOnly_(new Date());

  // 車検期限を新しい未来日に更新した行は、次回管理に戻すため車検状態を空欄へ戻す。
  // 期限切れの「更新済」は手作業中の可能性があるため触らない。
  if (!due || due.getTime() >= today.getTime()) {
    sheet.getRange(row, statusCol).clearContent();
  }
}

function normalizeVehicleInspectionUpdatedRows_(sheet) {
  if (!sheet || sheet.getName() !== "車検管理" || sheet.getLastRow() < 2) return;
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    normalizeVehicleInspectionUpdatedRow_(sheet, row);
  }
}

function ensureV9714MinorLayoutUpdates_() {
  ensureColumnAfter_("行事予定", "日付", "終了日");
}

function ensureColumnAfter_(sheetName, afterHeader, newHeader) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = getHeaders_(sheet);
  if (headers.includes(newHeader)) return;

  const afterCol = headers.indexOf(afterHeader) + 1;
  if (afterCol <= 0) return;

  sheet.insertColumnAfter(afterCol);
  sheet.getRange(1, afterCol + 1).setValue(newHeader);

  const expectedHeaders = getSheetHeaders_()[sheetName];
  if (expectedHeaders) {
    applyDataValidationByHeaders_(sheet, expectedHeaders);
    formatSheetByName_(sheetName);
  }
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

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const ids = values.map(rowValues => {
    const currentId = rowValues[idCol - 1];
    const hasBusinessData = isBusinessDataRow_(headers, rowValues, sheet.getName());

    // 既読・個人確認・通知・備考・IDだけの空行にはIDを残さない。
    if (!hasBusinessData) return [""];
    if (currentId) return [currentId];
    return [buildRecordId_(sheet.getName())];
  });

  sheet.getRange(2, idCol, ids.length, 1).setValues(ids);
}

function ensureIdForEditedRow_(sheet, row) {
  if (!sheet || row <= 1) return;

  const idHeader = SHEET_ID_HEADERS[sheet.getName()];
  if (!idHeader) return;

  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf(idHeader) + 1;
  if (idCol <= 0) return;

  const rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const hasBusinessData = isBusinessDataRow_(headers, rowValues, sheet.getName());
  const idCell = sheet.getRange(row, idCol);

  // 本体項目が空なら、過去に誤発行されたIDも消す。
  if (!hasBusinessData) {
    idCell.clearContent();
    return;
  }

  if (!idCell.getValue()) idCell.setValue(buildRecordId_(sheet.getName()));
}

function cleanActiveSheetOrphanIdsAndCheckboxes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || isAutoOutputSheet_(sheet.getName()) || SUPPORT_SHEETS.includes(sheet.getName())) {
    toast_("入力シートを開いてから実行してください");
    return;
  }

  const result = cleanOrphanIdsAndCheckboxesForSheet_(sheet, true);
  toast_("今のシートだけ整理しました（ID整理: " + result.cleanedIds + "）");
}

function refreshActiveSheetCheckboxes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || isAutoOutputSheet_(sheet.getName()) || SUPPORT_SHEETS.includes(sheet.getName())) {
    toast_("入力シートを開いてから実行してください");
    return;
  }

  setCheckboxesForDataRows(sheet);
  toast_("今のシートだけ既読チェックを再設定しました");
}

function cleanOrphanIdsAndCheckboxes() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "全入力シートのID・既読チェック整理",
    "この処理は全入力シートを確認する非常用メニューです。通常は『今のシートだけID・既読整理』を使ってください。続行しますか？",
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中のため、空行整理をスキップしました");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheets = getInputSheetNames_();
    let cleanedIds = 0;
    let processedSheets = 0;

    targetSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const result = cleanOrphanIdsAndCheckboxesForSheet_(sheet, false);
      cleanedIds += result.cleanedIds;
      processedSheets += result.processed ? 1 : 0;
    });

    toast_("全入力シートのID・既読チェックを整理しました（ID整理: " + cleanedIds + " / 対象: " + processedSheets + "）");
  } finally {
    lock.releaseLock();
  }
}

function cleanOrphanIdsAndCheckboxesForSheet_(sheet, refreshCheckboxes) {
  const result = { cleanedIds: 0, processed: false };
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return result;

  const sheetName = sheet.getName();
  const headers = getHeaders_(sheet);
  if (!headers.length) return result;

  const idHeader = SHEET_ID_HEADERS[sheetName];
  const idCol = idHeader ? headers.indexOf(idHeader) + 1 : 0;
  const lastRow = sheet.getLastRow();
  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  const hasDataFlags = values.map(rowValues => isBusinessDataRow_(headers, rowValues, sheetName));

  if (idCol > 0) {
    const currentIds = sheet.getRange(2, idCol, rowCount, 1).getValues();
    let changed = false;
    const nextIds = currentIds.map((row, index) => {
      const currentId = row[0];

      if (!hasDataFlags[index]) {
        if (currentId) {
          result.cleanedIds++;
          changed = true;
        }
        return [""];
      }

      if (currentId) return [currentId];
      changed = true;
      return [buildRecordId_(sheetName)];
    });

    if (changed) sheet.getRange(2, idCol, rowCount, 1).setValues(nextIds);
  }

  if (refreshCheckboxes) setCheckboxesForDataRows(sheet);
  result.processed = true;
  return result;
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
    {"日付": addDays_(today, 3), "終了日": addDays_(today, 5), "開始時刻": "09:00", "終了時刻": "12:00", "行事名": "サンプル安全講習", "内容": "安全教育", "担当": "鈴木", "状態": "予定", "備考": "SAMPLE"}
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


/**
 * GPTに貼りやすい診断用シートを作成する。
 * 実データを丸ごと出さず、シート構成・ヘッダー・件数・要注意点だけを一覧化する。
 */
function createGptDiagnosticSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "GPT診断用";
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  sheet.clearFormats();
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); } catch (e) {}
  sheet.setConditionalFormatRules([]);

  const rows = [["区分", "対象", "確認内容", "結果", "詳細", "GPTに見てほしい点"]];
  const nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  rows.push(["貼付用", "GPTへの依頼文", "このシートの内容をChatGPTへ貼る", "案内", buildGptDiagnosticPrompt_(), "この行と下の診断表を貼る"]);
  rows.push(["システム", "バージョン", "SYSTEM_VERSION", "情報", SYSTEM_VERSION, "v9.7系の最新版か確認"]);
  rows.push(["システム", "作成日時", "診断作成日時", "情報", nowText, "診断時点の状態"]);
  rows.push(["システム", "担当者", "設定シートから読み込み", "情報", getStaffMembers_().join(" / "), "実名なら社外共有前に匿名化"]);
  rows.push(["システム", "既読確認者", "個人確認列", "情報", getPersonalMembers_().join(" / "), "個人確認列の数と表示幅"]);

  appendGptMenuSpecDiagnostics_(rows);
  appendGptSheetStructureDiagnostics_(ss, rows);
  appendGptCompactLayoutDiagnostics_(ss, rows);
  appendGptAllCompactLayoutDiagnostics_(ss, rows);
  appendGptLegacyColumnDiagnostics_(ss, rows);
  appendGptScheduleDiagnostics_(ss, rows);
  appendGptAlertDiagnostics_(ss, rows);
  appendGptCheckboxDiagnostics_(ss, rows);
  appendGptPerformanceDiagnostics_(ss, rows);
  appendGptColorDiagnostics_(ss, rows);
  appendGptVehicleNumberDiagnostics_(ss, rows);
  appendGptAppSheetDiagnostics_(ss, rows);
  appendGptOperationAdvice_(rows);

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.getRange(1, 1, 1, rows[0].length)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 230);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 520);
  sheet.setColumnWidth(6, 420);
  sheet.getRange(1, 1, rows.length, rows[0].length).setWrap(true).setVerticalAlignment("middle");
  applyStatusBackground_(sheet, 4, 2, rows.length - 1);

  try { sheet.getFilter() && sheet.getFilter().remove(); } catch (e) {}
  try { sheet.getRange(1, 1, Math.max(rows.length, 2), rows[0].length).createFilter(); } catch (e) {}

  toast_("GPT診断用シートを作成しました。必要な範囲をコピーしてGPTに貼ってください。");
}

function buildGptDiagnosticPrompt_() {
  return [
    "社内共有業務管理システム v9.7系のGoogle Sheets + Apps Script + AppSheet連携を確認しています。",
    "下の診断表を見て、シート構成、ヘッダー、一覧反映、要確認一覧、既読チェックボックス、カラーリング/条件付き書式、表示設定、getLastRow/IDだけ空行、車番日付化、AppSheet設定、メニュー構成、保守用メニューの隔離、車検自動化を入れていない仕様で問題がありそうな点を指摘してください。",
    "実データは社外共有できないため、件数・ヘッダー・診断結果だけで判断してください。",
    "回答は、重大度順に『今すぐ直す』『ベータ前に確認』『後回しでよい』に分けてください。"
  ].join("\n");
}


function appendGptMenuSpecDiagnostics_(rows) {
  rows.push([
    "メニュー仕様",
    "メニュー構成",
    "社内管理の単一メニュー統合版",
    "OK",
    "2つ目のメニューが表示されない環境対策として、保守用メニューを社内管理の中に統合しています。",
    "社内管理の中に保守用サブメニューがまとまっているか確認"
  ]);

  rows.push([
    "メニュー仕様",
    "2メニュー構成",
    "社内管理・保守用（重い）が別メニューに出ない場合",
    "OK",
    "この版では別メニューが出なくても異常ではありません。社内管理内の保守用サブメニューに統合しています。",
    "別メニューが出ないこと自体をNG扱いしない"
  ]);

  rows.push([
    "メニュー仕様",
    "通常運用メニュー",
    "普段使う操作",
    "OK",
    "日常更新（一覧＋要確認） / 集計更新（未読・既読率・DB） / GPT診断用シートを作成 / 日報PDF / 一覧帳簿PDF をトップ付近に配置しています。",
    "日常操作と保守操作が混ざりすぎていないか確認"
  ]);

  rows.push([
    "メニュー仕様",
    "保守用サブメニュー",
    "保守用：初回・構成修復 / 表示・既読・ID修復 / サンプル / 重い処理・通常使用禁止 / 表示整理",
    "OK",
    "危険・重い処理はサブメニューへ隔離しています。",
    "新規ベータ作成や全体更新を日常操作で押しにくい配置か確認"
  ]);

  rows.push([
    "メニュー仕様",
    "通常使用禁止メニュー",
    "新規ベータ作成・全体更新・書式込み全体更新・旧列名移行・車検管理修復",
    "要確認",
    "これらは保守用サブメニュー内にある前提です。GASから実際のメニュー表示位置は自動取得できないため、画面上で目視確認してください。",
    "トップ直下に危険メニューが出ていないか確認"
  ]);

  rows.push([
    "メニュー仕様",
    "onOpen重複",
    "function onOpen が1個だけか",
    "要目視",
    "GASの実行中コードからプロジェクト内のonOpen個数は安全に自動判定できません。Apps Script検索で function onOpen が1件だけか確認してください。",
    "古い軽量版onOpenが別ファイルや同一ファイル下部に残っていないか確認"
  ]);

  rows.push([
    "車検仕様",
    "車検履歴自動追加",
    "現時点では入れない",
    "OK",
    "安定版戻し優先のため、車検期限変更時の履歴自動追加は未実装のままにしています。",
    "履歴自動化を今入れるべきか、後回しでよいか確認"
  ]);

  rows.push([
    "車検仕様",
    "車検状態更新時の自動クリア",
    "未来日に更新済なら状態だけ空欄へ戻す",
    "OK",
    "車検期限または車検予定日が今日以降で、車検状態が更新済の行は、次回管理に戻すため車検状態だけ空欄へ戻します。期限・予定日・履歴は自動削除しません。",
    "車検状態だけ空欄へ戻る仕様で問題ないか確認"
  ]);

  rows.push([
    "車検仕様",
    "車検管理の自動ソート",
    "現時点では後回し",
    "OK",
    "車検期限を変更しても車検管理シート自体を即時自動ソートする処理は、安定版戻しでは入れていません。一覧スケジュール側は日常更新でソートします。",
    "車検管理本体の自動ソートが必要か、手動/保守メニューで十分か確認"
  ]);

  const menuFunctions = [
    ["バージョン確認", "showSystemVersion", "トップ"],
    ["日常更新（一覧＋要確認）", "refreshDailyOnly", "トップ"],
    ["集計更新（未読・既読率・DB）", "refreshSummaryOnly", "トップ"],
    ["GPT診断用シートを作成", "createGptDiagnosticSheet", "トップ"],
    ["機能チェックシートを作成", "createFeatureCheckSheet", "トップ"],
    ["データ整合性チェック・軽微修正", "runDataConsistencyCheckAndFix", "トップ"],
    ["選択行の日報PDFを作成", "createDailyReportPdfForActiveRow", "トップ"],
    ["一覧帳簿PDFを作成", "createLedgerPdf", "トップ"],
    ["入力シートだけ設定反映", "refreshInputSheets", "保守用：初回・構成修復"],
    ["裏方シートを作成/修復", "ensureSupportSheets", "保守用：初回・構成修復"],
    ["手順シートを作成/更新", "createProcedureSheet", "保守用：初回・構成修復"],
    ["ベータ確認準備（軽量）", "prepareBetaCheckLight", "保守用：初回・構成修復"],
    ["新規ベータ作成（危険・既存データ注意）", "setupNewBetaClean", "保守用：初回・構成修復"],
    ["全入力シートの横スクロール軽減（統一）", "compactAllInputSheets", "保守用：表示・既読・ID修復"],
    ["免許・資格・備品の横スクロール軽減（統一）", "compactLicenseQualificationEquipmentSheets", "保守用：表示・既読・ID修復"],
    ["運転免許シートのズレ調整", "alignLicenseSheetLayout", "保守用：表示・既読・ID修復"],
    ["車検更新済を空欄に戻す", "clearUpdatedVehicleStatuses", "保守用：表示・既読・ID修復"],
    ["行事予定に終了日列を追加/修復", "ensureV9714MinorLayoutUpdates_", "保守用：表示・既読・ID修復"],
    ["今のシートだけ横スクロール軽減（統一）", "compactActiveSheet", "保守用：表示・既読・ID修復"],
    ["今のシートの非表示列をすべて表示", "showAllColumnsForActiveSheet", "保守用：表示・既読・ID修復"],
    ["今のシートだけ表示調整", "formatActiveSheetOnly", "保守用：表示・既読・ID修復"],
    ["今のシートだけID・既読整理", "cleanActiveSheetOrphanIdsAndCheckboxes", "保守用：表示・既読・ID修復"],
    ["今のシートだけ既読チェック再設定", "refreshActiveSheetCheckboxes", "保守用：表示・既読・ID修復"],
    ["表示幅・カラーリングを調整（入力シート・重め）", "formatInputSheetsForLongText", "保守用：表示・既読・ID修復"],
    ["既読表示・個人確認列を再調整（全体・重い）", "applyReadDisplayToAllSheets", "保守用：表示・既読・ID修復"],
    ["個人確認グループを作り直す（全体・重い）", "rebuildCheckGroups", "保守用：表示・既読・ID修復"],
    ["全入力シートのID・既読チェックを整理（非常用・重い）", "cleanOrphanIdsAndCheckboxes", "保守用：表示・既読・ID修復"],
    ["サンプル追加：主要予定系", "addSampleMainSchedules", "保守用：サンプル"],
    ["サンプル追加：車検・免許・資格・備品", "addSampleVehicleLicenseEquipment", "保守用：サンプル"],
    ["サンプル追加：日報・レシート・FB", "addSampleDailyReceiptFeedback", "保守用：サンプル"],
    ["ベータ用サンプルを削除（注意）", "deleteBetaSamples", "保守用：サンプル"],
    ["全体更新（保守用・重め）", "refreshAll", "保守用：重い処理・通常使用禁止"],
    ["書式込み全体更新（かなり重い）", "refreshAllWithFormat", "保守用：重い処理・通常使用禁止"],
    ["旧列名をv9.7.0へ移行", "migrateLegacyColumnsV970", "保守用：重い処理・通常使用禁止"],
    ["車検管理を修復（車番文字列＋色）", "repairVehicleInspectionSheet", "保守用：重い処理・通常使用禁止"],
    ["過去予定を過去一覧へ移動", "movePastItemsToArchive", "保守用：重い処理・通常使用禁止"],
    ["月次整理（過去移動＋月別）", "monthlyMaintenance", "保守用：重い処理・通常使用禁止"],
    ["裏方シートを非表示", "hideSupportSheets", "保守用：表示整理"],
    ["裏方シートを表示", "showSupportSheets", "保守用：表示整理"],
    ["一覧スケジュールを月別グループ化", "groupScheduleByMonth", "保守用：表示整理"],
    ["過去一覧を月別グループ化", "groupArchiveByMonth", "保守用：表示整理"],
    ["今のシートを月別グループ化", "groupActiveSheetByMonth", "保守用：表示整理"]
  ];

  menuFunctions.forEach(item => {
    const label = item[0];
    const fn = item[1];
    const place = item[2];
    const exists = isCallableFunctionByName_(fn);
    rows.push([
      "メニュー関数",
      label,
      fn + " / " + place,
      exists ? "OK" : "NG",
      exists ? "メニューから呼び出す関数が存在します。" : "メニューから呼び出す関数が見つかりません。",
      exists ? "配置と名称が意図通りか確認" : "メニューを押すとエラーになるため、関数名またはメニュー定義を修正"
    ]);
  });
}

function isCallableFunctionByName_(name) {
  try {
    return typeof eval(name) === "function";
  } catch (e) {
    return false;
  }
}


function appendGptSheetStructureDiagnostics_(ss, rows) {
  const headersMap = getSheetHeaders_();
  Object.keys(headersMap).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      rows.push(["シート構成", name, "シート存在", "NG", "シートがありません", "作成漏れか、名前変更の可能性"]);
      return;
    }

    const expected = headersMap[name];
    const actual = getHeaders_(sheet);
    const missing = expected.filter(h => !actual.includes(h));
    const extras = actual.filter(h => h && !expected.includes(h));
    const duplicated = findDuplicatedValues_(actual);
    const dataCount = Math.max(sheet.getLastRow() - 1, 0);

    rows.push(["シート構成", name, "存在・件数", "OK", "データ行数: " + dataCount + " / 列数: " + actual.length, "極端に多い場合は動作が重くなる可能性"]);
    rows.push(["ヘッダー", name, "必須ヘッダー", missing.length ? "NG" : "OK", missing.length ? "不足: " + missing.join(", ") : "必須ヘッダーOK", "列名変更後はAppSheet Regenerate Structureが必要"]);

    if (extras.length) {
      rows.push(["ヘッダー", name, "想定外ヘッダー", "確認", extras.join(", "), "旧列名や手入力列が混ざっていないか確認"]);
    }
    if (duplicated.length) {
      rows.push(["ヘッダー", name, "重複ヘッダー", "NG", duplicated.join(", "), "AppSheetや集計で不具合になりやすい"]);
    }
  });
}

function appendGptLegacyColumnDiagnostics_(ss, rows) {
  const legacyMap = {
    "車検管理": ["次回車検期限", "次回保険期限", "状態"],
    "電話履歴": ["日時", "メモ"],
    "備品修理管理": ["登録日", "場所", "担当"],
    "免許資格管理": ["担当", "資格名", "区分"]
  };

  Object.keys(legacyMap).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      if (name === "免許資格管理") rows.push(["旧仕様", name, "旧シート", "OK", "旧シートなし", "v9.7系では運転免許管理・資格管理に分離"]);
      return;
    }
    const headers = getHeaders_(sheet);
    const found = legacyMap[name].filter(h => headers.includes(h));
    const isOldLicenseSheet = name === "免許資格管理" && sheet.getLastRow() > 1;
    if (found.length || isOldLicenseSheet) {
      rows.push(["旧仕様", name, "旧列名・旧シート混在", "確認", found.join(", ") || "旧シートにデータあり", "v9.7系へ移行済みなら残さない方が安全"]);
    } else {
      rows.push(["旧仕様", name, "旧列名・旧シート混在", "OK", "旧仕様の混在なし", ""]);
    }
  });
}

function appendGptScheduleDiagnostics_(ss, rows) {
  const schedule = ss.getSheetByName("一覧スケジュール");
  if (!schedule || schedule.getLastRow() < 2) {
    rows.push(["一覧", "一覧スケジュール", "データ存在", "確認", "一覧スケジュールにデータがありません", "日常更新またはサンプル追加後に再確認"]);
    return;
  }

  const headers = getHeaders_(schedule);
  const typeCol = headers.indexOf("種類") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  if (typeCol <= 0) return;

  const values = schedule.getRange(2, 1, schedule.getLastRow() - 1, headers.length).getValues();
  const typeCounts = countByIndex_(values, typeCol - 1);
  const requiredTypes = ["車検", "運転免許", "資格", "備品修理", "工事予定", "電話履歴"];
  requiredTypes.forEach(type => {
    const count = typeCounts[type] || 0;
    rows.push(["一覧", type, "一覧反映", count > 0 ? "OK" : "確認", "件数: " + count, count > 0 ? "" : "サンプルがない/更新前/抽出条件違いの可能性"]);
  });

  rows.push(["一覧", "種類別件数", "一覧スケジュール内訳", "情報", objectToText_(typeCounts), "種類の偏りや不要な表示がないか確認"]);
  if (noticeCol > 0) rows.push(["一覧", "通知別件数", "通知内訳", "情報", objectToText_(countByIndex_(values, noticeCol - 1)), "期限切れ/今日/3日以内が妥当か確認"]);
}

function appendGptAlertDiagnostics_(ss, rows) {
  const sheet = ss.getSheetByName("要確認一覧");
  if (!sheet || sheet.getLastRow() < 2) {
    rows.push(["要確認", "要確認一覧", "データ存在", "確認", "要確認一覧にデータがありません", "期限系サンプルがない場合は正常。サンプルありなら更新確認"]);
    return;
  }

  const headers = getHeaders_(sheet);
  const noticeCol = headers.indexOf("通知") + 1;
  const typeCol = headers.indexOf("種類") + 1;
  const statusCol = headers.indexOf("状態") + 1;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  if (noticeCol > 0) rows.push(["要確認", "通知別件数", "通知内訳", "情報", objectToText_(countByIndex_(values, noticeCol - 1)), "期限切れ/今日/3日以内/差戻し/未精算の出方を確認"]);
  if (typeCol > 0) rows.push(["要確認", "種類別件数", "種類内訳", "情報", objectToText_(countByIndex_(values, typeCol - 1)), "要確認に出すべき種類が出ているか確認"]);
  if (statusCol > 0) rows.push(["要確認", "状態別件数", "状態内訳", "情報", objectToText_(countByIndex_(values, statusCol - 1)), "完了済みが残っていないか確認"]);
}

function appendGptCheckboxDiagnostics_(ss, rows) {
  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;
    const headers = getHeaders_(sheet);
    const checkboxHeaders = getCheckboxHeadersForSheet_(sheet);
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const checkEndRow = Math.min(sheet.getMaxRows(), lastRow + 30);
    const rowCount = Math.max(checkEndRow - 1, 1);
    const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();

    let blankRowsWithCheckbox = 0;
    let dataRowsWithoutCheckbox = 0;
    checkboxHeaders.forEach(header => {
      const col = headers.indexOf(header) + 1;
      if (col <= 0) return;
      const validations = sheet.getRange(2, col, rowCount, 1).getDataValidations();
      values.forEach((rowValues, i) => {
        const hasData = isBusinessDataRow_(headers, rowValues, name);
        const hasCheckbox = validations[i] && validations[i][0] && validations[i][0].getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX;
        if (!hasData && hasCheckbox) blankRowsWithCheckbox++;
        if (hasData && !hasCheckbox) dataRowsWithoutCheckbox++;
      });
    });

    const result = blankRowsWithCheckbox || dataRowsWithoutCheckbox ? "確認" : "OK";
    rows.push(["既読", name, "入力済み行だけチェックボックス", result, "空行チェック残り: " + blankRowsWithCheckbox + " / 入力済み行チェック不足: " + dataRowsWithoutCheckbox, "空行にチェックが並ぶ場合は『既読表示・個人確認列を再調整』を実行"]);
  });
}


/**
 * 処理が重くなる原因をGPT診断用に出す。
 * getLastRowが大きいのに業務データが少ない、IDだけ残った空行がある、などを確認する。
 */
function appendGptPerformanceDiagnostics_(ss, rows) {
  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;

    const headers = getHeaders_(sheet);
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const maxRows = sheet.getMaxRows();
    const rowCount = Math.max(lastRow - 1, 0);
    const idHeader = SHEET_ID_HEADERS[name] || "";
    const idCol = idHeader ? headers.indexOf(idHeader) + 1 : 0;

    let businessRows = 0;
    let blankRowsInsideLastRow = 0;
    let orphanIds = 0;
    let noticeOnlyRows = 0;
    let readOnlyRows = 0;

    if (rowCount > 0) {
      const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
      const noticeCol = headers.indexOf("通知") + 1;
      const readCol = headers.indexOf(READ_HEADER) + 1;
      const personalCols = getPersonalMembers_().map(n => headers.indexOf(n) + 1).filter(c => c > 0);

      values.forEach(rowValues => {
        const hasBusinessData = isBusinessDataRow_(headers, rowValues, name);
        if (hasBusinessData) {
          businessRows++;
          return;
        }

        blankRowsInsideLastRow++;
        if (idCol > 0 && rowValues[idCol - 1]) orphanIds++;
        if (noticeCol > 0 && rowValues[noticeCol - 1]) noticeOnlyRows++;

        const hasReadValue = readCol > 0 && (rowValues[readCol - 1] === true || rowValues[readCol - 1] === false);
        const hasPersonalValue = personalCols.some(c => rowValues[c - 1] === true || rowValues[c - 1] === false);
        if (hasReadValue || hasPersonalValue) readOnlyRows++;
      });
    }

    const lastRowGap = Math.max(rowCount - businessRows, 0);
    let result = "OK";
    if (orphanIds > 0) result = "確認";
    if (lastRowGap >= 50 || blankRowsInsideLastRow >= 50) result = "確認";
    if (lastRowGap >= 200 || orphanIds >= 20) result = "NG";

    rows.push([
      "重さ診断",
      name,
      "getLastRow/空行ID",
      result,
      "getLastRow: " + lastRow + " / 最大行: " + maxRows + " / 業務データ行: " + businessRows + " / lastRow内空行: " + blankRowsInsideLastRow + " / IDだけ空行: " + orphanIds + " / 通知だけ空行: " + noticeOnlyRows + " / 既読だけ空行: " + readOnlyRows,
      orphanIds > 0 ? "『今のシートだけID・既読整理』で先に整理。全シート一括は非常用" : "極端な乖離がなければOK"
    ]);
  });
}

/**
 * カラーリング/条件付き書式の状態をGPT診断用に出す。
 * 色そのものを自動判定するのではなく、色を付けるべき列に条件付き書式が入っているかを確認する。
 */
function appendGptColorDiagnostics_(ss, rows) {
  appendGptColorLegend_(rows);

  const targetSheetNames = Array.from(new Set([
    ...getInputSheetNames_(),
    "一覧スケジュール",
    "要確認一覧",
    "過去一覧",
    "日報レシート管理",
    "お知らせ",
    "個人ToDo",
    "フィードバック"
  ]));

  targetSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;

    const headers = getHeaders_(sheet);
    const expectedTargets = getColorDiagnosticTargetHeadersForSheet_(name).filter(h => headers.includes(h));
    const rules = sheet.getConditionalFormatRules() || [];
    const coveredHeaders = getConditionalFormatRuleTargetHeaders_(sheet, headers, rules);
    const missingTargets = expectedTargets.filter(h => !coveredHeaders.includes(h));
    const ruleCount = rules.length;

    let result = "OK";
    if (expectedTargets.length && ruleCount === 0) result = "NG";
    else if (missingTargets.length) result = "確認";
    else if (ruleCount > 90) result = "確認";

    rows.push([
      "カラーリング",
      name,
      "条件付き書式対象列",
      result,
      "ルール数: " + ruleCount + " / 期待列: " + (expectedTargets.join(", ") || "なし") + " / 実際に色ルールがある列: " + (coveredHeaders.join(", ") || "なし"),
      missingTargets.length ? "色付け対象なのに条件付き書式範囲がない列: " + missingTargets.join(", ") : "スクショで色味の見た目も確認"
    ]);
  });

  rows.push(["表示", "既読・個人確認列", "見出し縦書き", "案内", "既読、個人確認者列は setVerticalText(true) で縦書き設定", "既読だけ横向きなら『今のシートだけ表示調整』を実行"]);
}

function appendGptColorLegend_(rows) {
  const legends = [
    ["通知", "期限切れ", "淡い赤 #f4cccc", "期限超過を最優先で目立たせる"],
    ["通知", "今日", "淡い橙 #fce5cd", "当日対応を赤と淡い黄の中間で表示"],
    ["通知", "3日以内", "淡い黄 #fff2cc", "近い期限"],
    ["通知", "7日以内", "淡い緑 #d9ead3", "少し余裕あり"],
    ["通知", "30日以内", "淡い青灰 #d9eaf7", "遠めの期限"],
    ["状態", "未対応/未依頼/要確認/中止/失効", "赤〜赤紫系", "対応が必要なもの"],
    ["状態", "対応中/施工中/修理中/移動中", "青〜青淡い緑系", "進行中のもの"],
    ["状態", "完了/確認済/更新済/返却済/精算済", "濃い緑〜青淡い緑系", "完了済み"],
    ["電話", "未対応/対応中/折返し/完了", "赤/黄橙/深橙/濃緑", "電話対応の状態を独立して見分ける"],
    ["一覧", "種類", "種類別パステル色", "車検・免許・資格・備品・工事・電話などの色被りを減らす"]
  ];

  legends.forEach(r => {
    rows.push(["カラー基準", r[0], r[1], "情報", r[2], r[3]]);
  });
}

function getColorDiagnosticTargetHeadersForSheet_(sheetName) {
  const common = ["通知", "状態", "車検状態", "対応状況", "電話対応", READ_HEADER];
  const map = {
    "一覧スケジュール": ["種類", "通知", "状態", "電話対応", READ_HEADER],
    "要確認一覧": ["種類", "通知", "状態", "電話対応"],
    "日報レシート管理": ["通知", "経理確認", "精算状態"],
    "お知らせ": ["重要度", "通知", READ_HEADER],
    "フィードバック": ["対応状況"],
    "車検管理": ["通知", "車検状態", READ_HEADER],
    "運転免許管理": ["通知", "状態", READ_HEADER],
    "資格管理": ["通知", "状態", READ_HEADER],
    "備品修理管理": ["通知", "状態", READ_HEADER],
    "電話履歴": ["通知", "電話対応", READ_HEADER],
    "日報": ["状態", READ_HEADER],
    "個人ToDo": ["通知", "状態", READ_HEADER]
  };
  return map[sheetName] || common;
}

function getConditionalFormatRuleTargetHeaders_(sheet, headers, rules) {
  const covered = {};
  rules.forEach(rule => {
    const ranges = rule.getRanges ? rule.getRanges() : [];
    ranges.forEach(range => {
      if (range.getSheet().getName() !== sheet.getName()) return;
      const startCol = range.getColumn();
      const endCol = startCol + range.getNumColumns() - 1;
      for (let col = startCol; col <= endCol; col++) {
        const header = headers[col - 1];
        if (header) covered[header] = true;
      }
    });
  });
  return Object.keys(covered).sort((a, b) => String(a).localeCompare(String(b), "ja"));
}

function appendGptVehicleNumberDiagnostics_(ss, rows) {
  const sheet = ss.getSheetByName("車検管理");
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const col = headers.indexOf("車番") + 1;
  if (col <= 0) {
    rows.push(["車検", "車検管理", "車番列", "NG", "車番列がありません", "ヘッダーを確認"]);
    return;
  }

  const values = sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().flat();
  const bad = [];
  values.forEach((v, i) => {
    if (!v) return;
    if (v instanceof Date) bad.push("行" + (i + 2) + ": 日付型");
    const text = String(v);
    if (/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/.test(text)) bad.push("行" + (i + 2) + ": " + text);
  });

  rows.push(["車検", "車番", "日付化チェック", bad.length ? "NG" : "OK", bad.length ? bad.join(" / ") : "日付化なし", "車番列は文字列形式で固定"]);
}

function appendGptAppSheetDiagnostics_(ss, rows) {
  const targets = ["出先予定", "工事予定", "電話履歴", "車検管理", "運転免許管理", "資格管理", "備品修理管理", "日報", "日報レシート管理", "一覧スケジュール", "要確認一覧", "フィードバック"];
  targets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    const headers = getHeaders_(sheet);
    const idHeader = SHEET_ID_HEADERS[name] || "";
    const idOk = !idHeader || headers.includes(idHeader);
    rows.push(["AppSheet", name, "Key候補", idOk ? "OK" : "確認", idHeader ? idHeader : "自動出力/IDなし", idOk ? "AppSheet側でKey列・型を確認" : "ID列がないためKey設定を確認"]);
  });

  rows.push(["AppSheet", "時間列", "Time型候補", "情報", "行事予定: 開始時刻/終了時刻、電話履歴: 時間、社用車予約: 開始時刻/終了時刻", "AppSheet Regenerate後にTime型を確認"]);
  rows.push(["AppSheet", "チェック列", "Yes/No型候補", "情報", "既読、個人確認列、運転免許管理の免許種別", "AppSheet側でYes/No型を確認"]);
  rows.push(["AppSheet", "写真/URL列", "Image/URL型候補", "情報", "写真、レシート写真、PDFリンク", "AppSheet側でImage/URL型を確認"]);
}

function appendGptOperationAdvice_(rows) {
  rows.push(["運用", "日常更新", "普段使うメニュー", "案内", "社内管理 → 日常更新（一覧＋要確認）", "普段は重い全体更新を毎回押さない"]);
  rows.push(["運用", "保守更新", "機能チェック前", "案内", "全体更新（保守用・重め）→ 機能チェックシートを作成 → GPT診断用シートを作成", "大きな列変更後だけ実行"]);
  rows.push(["運用", "社外共有", "匿名化", "注意", "実名・会社名・電話番号・車番・取引先名・金額の扱いに注意", "GPTへ貼る前に診断表だけにする"]);
}

function findDuplicatedValues_(values) {
  const seen = {};
  const dup = {};
  values.forEach(v => {
    const key = String(v || "").trim();
    if (!key) return;
    if (seen[key]) dup[key] = true;
    seen[key] = true;
  });
  return Object.keys(dup);
}

function countByIndex_(values, index) {
  const counts = {};
  values.forEach(row => {
    const key = String(row[index] || "空欄").trim() || "空欄";
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function objectToText_(obj) {
  const keys = Object.keys(obj || {});
  if (!keys.length) return "なし";
  keys.sort((a, b) => String(a).localeCompare(String(b), "ja"));
  return keys.map(k => k + ": " + obj[k]).join(" / ");
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




function showAllColumnsForActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return;
  showAllColumns_(sheet);
  toast_("今のシートの非表示列をすべて表示しました: " + sheet.getName());
}

function compactAllInputSheets() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "全入力シートの横スクロール軽減（統一）",
    "全入力シートを、今のシートだけ横スクロール軽減と同じ基準で調整します。ID列・カレンダーID列などは非表示にしますが、データや列は削除しません。続行しますか？",
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureV9714MinorLayoutUpdates_();
  const targetNames = getInputSheetNames_();
  let count = 0;
  const skipped = [];

  targetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      skipped.push(name);
      return;
    }
    applyCompactLayoutForSheet_(sheet);
    count++;
  });

  toast_("全入力シートの横スクロール軽減（統一）を実行しました（" + count + "シート）" + (skipped.length ? " / 未作成: " + skipped.join(", ") : ""));
}

function compactActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || isAutoOutputSheet_(sheet.getName()) || SUPPORT_SHEETS.includes(sheet.getName())) {
    toast_("入力シートを開いてから実行してください");
    return;
  }
  ensureV9714MinorLayoutUpdates_();
  applyCompactLayoutForSheet_(sheet);
  toast_("今のシートの横スクロール軽減を実行しました: " + sheet.getName());
}

function applyCompactLayoutForSheet_(sheet) {
  const name = sheet.getName();

  // 免許・資格・備品は専用の圧縮レイアウトを使用。
  if (["運転免許管理", "資格管理", "備品修理管理"].includes(name)) {
    applyCompactLicenseEquipmentLayout_(sheet);
    return;
  }

  applyCompactGeneralSheetLayout_(sheet);
}

function applyCompactGeneralSheetLayout_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const sheetName = sheet.getName();
  const headers = getHeaders_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = headers.length;
  const personal = getPersonalMembers_();

  // 横スクロールを減らすが、固定線が邪魔にならないよう固定列は解除。
  sheet.setFrozenColumns(0);
  showAllColumns_(sheet);

  headers.forEach((header, index) => {
    const col = index + 1;
    if (!header) return;

    const width = getCompactColumnWidth_(sheetName, header, personal);
    sheet.setColumnWidth(col, width);

    // 縦書きは既読・個人確認だけ。
    // 状態・通知・担当まで縦書きにすると値が切れて見づらいため横書きに戻す。
    const shouldVertical = header === READ_HEADER || personal.includes(header);

    if (shouldVertical) {
      sheet.getRange(1, col, lastRow, 1)
        .setVerticalText(true)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
    } else {
      sheet.getRange(1, col, lastRow, 1)
        .setVerticalText(false)
        .setVerticalAlignment("middle")
        .setWrap(true);
    }

    if (DATE_HEADERS.includes(header) || TIME_HEADERS.includes(header) || [
      "通知", "状態", "車検状態", "電話対応", "経理確認", "精算状態", "重要度", "対応状況", "区分", "コピー有無", "返却済"
    ].includes(header) || shouldVertical) {
      sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1).setHorizontalAlignment("center");
    }
  });

  const idHeader = SHEET_ID_HEADERS[sheetName];
  if (idHeader) hideColumnByHeaderSafely_(sheet, idHeader);
  hideColumnByHeaderSafely_(sheet, CALENDAR_ID_HEADER);

  ["写真", "レシート写真", "PDFリンク"].forEach(header => hideColumnByHeaderSafely_(sheet, header));

  rebuildPersonalCheckGroupForSheet_(sheet);

  const maxRows = Math.min(sheet.getMaxRows(), Math.max(lastRow, 20));
  sheet.setRowHeight(1, 54);
  sheet.setRowHeights(2, Math.max(maxRows - 1, 1), 32);
  sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  applyColorRules(sheet);
  createFilterSafely_(sheet, lastCol);
}

function getCompactColumnWidth_(sheetName, header, personal) {
  if (personal.includes(header)) return 38;
  if (header === READ_HEADER) return 38;

  if (DATE_HEADERS.includes(header)) return 104;
  if (TIME_HEADERS.includes(header)) return 70;

  if (header === "通知") return 96;
  if (["状態", "車検状態", "電話対応", "対応状況"].includes(header)) return 88;
  if (["経理確認", "精算状態"].includes(header)) return 88;
  if (["重要度", "区分", "コピー有無", "返却済"].includes(header)) return 72;

  if (["担当", "入力者", "利用者", "投稿者", "記入者", "所有者", "修理依頼者"].includes(header)) return 92;
  if (["社用車", "車両名", "車番"].includes(header)) return 96;

  if (["行き先", "現場", "相手", "支払先", "備品名", "資格名", "タイトル", "依頼主"].includes(header)) return 130;
  if (["工事名", "会議名", "行事名"].includes(header)) return 150;

  if (["用件", "内容", "作業内容", "日報文章", "問題点", "明日の予定", "他現場状況", "対応メモ", "気になった内容", "対応方針"].includes(header)) return 190;

  if (["連絡先", "契約金額", "支払方法", "金額"].includes(header)) return 96;
  if (["備考", "メモ"].includes(header)) return 170;

  if (header === CALENDAR_ID_HEADER || String(header).endsWith("ID") || header === "ToDo_ID") return 40;
  if (["写真", "レシート写真", "PDFリンク"].includes(header)) return 40;

  return 92;
}



function alignLicenseSheetLayout() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("運転免許管理");
  if (!sheet) {
    toast_("運転免許管理シートが見つかりません");
    return;
  }
  applyCompactLicenseEquipmentLayout_(sheet);
  toast_("運転免許シートのズレ調整を実行しました");
}

function compactLicenseQualificationEquipmentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetNames = ["運転免許管理", "資格管理", "備品修理管理"];
  const missing = [];

  targetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      missing.push(name);
      return;
    }
    applyCompactLicenseEquipmentLayout_(sheet);
  });

  if (missing.length) {
    toast_("横スクロール軽減を実行しました。一部シートなし: " + missing.join(", "));
  } else {
    toast_("免許・資格・備品の横スクロール軽減（統一）を実行しました");
  }
}

function applyCompactLicenseEquipmentLayout_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const sheetName = sheet.getName();
  const headers = getHeaders_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = headers.length;
  const personal = getPersonalMembers_();

  // 全体調整・今のシートだけ調整・免許資格備品専用調整で同じ見た目にする。
  sheet.setFrozenColumns(0);
  showAllColumns_(sheet);

  headers.forEach((header, index) => {
    const col = index + 1;
    if (!header) return;

    let width = 92;

    if (sheetName === "運転免許管理") {
      if (header === "所有者") width = 120;
      else if (LICENSE_TYPE_HEADERS.includes(header)) width = 54;
      else if (["取得日", "更新期限"].includes(header)) width = 112;
      else if (header === "コピー有無") width = 96;
      else if (header === "状態") width = 96;
      else if (header === "通知") width = 104;
      else if (header === READ_HEADER || personal.includes(header)) width = 40;
      else if (header === "備考") width = 180;
      else if (header === "免許ID") width = 40;
    } else if (sheetName === "資格管理") {
      if (header === "所有者") width = 112;
      else if (header === "資格名") width = 165;
      else if (header === "区分") width = 74;
      else if (["取得日", "更新期限"].includes(header)) width = 106;
      else if (header === "コピー有無") width = 96;
      else if (header === "状態") width = 96;
      else if (header === "通知") width = 104;
      else if (header === READ_HEADER || personal.includes(header)) width = 40;
      else if (header === "備考") width = 160;
      else if (header === "資格ID") width = 40;
    } else if (sheetName === "備品修理管理") {
      // 通知列までがなるべく1画面に収まるよう、内容列だけ少し広めにして他は圧縮。
      if (header === "購入日") width = 96;
      else if (header === "備品名") width = 120;
      else if (header === "修理業者") width = 115;
      else if (header === "内容") width = 155;
      else if (header === "修理依頼者") width = 92;
      else if (["修理依頼日", "返却予定日"].includes(header)) width = 96;
      else if (header === "返却済") width = 62;
      else if (header === "状態") width = 92;
      else if (header === "通知") width = 104;
      else if (header === READ_HEADER || personal.includes(header)) width = 40;
      else if (header === "備考") width = 150;
      else if (header === "修理ID") width = 40;
    }

    sheet.setColumnWidth(col, width);

    // 縦書きは、免許種別・既読・個人確認だけ。
    // 状態・通知・コピー有無・返却済は横書きで色と文字を読めるようにする。
    const shouldVertical = LICENSE_TYPE_HEADERS.includes(header) || header === READ_HEADER || personal.includes(header);

    const range = sheet.getRange(1, col, lastRow, 1);
    if (shouldVertical) {
      range
        .setVerticalText(true)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
    } else {
      range
        .setVerticalText(false)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
    }

    if (DATE_HEADERS.includes(header) || ["コピー有無", "状態", "通知", "区分", "返却済"].includes(header) || shouldVertical) {
      sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    } else {
      sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1)
        .setVerticalAlignment("middle");
    }
  });

  // 内部列は削除せず非表示。
  const idHeader = SHEET_ID_HEADERS[sheetName];
  if (idHeader) hideColumnByHeaderSafely_(sheet, idHeader);
  [CALENDAR_ID_HEADER, "写真", "レシート写真", "PDFリンク"].forEach(header => hideColumnByHeaderSafely_(sheet, header));

  rebuildPersonalCheckGroupForSheet_(sheet);

  const maxRows = Math.min(sheet.getMaxRows(), Math.max(lastRow, 20));
  sheet.setRowHeight(1, sheetName === "運転免許管理" ? 76 : 64);
  sheet.setRowHeights(2, Math.max(maxRows - 1, 1), 34);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  applyColorRules(sheet);
  createFilterSafely_(sheet, lastCol);
}

function hideColumnByHeaderSafely_(sheet, header) {
  if (!sheet || !header) return;
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(header) + 1;
  if (col <= 0) return;
  try {
    sheet.hideColumns(col);
  } catch (e) {
    console.log("hideColumnByHeaderSafely_ skipped: " + sheet.getName() + " / " + header + " / " + e.message);
  }
}


function appendGptAllCompactLayoutDiagnostics_(ss, rows) {
  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      rows.push(["表示", name, "全入力シート横スクロール軽減", "NG", "シートがありません", "シート名を確認"]);
      return;
    }

    const headers = getHeaders_(sheet);
    const idHeader = SHEET_ID_HEADERS[name];
    const idCol = idHeader ? headers.indexOf(idHeader) + 1 : 0;
    const hiddenId = idCol > 0 ? sheet.isColumnHiddenByUser(idCol) : false;
    const calendarCol = headers.indexOf(CALENDAR_ID_HEADER) + 1;
    const hiddenCalendar = calendarCol > 0 ? sheet.isColumnHiddenByUser(calendarCol) : true;
    const frozenCols = sheet.getFrozenColumns();
    const wideColumns = [];

    headers.forEach((header, index) => {
      const col = index + 1;
      if (!header) return;
      const width = sheet.getColumnWidth(col);
      if (width >= 210 && !["内容", "作業内容", "日報文章", "備考", "気になった内容"].includes(header)) {
        wideColumns.push(header + ":" + width);
      }
    });

    let status = "OK";
    const notes = [];
    if (frozenCols > 0) {
      status = "要確認";
      notes.push("固定列あり:" + frozenCols);
    }
    if (wideColumns.length) {
      status = "要確認";
      notes.push("幅広列:" + wideColumns.join(", "));
    }

    rows.push([
      "表示",
      name,
      "全入力シート横スクロール軽減",
      status,
      "列数: " + headers.length + " / 固定列: " + frozenCols + " / ID列非表示: " + (hiddenId ? "はい" : "いいえ") + " / カレンダーID非表示: " + (hiddenCalendar ? "はい" : "いいえ") + (notes.length ? " / " + notes.join(" / ") : ""),
      "横スクロールが大きい場合は『全入力シートの横スクロール軽減（統一）』を実行。全体調整・個別調整・免許資格備品専用調整で同じ列幅基準を使用。運転免許は免許種別列を少し広めにしてズレを抑制"
    ]);
  });
}


function appendGptCompactLayoutDiagnostics_(ss, rows) {
  ["運転免許管理", "資格管理", "備品修理管理"].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) {
      rows.push(["表示", name, "横スクロール軽減", "NG", "シートがありません", "シート名を確認"]);
      return;
    }

    const headers = getHeaders_(sheet);
    const idHeader = SHEET_ID_HEADERS[name];
    const idCol = idHeader ? headers.indexOf(idHeader) + 1 : 0;
    const hiddenId = idCol > 0 ? sheet.isColumnHiddenByUser(idCol) : false;
    const wideColumns = [];

    headers.forEach((header, index) => {
      const col = index + 1;
      if (!header) return;
      const width = sheet.getColumnWidth(col);
      if (width >= 190 && !["内容", "備考"].includes(header)) {
        wideColumns.push(header + ":" + width);
      }
    });

    const status = wideColumns.length ? "要確認" : "OK";
    rows.push([
      "表示",
      name,
      "横スクロール軽減",
      status,
      "列数: " + headers.length + " / ID列非表示: " + (hiddenId ? "はい" : "いいえ") + " / 幅広列: " + (wideColumns.length ? wideColumns.join(", ") : "なし"),
      "横スクロールが大きい場合は『免許・資格・備品の横スクロール軽減（統一）』を実行。固定列は使わない設定"
    ]);
  });
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
  applyPersonalCheckColumnLayout_(sheet);
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
  const rowCount = Math.max(getApplyRowCount_(sheet), 1);
  const noticeCol = headers.indexOf("通知") + 1;
  const statusCol = Math.max(headers.indexOf("状態") + 1, headers.indexOf("車検状態") + 1, headers.indexOf("対応状況") + 1);
  const phoneCol = headers.indexOf("電話対応") + 1;
  const readCol = headers.indexOf(READ_HEADER) + 1;
  const typeCol = headers.indexOf("種類") + 1;
  const importanceCol = headers.indexOf("重要度") + 1;
  const accountingCol = headers.indexOf("経理確認") + 1;
  const settlementCol = headers.indexOf("精算状態") + 1;
  const copyCol = headers.indexOf("コピー有無") + 1;
  const returnedCol = headers.indexOf("返却済") + 1;
  const categoryCol = headers.indexOf("区分") + 1;

  // v9.7.14 淡色カラー設計:
  // 1) 目に強すぎる濃色をやめ、全体をパステル寄りにする。
  // 2) ただし「期限切れ/差戻し/中止/失効」は淡い赤系で分かるように残す。
  // 3) 似ている色は、赤・橙・黄・緑・青・紫・茶灰に系統分けして被りを減らす。
  // 4) 一覧の種類は淡い色でも見分けやすいように、色相をずらす。

  // 通知：期限・重要度。淡色でも優先度が分かるように赤→橙→黄→緑→青灰で整理。
  addTextRule_(rules, sheet, noticeCol, "期限切れ", "#f4cccc", rowCount); // 淡い赤
  addTextRule_(rules, sheet, noticeCol, "今日", "#fce5cd", rowCount); // 淡い橙
  addTextRule_(rules, sheet, noticeCol, "3日以内", "#fff2cc", rowCount); // 淡い黄
  addTextRule_(rules, sheet, noticeCol, "7日以内", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, noticeCol, "30日以内", "#d9eaf7", rowCount); // 淡い青灰
  addTextRule_(rules, sheet, noticeCol, "重要", "#ead1dc", rowCount); // 淡い赤紫
  addTextRule_(rules, sheet, noticeCol, "未確認", "#fff7d6", rowCount); // さらに淡い黄
  addTextRule_(rules, sheet, noticeCol, "差戻し", "#ead1dc", rowCount); // 淡い赤紫
  addTextRule_(rules, sheet, noticeCol, "未精算", "#fde9d9", rowCount); // 淡い橙

  // 状態：進捗・完了・異常を淡色で分ける。
  addTextRule_(rules, sheet, statusCol, "予定", "#eeeeee", rowCount); // 淡い灰
  addTextRule_(rules, sheet, statusCol, "移動中", "#d9eaf7", rowCount); // 淡い青
  addTextRule_(rules, sheet, statusCol, "着工前", "#e4dfec", rowCount); // 淡い紫
  addTextRule_(rules, sheet, statusCol, "施工中", "#cfe2f3", rowCount); // 淡い青
  addTextRule_(rules, sheet, statusCol, "進行中", "#d0e0e3", rowCount); // 淡い青緑
  addTextRule_(rules, sheet, statusCol, "未対応", "#f4cccc", rowCount); // 淡い赤
  addTextRule_(rules, sheet, statusCol, "対応中", "#fff2cc", rowCount); // 淡い黄
  addTextRule_(rules, sheet, statusCol, "折返し", "#fce5cd", rowCount); // 淡い橙
  addTextRule_(rules, sheet, statusCol, "要確認", "#ead1dc", rowCount); // 淡い赤紫
  addTextRule_(rules, sheet, statusCol, "予約済", "#d0e0e3", rowCount); // 淡い青緑
  addTextRule_(rules, sheet, statusCol, "実施済", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, statusCol, "更新予定", "#fce5cd", rowCount); // 淡い橙
  addTextRule_(rules, sheet, statusCol, "有効", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, statusCol, "未依頼", "#f4cccc", rowCount); // 淡い赤
  addTextRule_(rules, sheet, statusCol, "依頼済", "#fce5cd", rowCount); // 淡い橙
  addTextRule_(rules, sheet, statusCol, "修理中", "#cfe2f3", rowCount); // 淡い青
  addTextRule_(rules, sheet, statusCol, "修理不可", "#d7ccc8", rowCount); // 淡い茶灰
  addTextRule_(rules, sheet, statusCol, "下書き", "#eeeeee", rowCount); // 淡い灰
  addTextRule_(rules, sheet, statusCol, "提出済", "#e4dfec", rowCount); // 淡い紫
  addTextRule_(rules, sheet, statusCol, "確認済", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, statusCol, "差戻し", "#ead1dc", rowCount); // 淡い赤紫
  addTextRule_(rules, sheet, statusCol, "返却済", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, statusCol, "完了", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, statusCol, "更新済", "#cfe8e6", rowCount); // 淡い青緑
  addTextRule_(rules, sheet, statusCol, "延期", "#fce5cd", rowCount); // 淡い橙
  addTextRule_(rules, sheet, statusCol, "中止", "#f4cccc", rowCount); // 淡い赤
  addTextRule_(rules, sheet, statusCol, "期限切れ", "#f4cccc", rowCount); // 淡い赤
  addTextRule_(rules, sheet, statusCol, "失効", "#ead1dc", rowCount); // 淡い赤紫

  // 電話対応：淡色だが、未対応/対応中/折返し/完了を区別。
  addTextRule_(rules, sheet, phoneCol, "未対応", "#f4cccc", rowCount);
  addTextRule_(rules, sheet, phoneCol, "対応中", "#fff2cc", rowCount);
  addTextRule_(rules, sheet, phoneCol, "折返し", "#fce5cd", rowCount);
  addTextRule_(rules, sheet, phoneCol, "完了", "#d9ead3", rowCount);

  // レシート・経理系：経理確認と精算状態が通知と似すぎないように淡色で整理。
  addTextRule_(rules, sheet, accountingCol, "未確認", "#fff7d6", rowCount);
  addTextRule_(rules, sheet, accountingCol, "確認中", "#fff2cc", rowCount);
  addTextRule_(rules, sheet, accountingCol, "確認済", "#d9ead3", rowCount);
  addTextRule_(rules, sheet, accountingCol, "差戻し", "#ead1dc", rowCount);
  addTextRule_(rules, sheet, settlementCol, "未精算", "#fde9d9", rowCount);
  addTextRule_(rules, sheet, settlementCol, "精算中", "#fff2cc", rowCount);
  addTextRule_(rules, sheet, settlementCol, "精算済", "#cfe8e6", rowCount);

  // コピー有無・返却済・区分:
  // 同じプルダウン項目は、シートが違ってもなるべく同じ色にそろえる。
  addTextRule_(rules, sheet, copyCol, "有", "#d9ead3", rowCount);
  addTextRule_(rules, sheet, copyCol, "無", "#f4cccc", rowCount);
  addTextRule_(rules, sheet, copyCol, "未確認", "#fff2cc", rowCount);
  addTextRule_(rules, sheet, returnedCol, "済", "#d9ead3", rowCount);
  addTextRule_(rules, sheet, returnedCol, "未", "#fce5cd", rowCount);
  addTextRule_(rules, sheet, categoryCol, "燃料", "#d9eaf7", rowCount);
  addTextRule_(rules, sheet, categoryCol, "駐車場", "#d0e0e3", rowCount);
  addTextRule_(rules, sheet, categoryCol, "高速代", "#cfe2f3", rowCount);
  addTextRule_(rules, sheet, categoryCol, "消耗品", "#fff2cc", rowCount);
  addTextRule_(rules, sheet, categoryCol, "資材", "#e4dfec", rowCount);
  addTextRule_(rules, sheet, categoryCol, "工具", "#fce5cd", rowCount);
  addTextRule_(rules, sheet, categoryCol, "修理", "#fde9d9", rowCount);
  addTextRule_(rules, sheet, categoryCol, "その他", "#eeeeee", rowCount);

  // 一覧スケジュールの種類色：淡いパステルで、同じ系統が隣りすぎないように分離。
  addTextRule_(rules, sheet, typeCol, "車検", "#f4cccc", rowCount); // 淡赤
  addTextRule_(rules, sheet, typeCol, "運転免許", "#fce5cd", rowCount); // 淡橙
  addTextRule_(rules, sheet, typeCol, "資格", "#fff2cc", rowCount); // 淡黄
  addTextRule_(rules, sheet, typeCol, "備品修理", "#e4dfec", rowCount); // 淡紫
  addTextRule_(rules, sheet, typeCol, "工事予定", "#cfe2f3", rowCount); // 淡青
  addTextRule_(rules, sheet, typeCol, "電話履歴", "#fde9d9", rowCount); // 淡い橙桃
  addTextRule_(rules, sheet, typeCol, "出先予定", "#d9ead3", rowCount); // 淡緑
  addTextRule_(rules, sheet, typeCol, "会議", "#d0e0e3", rowCount); // 淡青緑
  addTextRule_(rules, sheet, typeCol, "行事", "#ead1dc", rowCount); // 淡桃
  addTextRule_(rules, sheet, typeCol, "作業状況", "#d9eaf7", rowCount); // 淡青灰
  addTextRule_(rules, sheet, typeCol, "社用車予約", "#d9ead3", rowCount); // 淡緑
  addTextRule_(rules, sheet, typeCol, "日報", "#eaf2d3", rowCount); // 淡ライム
  addTextRule_(rules, sheet, typeCol, "日報レシート", "#d7ccc8", rowCount); // 淡茶灰
  addTextRule_(rules, sheet, typeCol, "ToDo", "#e4dfec", rowCount); // 淡紫
  addTextRule_(rules, sheet, typeCol, "お知らせ", "#ead1dc", rowCount); // 淡赤紫

  addTextRule_(rules, sheet, importanceCol, "高", "#ead1dc", rowCount);
  addTextRule_(rules, sheet, importanceCol, "中", "#fff2cc", rowCount);
  addTextRule_(rules, sheet, importanceCol, "低", "#d9ead3", rowCount);

  // 既読がFALSEのセルだけ淡い淡い黄。空行はチェックボックスを消すので対象外になりやすい。
  if (readCol > 0) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied("=$" + columnLetter_(readCol) + "2=FALSE")
        .setBackground("#fff7d6")
        .setRanges([sheet.getRange(2, readCol, rowCount, 1)])
        .build()
    );
  }

  sheet.setConditionalFormatRules(rules);
}
function addTextRule_(rules, sheet, col, text, color, rowCount) {
  if (!col || col <= 0) return;
  const range = sheet.getRange(2, col, Math.max(rowCount || getApplyRowCount_(sheet), 1), 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(color).setRanges([range]).build());
}

function columnLetter_(col) {
  let temp = "";
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    n = Math.floor((n - rem - 1) / 26);
  }
  return temp;
}

function applyStatusBackground_(sheet, col, startRow, numRows) {
  if (numRows <= 0) return;
  const range = sheet.getRange(startRow, col, numRows, 1);
  const values = range.getValues();
  const backgrounds = values.map(r => [String(r[0]) === "OK" ? "#d9ead3" : String(r[0]) === "NG" ? "#ea9999" : "#fff2cc"]);
  range.setBackgrounds(backgrounds);
}


/**
 * 既読・個人確認列の見た目だけを全シートで再調整する。
 * 空行チェックボックス削除、入力済み行だけチェックボックス表示、個人確認列の縦書き・折りたたみをまとめて行う。
 */
function applyReadDisplayToAllSheets() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中のため、既読表示の再調整をスキップしました");
    return;
  }

  try {
    Object.keys(getSheetHeaders_()).forEach(name => {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
      if (!sheet) return;
      setCheckboxesForDataRows(sheet);
      applyPersonalCheckColumnLayout_(sheet);
      applyColorRules(sheet);
    });
    rebuildCheckGroups();
    toast_("既読表示・個人確認列を再調整しました");
  } finally {
    lock.releaseLock();
  }
}

function applyPersonalCheckColumnLayout_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = getHeaders_(sheet);
  const members = getPersonalMembers_();
  const lastRow = Math.max(sheet.getLastRow(), 1);

  // v9.7.11:
  // setTextRotation(90) は「文字を横向きに回転」するだけで、
  // Google Sheets上では見た目が横向き/寝た文字に見える場合がある。
  // 既読・個人確認列の見出しは setVerticalText(true) で本当の縦書きにする。
  const checkHeaders = [READ_HEADER, ...members];

  checkHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;

    sheet.setColumnWidth(col, 44);

    try {
      const headerCell = sheet.getRange(1, col);
      headerCell
        .setVerticalText(true)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(false);

      if (lastRow >= 2) {
        sheet.getRange(2, col, lastRow - 1, 1)
          .setVerticalText(false)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      }
    } catch (e) {
      console.log(sheet.getName() + " の既読/個人確認列レイアウト調整をスキップ: " + e.message);
    }
  });
}

function rebuildCheckGroups() {
  resetColumnGroups_();
  createPersonalCheckColumnGroups_();
  toast_("個人確認グループを作り直しました");
}

function resetColumnGroups_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(getSheetHeaders_()).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;
    try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (e) {}
    for (let i = 0; i < 10; i++) {
      try { sheet.getRange(1, 1, 1, sheet.getLastColumn()).shiftColumnGroupDepth(-1); } catch (e) {}
    }
  });
}

function createPersonalCheckColumnGroups_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = getPersonalMembers_();
  Object.keys(getSheetHeaders_()).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;
    const headers = getHeaders_(sheet);
    const personalCols = members.map(member => headers.indexOf(member) + 1).filter(col => col > 0);
    if (!personalCols.length) return;

    const startCol = Math.min(...personalCols);
    const endCol = Math.max(...personalCols);
    try {
      // 既読列は常時表示。折りたたむのは個人確認列だけ。
      sheet.getRange(1, startCol, Math.max(sheet.getMaxRows(), 1), endCol - startCol + 1).shiftColumnGroupDepth(1);
    } catch (e) {
      console.log(name + " の個人確認グループ作成をスキップ: " + e.message);
    }
  });
}


function rebuildPersonalCheckGroupForSheet_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const headers = getHeaders_(sheet);
  const members = getPersonalMembers_();
  const personalCols = members
    .map(member => headers.indexOf(member) + 1)
    .filter(col => col > 0);

  if (!personalCols.length) return;

  const startCol = Math.min(...personalCols);
  const endCol = Math.max(...personalCols);
  const width = endCol - startCol + 1;

  try {
    // 対象シートの個人確認列だけ、既存グループを軽く解除して作り直す。
    // 全シートのグループを触らないので、横スクロール軽減メニュー用として安全。
    const groupRange = sheet.getRange(1, startCol, Math.max(sheet.getMaxRows(), 1), width);
    for (let i = 0; i < 5; i++) {
      try {
        groupRange.shiftColumnGroupDepth(-1);
      } catch (e) {
        break;
      }
    }
    groupRange.shiftColumnGroupDepth(1);
  } catch (e) {
    console.log(sheet.getName() + " の個人確認グループ作成をスキップ: " + e.message);
  }
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

/**
 * シート順を安全に並び替える。
 *
 * 新セットアップ中は、SHEET_ORDER に定義されていても
 * まだ存在しないシートがある場合があります。
 * その状態で i + 1 を移動先に使うと、実在シート数を超えて
 * Exception: 引数が無効です が出るため、存在するシートだけを
 * position = 1, 2, 3... の順に詰めて移動します。
 */
function reorderSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let position = 1;

  SHEET_ORDER.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    try {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(position);
      position++;
    } catch (e) {
      console.log("シート並び替えをスキップ: " + name + " / " + e.message);
    }
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
































