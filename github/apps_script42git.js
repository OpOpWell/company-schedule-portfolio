
/**
 * 社内共有業務管理システム v9.7.20 電話履歴 シンプル日時ソート版（完全1ファイル貼り替え用）
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
 * v9.7.17:
 * - 電話対応を「空欄に戻す」にした時、対応完了日時も空欄へ戻す。
 * - 対応完了日時の表示形式を yyyy/mm/dd hh:mm に統一。
 * - 対応完了日時を日付列ではなく日時列として扱い、日付だけ表示になる問題を修正。
 * v9.7.18:
 * - 電話対応を空欄に戻した時、通知も空欄へ戻す。
 * - 電話履歴の通知判定を専用化し、空欄行が期限切れへ戻る問題を修正。
 * v9.7.19:
 * - 完了から空欄へ戻した時も電話履歴を再ソートし、完了グループ下部から未完了側へ戻す。
 * - 電話対応を変更した場合は、未対応/対応中/折返し/空欄/その他/完了の順に整理する。
 * v9.7.20:
 * - 電話履歴の並び順を簡素化。完了以外を上、完了を下にする。
 * - 未対応/対応中/折返し/空欄の細かい状態順ソートはやめ、各グループ内は日付・時間の新しい順に統一。
 *
 * v9.7.28:
 * - 横長の資格入力シート方式を廃止。
 * - 資格管理1シートに集約し、所有者ごとの並び替え・行グループ化に変更。
 * - 資格管理の重複チェックシートを追加。
 *
 * v9.7.31:
 * - 横長チェック入力の常用をやめ、資格一括登録シート方式へ変更。
 * - 資格マスタと設定シートを元に、資格名・社員名のプルダウンを最小範囲だけ反映。
 * - 資格一括登録から資格管理へ安全に反映し、同じ人・同じ区分・同じ資格は重複追加しない。
 *
 * v9.7.33:
 * - 資格管理の資格ID列を自動非表示にし、折り返しを止めて行高が大きくなる問題を修正。
 * - 資格管理専用のコンパクト表示メニューを追加。
 *
 * v9.7.34:
 * - 資格かんたん登録シートを追加。
 * - 1行に複数資格を「、」「,」「改行」区切りで貼り付け、資格管理へ縦持ち展開できるようにした。
 * - 資格一括登録は残しつつ、普段の初回入力は資格かんたん登録を推奨。
 *
 * v9.7.37:
 * - 資格かんたん登録の「追加する資格」は「区分｜資格名」表示に変更。
 * - 取得区分の混在を防止。施工管理技士系だけ技士/技士補を許可し、それ以外は保有扱い。
 * - 選択時は区分を自動補正し、資格名まとめには資格名だけ追記。
 * - 区分違いの混在を防ぐ注意表示と、社員別資格一覧の人ごと折りたたみを維持。
 *
 * v9.7.47:
 * - 資格かんたん登録の「区分」と「追加する資格」プルダウンのズレを修正。
 * - 区分変更時に編集行の資格候補を必ず再構築し、古い候補を残さない。
 * - 資格かんたん登録の区分別プルダウンを全行再構築する保守メニューを追加。
 *
 * v9.7.48:
 * - 1. 資格マスタに「通知日前」列を正式追加。
 * - 2. 資格管理/資格かんたん登録の資格プルダウンを資格マスタ表示順・区分連動へ統一。
 * - 3. 資格管理の所有者グループを、ソート後に作り直す処理を強化。
 * - 4. 社員別資格一覧を社員見出し＋明細グループで見やすく再生成。
 * - 5. 車検予定日は「車検に出す予定日」とし、車検履歴へ別列で記録。
 * - 6. 車検管理の車両名プルダウンを設定シート基準に変更。
 * - 7. 工事予定の契約金額は円表示、税込/税抜/不明は短い「税」列のプルダウンで管理。
 * - 8. 備品修理管理の修理業者を設定シートへ自動マスタ追加。
 * - 9. 行事予定は備考を手前にし、内容列は削除せず非表示移行。
 * - 10. 裏方シート非表示対象を再整理。
 *
 * v9.7.48c:
 * - 社用版では担当者・既読確認者のサンプル名を自動補充しない。
 * - 設定シートに会社用の担当者がある場合、山田/鈴木などを戻さない。
 * - 「会社用：サンプル担当者を設定シートから削除」メニューを追加。
 * - 社員別資格一覧の見出しに取得数を表示。資格管理本体には取得数列を持たせない。
 *
 * v9.7.48d:
 * - 設定シートだけでなく、既存入力シートに残った山田/高橋/鈴木などの個人確認列も削除できるようにした。
 * - 「既読確認者」設定を正として、各入力シートの個人確認列を再構成する保守メニューを追加。
 * - 入力シートだけ設定反映時にも、古い個人確認列を設定シート基準で整理する。
 *
 * v9.7.48e:
 * - 資格管理を手入力したとき、取得区分・保有状況を勝手に自動入力しないように変更。
 * - 資格管理のonEditは、区分と資格名の相互補正、通知更新、要確認メモだけに限定。
 * - 取得区分・保有状況の自動補完は、資格かんたん登録から資格管理へ反映する時と、手動の安全補正メニューだけに限定。
 * v9.7.48f:
 * - 社員別資格一覧の明細から所有者列を外し、社員名は見出し行だけに表示。
 * - 資格管理・社員別資格一覧の資格区分、取得区分、保有状況、状態、通知の淡色カラーリングを強化。
 * v9.7.48i:
 * - GPT診断用シートの期待ヘッダーを、社員別資格一覧の所有者列なし仕様に合わせて修正。
 * - 資格一括登録、資格重複チェック、資格別保有者一覧の旧「保有区分」診断を減らすため、診断前に軽量ヘッダー修復を実行。
 * - 資格管理だけ既読・個人確認チェックボックスを再設定する軽量メニューを追加。
 * v9.7.48j:
 * - 備品修理管理の修理業者は、候補を表示しつつマスタ外の手入力も許可する。
 * - 手入力された修理業者を設定シートへ自動追加し、修理業者行を重複なしでソートする。
 * - GPT診断用シートに、修理業者マスタ・プルダウン候補外入力許可の診断を追加。
 *
 * v9.7.48k:
 * - 修理業者マスタ＋プルダウン修復をトップメニューにも追加。
 * - 予定・車検・工事・備品サブメニュー内でも修理業者修復を上側に移動し、見つけやすくした。
 *
 * v9.7.48p2:
 * - 資格管理の既読・個人確認チェック不足を強制復旧する専用処理を追加。
 * - 資格管理だけ入力済み行を判定し、既読＋個人確認列にチェックボックスを再挿入。
 *
 * v9.7.48p3:
 * - GPT診断用シート作成時にも資格管理の既読・個人確認チェックを再確認する。
 * - 資格管理の表示/入力規則調整後に、最後もう一度チェックボックスを強制再設定する。
 * - 診断表が修復直後のチェックボックスを取りこぼして860件不足のまま出る問題を抑制。
 * - 空行にはチェックボックスを出さず、既存 TRUE/FALSE は可能な範囲で保持。
 * v9.7.48p5:
 * - 後付けの保守メニュー実行が多くなったため、「会社用：ベータ前チェックパック」を追加。
 * - ベータ前に必要な軽量修復、IDだけ空行整理、資格管理チェック再設定、一覧/集計/GPT診断を一括実行。
 * - 新規ベータ作成・旧列名移行・書式込み全体更新などの危険/重い処理は含めない。
 *
 * v9.7.48p6:
 * - 各入力シートへ AUTO_TEST_ 付きのテスト行を自動投入し、一覧反映・ID・既読チェック・車検履歴・修理業者マスタを自動確認する「会社用：自動テストパック」を追加。
 * - テストデータは「会社用：テストデータ削除」でまとめて削除できるようにした。
 * - 本番データと混ざらないよう、テスト行の備考・内容・車番・修理業者へ AUTO_TEST_ マーカーを入れる。
 *
 * v9.7.48p10:
 * - テストデータ削除を強化し、過去版のフィードバック自動テスト残骸（対象シート=自動テスト、対応方針=自動テスト等）も安全に削除対象へ含める。
 * - AUTO_TEST_ 行削除後に自動テスト結果シートへ削除件数と残骸チェック結果を残す。
 * - 設定シートの AUTO_TEST_ 修理業者、一覧・要確認・集計側の AUTO_TEST_ 行もまとめて掃除する。

 * v9.7.48p14:
 * - ID列・カレンダーID列が表示調整後に再表示される問題を修正。
 * - 日常更新、表示調整、診断作成後にも内部ID列を自動で非表示へ戻す。
 * - AppSheet用IDは削除せず、非表示だけにする。
 *
 * v9.7.48p15:
 * - ダッシュボードが汎用表示調整や古いフィルタ/行グループ/余分な列の影響で崩れる問題を修正。
 * - ダッシュボード作成時にフィルタ・結合・条件付き書式・余分な行列をリセットし、固定の見やすい表示へ戻す。
 * - 「会社用：ダッシュボードを作り直す」メニューを追加。
 *
 * v9.7.48p23:
 * - 資格管理の所有者グループ作成時に cleanupQualificationFalseDisplay_ が未定義で止まる問題を修正。
 * - p19で追加した cleanupQualificationManagementFalseDisplay_ へ統一し、旧関数名の互換ラッパーも追加。
 * - 人ごとグループ化メニューでも、FALSE表示掃除・チェックボックス復旧・ID列非表示を安全に再実行する。
 *
 * v9.7.48p24:
 * - 資格管理C列「資格名」を、B列「区分」に応じて1行ずつ厳密に入力規則設定する方式へ変更。
 * - 通常の入力シート設定反映でC列が全資格候補へ戻る問題を防止。
 * - 区分が資格なら資格だけ、技能講習なら技能講習だけを表示するよう再修正。
 *
 * v9.7.48p16:
 * - 「会社用：資格管理だけ整える」実行時にも、資格管理の列ズレ安全修復を必ず実行するよう修正。
 * - F列「取得日」が空欄見出しになる事故を、資格管理パックだけで復旧できるようにした。
 * - 安全教育「除雪講習」など、紙版資格マスタ由来の不足候補補完も資格管理パックに含めた。
 *
 * v9.7.48p17:
 * - 資格管理で区分を変更した時、資格名プルダウンをその区分の資格だけに絞り、資格マスタ表示順で並ぶよう強化。
 * - 区分と資格名が食い違う既存値は、区分変更時に資格名を空欄へ戻して誤登録を防止。
 * - 資格管理パック実行時にも区分別資格名プルダウンを再ソートする。
 *
 * v9.7.48p18:
 * - 社内共有時に迷わないよう、社内管理メニューを「日常用」「保守用」「危険操作」に整理。
 * - 通常メニューの表示項目を減らし、自動テスト・詳細修復・サンプル・移行系はサブメニューへ移動。
 * - 裏方シート非表示対象に自動テスト結果を追加し、設定・資格マスタ・ダッシュボードは表示のままにする説明へ修正。
 *
 * v9.7.48p19:
 * - 資格管理の既読・個人確認列が FALSE 文字表示に戻る問題を修正。
 * - 資格管理パックの最後で、備考など通常列へ漏れた TRUE/FALSE を掃除し、チェックボックス入力規則を最終再設定する。
 * - 「資格管理のFALSE表示を修復」メニューを追加。
 *
 * v9.7.48p20:
 * - 「会社用：資格管理だけ整える」を軽量化し、社員別/資格別一覧作成や所有者グループ作成など重い処理を分離。
 * - 資格管理パック実行時の起動時間超過を避けるため、列修復・区分別資格名プルダウン・FALSE表示修復・ID非表示だけを日常側で実行。
 * - ID列/カレンダーID列をセットアップ・設定反映・資格管理修復後に標準で非表示へ戻す。
 * - 社員別資格一覧/資格別保有者一覧は「資格一覧だけ更新」メニューで別実行に変更。
 *
 * v9.7.48p21:
 * - 資格マスタ内に旧区分の重複候補が残っていても、紙版標準マスタの区分を優先して資格名候補を作るよう修正。
 * - 資格管理C列「資格名」は、B列「区分」ごとに直接リスト型の入力規則を作り直し、隠し補助シートや旧入力規則の影響を受けにくくした。
 * - 区分が「資格」の行では、技能講習・特別教育・安全教育・免許の候補が出ないようにした。
 *
 * v9.7.48p22:
 * - 資格管理の並び替え・区分別資格名プルダウン修復時に、既存の入力規則が先に残って C列 setValues が止まる問題を修正。
 * - C列「資格名」は、値を整理してから区分別入力規則を貼り直す順序に変更。
 * - 資格管理のソート時は一時的に入力規則を外し、値の移動後に再設定する安全処理へ変更。
 *
 * v9.7.48p25:
 * - 高所作業車、車両系建設機械（締固）を特別教育へ統一。
 * - 資格管理を編集した時、所有者＋資格マスタ表示順で軽量自動ソートする処理を追加。
 * - 資格管理C列の区分別入力規則を、行単位から連続範囲単位へまとめて貼るよう効率化。
 *
 * v9.7.48p29:
 * - 資格マスタIDなどの内部ID列が紙順マスタ再作成後に表示される問題を修正。
 * - 資格マスタ整形・紙順マスタ修復・通常表示調整の最後にID列非表示を標準適用。
 * - ID列/カレンダーID列は削除せず、各シートでデフォルト非表示に戻す。


 * v9.7.48p26:
 * - 「会社用：資格管理だけ整える（軽量）」の起動時間超過対策。
 * - 軽量パックから全行ソート、社員別/資格別一覧、重いカラーリング、二重チェックボックス復旧を外した。
 * - 資格管理C列の区分別プルダウンは、連続範囲単位で維持する。
 * - FALSE掃除と既読チェックボックス復旧を1回読み取り中心の軽量処理へ変更。
 *
 * v9.7.48p27:
 * - 資格管理で1件入力するたびに自動ソートしていた処理を停止。
 * - 入力行が移動して、次の空行の資格名プルダウンが消えたように見える問題を修正。
 * - 資格管理C列の区分別プルダウンを、実データ最終行＋予備行まで維持する。
 * - 並び替えは手入力中に自動実行せず、必要な時だけ保守メニューから実行する。

 * v9.7.48p30:
 * - 資格管理の並び順が入力順のまま残る問題を修正。
 * - 「会社用：資格管理だけ整える（軽量）」でも、値だけの軽量ソートを実行する。
 * - 資格マスタシートに旧表示順が残っていても、紙版マスタ順を優先して資格管理・プルダウンを並べる。
 * - 入力中のonEdit自動ソートは引き続き停止し、入力後のメニュー実行で並べ替える。
 *
 * v9.7.48p34:
 * - 資格管理で行挿入後に所有者・区分・資格名・取得区分などのプルダウンが消える問題を再修正。
 * - 入力規則復旧時に先に clearDataValidations() で消してから貼る方式をやめ、上書き貼りに変更。途中停止してもプルダウンだけ消えた状態を残しにくくした。
 * - 資格管理軽量パック・プルダウン復旧メニュー実行時に、行挿入検知トリガーも自動確認する。
 * - onChangeは行/列挿入・削除・OTHERだけを対象にし、ロックで多重実行を抑制する。
 * v9.7.48p38:
 * - 「資格管理だけ整える（軽量）」を超軽量化。
 * - 通常メニューでは、ヘッダー最小確認・資格管理プルダウン復旧・ID非表示だけ実行する。
 * - ソート、FALSE掃除、チェックボックス全復旧、社員別一覧作成は別メニューへ分離。
 * - onChangeはOTHERを拾わず、行/列挿入・削除だけに限定して多重起動を抑制。
 *
 * v9.7.48p39:
 * - 工事予定の「連絡先」を文字列固定にし、0128など先頭の0が消えないよう修正。
 * - 工事予定の契約金額・税・状態修復時にも、連絡先列を文字列形式へ戻す。
 * - 連絡先セル編集時に、可能な範囲で入力値を文字列として保持する軽量処理を追加。
 *
 * v9.7.48p43:
 * - 運転免許更新履歴シートを追加。免許証有効期限・コピー有無・状態を変更した時だけ軽量に履歴化する。
 * - 資格更新履歴シートを追加。資格管理の更新期限・コピー有無・状態を変更した時だけ軽量に履歴化する。
 * - 履歴は全編集履歴ではなく、更新確認で必要な項目だけを残す設計にし、普段の入力処理を重くしない。

 * v9.7.48p44:
 * - GitHub/ポートフォリオ公開向けに、コード内のデフォルト社用車候補をサンプル名へ匿名化。
 * - 実運用の担当者・既読確認者・社用車・車両名は、引き続き設定シートを正として扱う。
 * - 既存の設定シートに会社用データがある場合は、サンプル候補を自動追加しない方針を維持。
 * - p43の運転免許更新履歴・資格更新履歴、p39の工事連絡先文字列固定、p38の超軽量化は維持。

 */

const SYSTEM_VERSION = "v9.7.48p44-github-safe-sample-settings";
const AUTO_TEST_MARKER_PREFIX = "AUTO_TEST_";
const READ_HEADER = "既読";
const CLEAR_LABEL = "空欄に戻す";
const CALENDAR_ID_HEADER = "カレンダーID";
const SUMMARY_DIRTY_KEY = "SUMMARY_DIRTY";
let IS_BATCH_SETUP_ = false;

// 社用版では、担当者・既読確認者は設定シートを正とする。
// ここにサンプル名を入れると、設定シート修復時に山田/鈴木などが戻るため空にする。
const DEFAULT_STAFF_MEMBERS = [];
const DEFAULT_PERSONAL_MEMBERS = [];

// 旧サンプル名を設定シートから手動削除するための安全な掃除リスト。
// 実在社員に同名がいる場合は、このメニューを実行せず、設定シートを手で確認する。
const SAMPLE_STAFF_MEMBERS_FOR_CLEANUP = [
  "山田", "鈴木", "田中", "高橋", "伊藤", "中村", "小林", "吉田", "山本", "佐々木", "松本", "井上", "木村", "林", "清水"
];
const SAMPLE_PERSONAL_MEMBERS_FOR_CLEANUP = ["山田", "高橋", "鈴木"];

const DEFAULT_COMPANY_VEHICLES = [
  "サンプル社用車A",
  "サンプル社用車B",
  "サンプル作業車A",
  "サンプル作業車B",
  "サンプルトラックA",
  "サンプルトラックB",
  "サンプル軽車両A",
  "サンプル軽車両B",
  "サンプル予備車A",
  "サンプル予備車B"
];

const MIN_DATA_ROWS_FOR_VALIDATION = 300;
// p33: 資格管理は行挿入時に入力規則が外れやすいため、
// 通常の最終行＋少し先だけでなく、最低500行までは名前/資格名プルダウンを維持する。
// 全列・全1000行へ毎回貼ると重くなるため、500行＋最終行から200行先を上限にする。
const QUALIFICATION_MANAGEMENT_DROPDOWN_MIN_ROWS = 200;
const QUALIFICATION_MANAGEMENT_DROPDOWN_BUFFER_ROWS = 80;
const QUALIFICATION_MANAGEMENT_ONEDIT_BUFFER_ROWS = 30;
const AUTO_EXTEND_ROWS_BUFFER = 100;
const AUTO_EXTEND_TRIGGER_MARGIN = 30;

const DATE_HEADERS = [
  "日付", "投稿日", "登録日", "購入日", "契約日", "開始日", "終了日", "車検期限", "車検予定日", "新車検期限", "取得日", "更新期限", "免許証交付日", "免許証有効期限", "修理依頼日", "返却予定日", "作成日時", "対応日", "移動日", "期限"
];

const TIME_HEADERS = ["時間", "開始時刻", "終了時刻"];
// 数字だけでも先頭0を保持したい列。電話番号・車番・工事連絡先は文字列固定にする。
const VEHICLE_NUMBER_HEADERS = ["車番", "折り返し電話番号", "連絡先"];
const NUMBER_HEADERS = ["金額"];

const LICENSE_TYPE_HEADERS = ["普通", "準中型", "中型", "大型", "大型特殊", "けん引", "二種"];
const DRIVER_LICENSE_ISSUE_DATE_HEADER = "免許証交付日";
const DRIVER_LICENSE_EXPIRY_DATE_HEADER = "免許証有効期限";
const DRIVER_LICENSE_DETAIL_SHEET_NAME = "運転免許明細";
const DRIVER_LICENSE_UPDATE_HISTORY_SHEET_NAME = "運転免許更新履歴";
const QUALIFICATION_UPDATE_HISTORY_SHEET_NAME = "資格更新履歴";

// p42: 運転免許管理は紙フォルダ照合で下の空行まで入力するため、
// 20行以降だけ行高・折り返しが戻らない問題を防ぐ。
// 全1000行へ毎回書式をかけると重くなるため、通常は最低120行までに限定する。
const DRIVER_LICENSE_DISPLAY_MIN_ROWS = 120;
const DRIVER_LICENSE_DETAIL_DISPLAY_MIN_ROWS = 120;

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
  "運転免許明細": "免許明細ID",
  "運転免許更新履歴": "免許更新履歴ID",
  "資格管理": "資格ID",
  "資格更新履歴": "資格更新履歴ID",
  "資格マスタ": "資格マスタID",
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
  "運転免許明細",
  "運転免許更新履歴",
  "資格管理",
  "資格更新履歴",
  "資格マスタ",
  "資格表記ゆれマスタ",
  "資格かんたん登録",
  "資格一括登録",
  "資格重複チェック",
  "社員別資格一覧",
  "資格別保有者一覧",
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
  // v9.7.48p18: 日常運用で触る「設定」「資格マスタ」「ダッシュボード」は隠さない。
  "SQL設計", "移行対応表", "SQLサンプル集",
  "帳簿PDF履歴", "帳簿出力設定",
  "担当別未読", "既読率集計",
  "手順シート", "社内管理説明",
  "機能チェック", "GPT診断用", "自動テスト結果",
  "資格重複チェック", "資格表記ゆれマスタ", "資格プルダウン補助"
];

/**
 * v9.7.17 注意:
 * - Apps Script 側に古い .gs ファイルが残っていると、古い onOpen() が実行されて
 *   メニューが「日常更新・集計更新・GPT診断・PDF」だけに戻ることがあります。
 * - このコードは「1ファイルだけ残して全文貼り替え」する前提です。
 * - onOpen() はこの1個だけにしてください。
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // v9.7.48p18:
  // 会社共有時に迷わないよう、通常メニューは日常運用で押すものだけに絞る。
  // 詳細修復、自動テスト、サンプル、移行、削除系は下部の保守/危険メニューへ隔離する。

  ui.createMenu("社内管理")
    .addItem("バージョン確認", "showSystemVersion")
    .addSeparator()

    // 日常運用で主に使う項目。
    .addItem("日常更新（一覧＋要確認）", "refreshDailyOnly")
    .addItem("集計更新（未読・既読率・DB）", "refreshSummaryOnly")
    .addItem("GPT診断用シートを作成", "createGptDiagnosticSheet")
    .addSeparator()

    // 会社用の軽量修復。迷いやすい詳細メニューはサブメニューへ移動。
    .addItem("会社用：ベータ前チェックパック", "setupCompanyBetaPreflightPack")
    .addItem("会社用：資格管理だけ整える（超軽量）", "setupCompanyQualificationPack")
    .addItem("会社用：ID列・カレンダーID列を隠す", "hideSystemColumnsForAllInputSheets")
    .addItem("会社用：ダッシュボードを作り直す", "repairDashboardSheet")
    .addItem("会社用：裏方シートを整理", "setupCompanyHiddenSheetsPack")
    .addSeparator()

    .addSubMenu(
      ui.createMenu("保守用：よく使う修復")
        .addItem("会社用：初回導入パック", "setupCompanyDeploymentPack")
        .addItem("会社用：車検・工事・備品・行事を整える", "setupCompanyOperationPack")
        .addItem("会社用：修理業者マスタ＋プルダウン修復", "repairRepairVendorMasterAndDropdown")
        .addItem("会社用：サンプル担当者・既読列を完全整理", "cleanupCompanySampleStaffAndPersonalColumns")
        .addSeparator()
        .addItem("入力シートだけ設定反映（超軽量）", "refreshInputSheets")
        .addItem("入力シート設定反映（完全・重い）", "refreshInputSheetsFullMaintenance")
        .addItem("免許・資格の更新履歴シートを作成/修復", "repairLicenseAndQualificationUpdateHistorySheets")
        .addItem("データ整合性チェック・軽微修正", "runDataConsistencyCheckAndFix")
        .addItem("機能チェックシートを作成", "createFeatureCheckSheet")
        .addItem("裏方シートを作成/修復", "ensureSupportSheets")
        .addItem("手順シートを作成/更新", "createProcedureSheet")
    )

    .addSubMenu(
      ui.createMenu("保守用：資格管理")
        .addItem("資格管理だけ整える（超軽量・通常はこちら）", "setupCompanyQualificationPack")
        .addItem("社員別・資格別資格一覧だけ更新", "setupCompanyQualificationListPack")
        .addItem("資格管理の列ズレ＋区分別プルダウン復旧（重め）", "repairQualificationHeaderDropdownSortSafe")
        .addItem("資格管理の資格名プルダウンだけ直す（行ごと厳密）", "repairQualificationManagementCategoryNameSort")
        .addItem("資格管理の名前・資格名プルダウンを復旧", "repairQualificationManagementDropdowns")
        .addItem("資格管理：行挿入後プルダウンを復旧", "repairQualificationManagementDropdownsAfterRowInsert")
        .addItem("資格管理：行挿入検知トリガーを設定", "installQualificationManagementRowInsertTrigger")
        .addItem("資格管理のプルダウンを完全復旧", "rebuildQualificationManagementDependentDropdowns")
        .addItem("資格管理の既読チェックを軽量再設定", "refreshQualificationManagementCheckboxes")
        .addItem("資格管理のFALSE表示を修復", "repairQualificationManagementFalseDisplay")
        .addItem("資格管理を所有者＋マスタ順に整理（空欄下）", "sortQualificationManagementByOwner")
        .addItem("資格管理を紙マスタ順に並べ替え（軽量）", "sortQualificationManagementByPaperOrderLight")
        .addItem("資格管理の所有者グループを完全作り直し", "groupQualificationManagementByOwner")
        .addSeparator()
        .addItem("資格マスタ通知日前列を修復", "repairQualificationMasterNoticeAndAlias")
        .addItem("資格マスタを紙の標準に戻す（手動変更上書き）", "rebuildQualificationMasterClean")
        .addItem("資格プルダウン補助を作成/更新", "ensureQualificationDropdownHelperSheet")
        .addItem("資格診断ヘッダーを軽量修復", "repairQualificationDiagnosticHeaders")
        .addItem("資格管理カラーリングを再調整", "formatQualificationManagementColors")
        .addSeparator()
        .addItem("資格かんたん登録シートを作成/修復", "createQualificationEasyPasteSheet")
        .addItem("資格かんたん登録の区分別プルダウンを再構築", "rebuildQualificationEasyPasteDependentDropdowns")
        .addItem("資格かんたん登録を資格管理へ反映（新規のみ）", "applyQualificationEasyPasteToManagement")
        .addItem("資格かんたん登録で既存資格も更新（空欄は無視）", "applyQualificationEasyPasteUpdateExisting")
        .addItem("資格かんたん登録をクリア（ヘッダー残す）", "clearQualificationEasyPasteSheet")
        .addSeparator()
        .addItem("社員別資格一覧を作成/更新（グループなし）", "createEmployeeQualificationListSheet")
        .addItem("資格別保有者一覧を作成/更新（任意）", "createQualificationHolderListSheet")
        .addItem("資格管理の重複チェック", "createQualificationDuplicateCheckSheet")
        .addSeparator()
        .addItem("資格更新履歴を作成/修復", "createQualificationUpdateHistorySheet")
        .addItem("資格更新履歴を新しい順に整理", "sortQualificationUpdateHistorySheet")
    )

    .addSubMenu(
      ui.createMenu("保守用：車検・工事・電話など")
        .addItem("車検更新済を処理（履歴＋期限更新）", "clearUpdatedVehicleStatuses")
        .addItem("車検履歴列＋車両名プルダウンを修復", "repairVehicleHistoryAndVehicleNameDropdowns")
        .addItem("車検履歴の日付・担当入力規則を修復", "repairVehicleHistoryValidationOnly")
        .addItem("選択行の車検履歴を車検管理から再同期", "syncSelectedVehicleHistoryRowFromVehicleManagement")
        .addItem("設定シートに社用車一覧を追加/修復", "repairCompanyVehicleSettings")
        .addSeparator()
        .addItem("工事予定の未契約・通知完了・色を修復", "repairConstructionScheduleDropdownAndColors")
        .addItem("工事予定の契約金額・税・状態・連絡先を修復", "repairConstructionContractAmountAndTax")
        .addSeparator()
        .addItem("電話履歴を未完了上・完了下に整理", "sortPhoneHistorySheet")
        .addItem("運転免許シートの表示・行高調整", "alignLicenseSheetLayout")
        .addItem("運転免許ヘッダー・明細を修復（軽量）", "repairDriverLicenseSheetsLight")
        .addItem("運転免許管理を所有者50音順に整理", "sortDriverLicenseManagementByOwnerKana")
        .addItem("運転免許明細を作成/修復", "createDriverLicenseDetailSheet")
        .addItem("運転免許明細を所有者50音＋免許種類順に整理", "sortDriverLicenseDetailByOwnerKana")
        .addItem("運転免許更新履歴を作成/修復", "createDriverLicenseUpdateHistorySheet")
        .addItem("運転免許更新履歴を新しい順に整理", "sortDriverLicenseUpdateHistorySheet")
        .addItem("フィードバック列追加・順番修復", "repairFeedbackColumnUpdates")
        .addItem("行事予定の備考手前・内容非表示を修復", "repairEventScheduleColumns")
    )

    .addSubMenu(
      ui.createMenu("保守用：表示・帳票")
        .addItem("今のシートだけ表示調整", "formatActiveSheetOnly")
        .addItem("今のシートだけコンパクト表示", "compactActiveSheet")
        .addItem("今のシートだけ標準表示に戻す", "standardizeActiveSheetReadable")
        .addItem("全入力シートを標準表示に整える（少し大きめ）", "standardizeAllInputSheetsReadable")
        .addItem("今のシートの非表示列をすべて表示", "showAllColumnsForActiveSheet")
        .addItem("全入力シートのID列・カレンダーID列を非表示", "hideSystemColumnsForAllInputSheets")
        .addSeparator()
        .addItem("裏方シートを非表示", "hideSupportSheets")
        .addItem("裏方シートを表示", "showSupportSheets")
        .addItem("一覧スケジュールを月別グループ化", "groupScheduleByMonth")
        .addItem("過去一覧を月別グループ化", "groupArchiveByMonth")
        .addItem("今のシートを月別グループ化", "groupActiveSheetByMonth")
        .addSeparator()
        .addItem("選択行の日報PDFを作成", "createDailyReportPdfForActiveRow")
        .addItem("一覧帳簿PDFを作成", "createLedgerPdf")
    )

    .addSubMenu(
      ui.createMenu("保守用：自動テスト")
        .addItem("自動テストパック（使い方）", "setupCompanyAutoTestPack")
        .addItem("自動テスト①入力だけ", "setupCompanyAutoTestInputOnly")
        .addItem("自動テスト②更新チェック", "setupCompanyAutoTestUpdateOnly")
        .addItem("自動テスト③GPT診断", "setupCompanyAutoTestDiagnosticOnly")
        .addItem("テストデータ削除", "deleteCompanyAutoTestData")
    )

    .addSubMenu(
      ui.createMenu("危険操作・通常使用禁止")
        .addItem("全体更新（保守用・重め）", "refreshAll")
        .addItem("書式込み全体更新（かなり重い）", "refreshAllWithFormat")
        .addItem("旧列名をv9.7.0へ移行", "migrateLegacyColumnsV970")
        .addItem("新規ベータ作成（危険・既存データ注意）", "setupNewBetaClean")
        .addItem("ベータ確認準備（軽量）", "prepareBetaCheckLight")
        .addSeparator()
        .addItem("サンプル追加：主要予定系", "addSampleMainSchedules")
        .addItem("サンプル追加：車検・免許・資格・備品", "addSampleVehicleLicenseEquipment")
        .addItem("サンプル追加：日報・レシート・FB", "addSampleDailyReceiptFeedback")
        .addItem("ベータ用サンプルを削除（注意）", "deleteBetaSamples")
        .addSeparator()
        .addItem("資格かんたん登録で強制上書き（注意）", "applyQualificationEasyPasteForceOverwrite")
        .addItem("横長資格入力シートを非表示", "hideOldQualificationInputSheets")
        .addItem("横長資格入力シートを削除（注意）", "deleteOldQualificationInputSheets")
        .addItem("車検管理を修復（車番文字列＋色）", "repairVehicleInspectionSheet")
        .addItem("過去予定を過去一覧へ移動", "movePastItemsToArchive")
        .addItem("月次整理（過去移動＋月別）", "monthlyMaintenance")
    )

    .addToUi();

  // p29: 開いた時点でも内部ID列を通常表示へ戻さない。
  // 失敗してもメニュー表示を止めない。
  try { hideSystemColumnsForAllInputSheets_({silent: true}); } catch (e) { console.log("onOpen hide system columns skip: " + e.message); }
}

function showSystemVersion() {
  toast_("社内共有業務管理システム " + SYSTEM_VERSION + " が動いています");
}

function acquireCompanyPackLock_(label) {
  const lock = LockService.getDocumentLock();
  try {
    // 通常の30秒 tryLock だと、GPT診断や前回処理の直後に「他の処理中です」が出やすい。
    // 会社用パックは手動で押す処理なので、最大3分だけ待ってから開始する。
    lock.waitLock(180000);
    return lock;
  } catch (e) {
    toast_(label + "は他の処理中です。1〜3分待ってから再実行してください。");
    console.log(label + " lock wait failed: " + e.message);
    return null;
  }
}

function releaseCompanyPackLock_(lock) {
  try {
    if (lock) lock.releaseLock();
  } catch (e) {
    console.log("会社用パックのロック解放をスキップ: " + e.message);
  }
}

function setupCompanyDeploymentPack() {
  try {
    ensureSettingsSheet_();
    rebuildPersonalCheckColumnsFromSettings_();
    setupCompanyQualificationPack();
    setupCompanyOperationPack();
    createFeatureCheckSheet();
    toast_("会社用初回導入パックが完了しました。機能チェックシートを確認してください。");
  } catch (err) {
    toast_("会社用初回導入パックでエラー: " + err.message);
    throw err;
  }
}


function setupCompanyBetaPreflightPack() {
  const lock = acquireCompanyPackLock_("ベータ前チェックパック");
  if (!lock) return;

  try {
    clearSettingsCache_();
    ensureSettingsSheet_();

    // ベータ前チェックは「安全な軽量修復」だけをまとめる。
    // 新規ベータ作成、旧列名移行、書式込み全体更新、過去一覧移動などは含めない。
    ensureV9714MinorLayoutUpdates_();
    repairQualificationDiagnosticHeaders_();
    repairEventScheduleColumns_();
    repairVehicleHistoryAndVehicleNameDropdowns_();
    repairConstructionContractAmountAndTax_();
    rebuildRepairVendorMasterFromEquipmentSheet_();
    repairRepairVendorDropdown_();

    // 設定シートの担当者・既読確認者・社用車・修理業者などを入力シートへ反映。
    refreshInputSheets();

    // 空行に残ったIDだけを軽く整理。確認ダイアログ付きの非常用一括メニューは使わない。
    const idCleanup = cleanInputSheetsOrphanIdsLight_();

    // 資格管理はチェックボックス表示が崩れやすいため、最後に専用処理でもう一度整える。
    const qualificationCheck = refreshQualificationManagementCheckboxes_();

    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    clearSummaryDirty_();

    // 最後にGPT診断用シートを作成する。
    // createGptDiagnosticSheet 内でも資格管理チェックを再確認する。
    createGptDiagnosticSheet();

    toast_(
      "ベータ前チェックパックが完了しました。GPT診断用シートを確認してください。"
      + " ID整理:" + idCleanup.cleanedIds
      + " / 資格チェック:" + qualificationCheck.checkboxCount + "個"
    );
  } catch (err) {
    toast_("ベータ前チェックパックでエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}



/**
 * 会社用：自動テストパック。
 * p10では時間超過防止のため、一括実行せず3分割で実行する。
 */
function setupCompanyAutoTestPack() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "自動テストパックの使い方",
    "起動時間の最大値を超えないよう、p10では自動テストを3分割しています。\n\n"
    + "1) 会社用：自動テスト①入力だけ\n"
    + "2) 会社用：自動テスト②更新チェック\n"
    + "3) 会社用：自動テスト③GPT診断\n\n"
    + "確認後は『会社用：テストデータ削除』を実行してください。",
    ui.ButtonSet.OK
  );
}


/**
 * 会社用：自動テスト①入力だけ。
 * AUTO_TEST_ 行を各入力シートへ入れるだけ。重い一覧更新・GPT診断は行わない。
 */
function setupCompanyAutoTestInputOnly() {
  const lock = acquireCompanyPackLock_("自動テスト①入力だけ");
  if (!lock) return;

  try {
    clearSettingsCache_();
    ensureSettingsSheet_();

    // 前回のAUTO_TEST_行が残っていると判定がぶれるため、先に軽量削除する。
    const beforeDelete = deleteCompanyAutoTestData_({silent: true, skipRefresh: true});

    // 入力に必要な軽量修復だけ。ここでは一覧更新・GPT診断はしない。
    ensureV9714MinorLayoutUpdates_();
    repairEventScheduleColumns_();
    repairVehicleHistoryAndVehicleNameDropdowns_();
    repairConstructionContractAmountAndTax_();
    repairRepairVendorDropdown_();

    const marker = buildCompanyAutoTestMarker_();
    saveCompanyAutoTestMarker_(marker);

    const writeResult = writeCompanyAutoTestRows_(marker);

    // 車検更新済テストはAUTO_TEST_行だけを履歴化する。
    const vehicleResult = processCompanyAutoTestVehicleUpdate_(marker);

    // 修理業者はAUTO_TEST_行を入れたあとに設定へ拾う。
    const addedVendors = rebuildRepairVendorMasterFromEquipmentSheet_();
    sortRepairVendorSettingsRows_();
    repairRepairVendorDropdown_();

    const checkRows = [
      autoTestRow_("情報", "実行ステップ", "OK", "①入力だけ"),
      autoTestRow_("情報", "テストマーカー", "OK", marker),
      autoTestRow_("事前整理", "前回テストデータ削除", "情報", "削除行: " + beforeDelete.deletedRows),
      autoTestRow_("自動入力", "入力シート投入", writeResult.totalRows > 0 ? "OK" : "NG", "追加行: " + writeResult.totalRows),
      autoTestRow_("車検", "更新済→車検履歴", vehicleResult.added > 0 ? "OK" : "要確認", "履歴追加: " + vehicleResult.added + " / 状態空欄戻し: " + vehicleResult.cleared + (vehicleResult.reason ? " / " + vehicleResult.reason : "")),
      autoTestRow_("修理業者", "手入力→設定シート反映", addedVendors >= 0 ? "情報" : "要確認", "追加処理: " + addedVendors),
      autoTestRow_("次の操作", "自動テスト②更新チェック", "案内", "次に『会社用：自動テスト②更新チェック』を実行してください")
    ];
    createCompanyAutoTestResultSheet_(checkRows);

    toast_(
      "自動テスト①入力だけが完了しました。次に『自動テスト②更新チェック』を実行してください。"
      + " 追加:" + writeResult.totalRows
      + " / 車検履歴:" + vehicleResult.added
    );
  } catch (err) {
    toast_("自動テスト①入力だけでエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}


/**
 * 会社用：自動テスト②更新チェック。
 * ①で入れたAUTO_TEST_行を使い、一覧・要確認・ID・既読を確認する。
 * GPT診断は時間超過防止のため③へ分離。
 */
function setupCompanyAutoTestUpdateOnly() {
  const lock = acquireCompanyPackLock_("自動テスト②更新チェック");
  if (!lock) return;

  try {
    clearSettingsCache_();
    const marker = getCompanyAutoTestMarker_();
    if (!marker) {
      toast_("AUTO_TEST_ 行が見つかりません。先に『自動テスト①入力だけ』を実行してください。");
      return;
    }

    // 入力規則・ID・既読を軽く再設定。
    refreshInputSheets();
    refreshQualificationManagementCheckboxes_();

    // 一覧・要確認・集計を作る。GPT診断は別。
    createScheduleList();
    createAlertList();
    createAssigneeUnreadSummary();
    createReadRateSummary();
    createDashboard();
    clearSummaryDirty_();

    const inputNames = getInputSheetNames_();
    const created = inputNames.map(name => ({sheetName: name, count: countRowsContainingMarkerByName_(name, marker)}));
    const totalRows = created.reduce((sum, item) => sum + item.count, 0);
    const writeResult = {marker: marker, created: created, totalRows: totalRows, context: getCompanyAutoTestContext_(marker)};
    const vehicleResult = {
      added: countRowsContainingMarkerByName_("車検履歴", marker),
      cleared: 0,
      reason: "②更新チェック時点の履歴行数"
    };
    const vendorCount = countRowsContainingMarkerByName_("設定", marker);

    const checkRows = buildCompanyAutoTestCheckRows_(marker, writeResult, vehicleResult, vendorCount, 0);
    checkRows.push(autoTestRow_("次の操作", "自動テスト③GPT診断", "案内", "必要なら次に『会社用：自動テスト③GPT診断』を実行してください"));
    createCompanyAutoTestResultSheet_(checkRows);

    toast_(
      "自動テスト②更新チェックが完了しました。自動テスト結果を確認してください。"
      + " 判定:" + countCompanyAutoTestNg_(checkRows) + "件NG"
    );
  } catch (err) {
    toast_("自動テスト②更新チェックでエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}


/**
 * 会社用：自動テスト③GPT診断。
 * GPT診断用シート作成だけを単独実行する。
 */
function setupCompanyAutoTestDiagnosticOnly() {
  const lock = acquireCompanyPackLock_("自動テスト③GPT診断");
  if (!lock) return;

  try {
    createGptDiagnosticSheet();
    toast_("自動テスト③GPT診断が完了しました。GPT診断用シートを確認してください。");
  } catch (err) {
    toast_("自動テスト③GPT診断でエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}


/**
 * 会社用：テストデータ削除。
 * p10では削除を軽量化しつつ、AUTO_TEST_ と過去版フィードバック残骸をまとめて削除する。
 */
function deleteCompanyAutoTestData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "テストデータ削除",
    "AUTO_TEST_ を含む自動テスト行を削除します。続行しますか？\n\n"
    + "p10では時間超過防止のため、削除後の一覧更新・GPT診断は自動では行いません。\n"
    + "過去版で残りやすかったフィードバックの自動テスト行も削除対象にします。",
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const lock = acquireCompanyPackLock_("テストデータ削除");
  if (!lock) return;

  try {
    const result = deleteCompanyAutoTestData_({silent: true, skipRefresh: true});
    saveCompanyAutoTestMarker_("");

    const detail = Object.keys(result.deletedBySheet || {})
      .map(name => name + ":" + result.deletedBySheet[name])
      .join(" / ") || "対象なし";

    const rows = [
      autoTestRow_("削除", "AUTO_TEST_ データ", "OK", "削除行: " + result.deletedRows),
      autoTestRow_("削除", "シート別削除", "情報", detail),
      autoTestRow_("削除", "残骸チェック", result.remainingAutoTestRows === 0 ? "OK" : "要確認", "残り: " + result.remainingAutoTestRows),
      autoTestRow_("次の操作", "日常更新（一覧＋要確認）", "案内", "必要なら日常更新を実行してください"),
      autoTestRow_("次の操作", "GPT診断用シートを作成", "案内", "必要ならGPT診断を作り直してください")
    ];
    createCompanyAutoTestResultSheet_(rows);

    toast_("AUTO_TEST_ テストデータを削除しました（削除行: " + result.deletedRows + " / 残り:" + result.remainingAutoTestRows + "）。必要なら日常更新とGPT診断を別で実行してください。");
  } catch (err) {
    toast_("テストデータ削除でエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}


function saveCompanyAutoTestMarker_(marker) {
  try {
    const props = PropertiesService.getDocumentProperties();
    if (marker) {
      props.setProperty("COMPANY_AUTO_TEST_LAST_MARKER", marker);
    } else {
      props.deleteProperty("COMPANY_AUTO_TEST_LAST_MARKER");
    }
  } catch (e) {}
}


function getCompanyAutoTestMarker_() {
  try {
    const saved = PropertiesService.getDocumentProperties().getProperty("COMPANY_AUTO_TEST_LAST_MARKER");
    if (saved && countRowsContainingMarkerByName_("一覧スケジュール", saved) + countRowsContainingMarkerAcrossInputSheets_(saved) > 0) {
      return saved;
    }
  } catch (e) {}
  return findLatestCompanyAutoTestMarker_();
}


function countRowsContainingMarkerAcrossInputSheets_(marker) {
  return getInputSheetNames_().reduce((sum, name) => sum + countRowsContainingMarkerByName_(name, marker), 0);
}


function findLatestCompanyAutoTestMarker_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = uniqueList_(getInputSheetNames_().concat(["車検履歴", "一覧スケジュール", "要確認一覧"]));
  const re = /AUTO_TEST_\d{8}_\d{6}/g;
  const found = [];

  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
    values.forEach(row => {
      row.forEach(value => {
        const text = String(value || "");
        const matches = text.match(re);
        if (matches) matches.forEach(m => found.push(m));
      });
    });
  });

  if (!found.length) return "";
  const unique = uniqueList_(found).sort();
  return unique[unique.length - 1];
}


function buildCompanyAutoTestMarker_() {
  const tz = Session.getScriptTimeZone();
  return AUTO_TEST_MARKER_PREFIX + Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");
}


function getCompanyAutoTestContext_(marker) {
  clearSettingsCache_();
  const staff = getStaffMembers_();
  const vehicles = getCompanyVehicles_();
  const vehicleNames = getVehicleNames_();
  const qualification = getFirstQualificationForAutoTest_();

  const assignee1 = staff[0] || "";
  const assignee2 = staff[1] || assignee1;
  const assignee3 = staff[2] || assignee1;
  const vehicle = vehicles[0] || "";
  const vehicleName = vehicleNames[0] || vehicle || "";

  return {
    marker: marker,
    assignee1: assignee1,
    assignee2: assignee2,
    assignee3: assignee3,
    vehicle: vehicle,
    vehicleName: vehicleName,
    qualificationCategory: qualification.category,
    qualificationName: qualification.name,
    repairVendor: marker + "_修理業者",
    vehicleNumber: marker + "-車番"
  };
}


function getFirstQualificationForAutoTest_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("資格マスタ");
  if (sheet && sheet.getLastRow() >= 2) {
    const headers = getHeaders_(sheet);
    const categoryCol = headers.indexOf("区分");
    const nameCol = headers.indexOf("資格名");
    const visibleCol = headers.indexOf("一覧表示");
    if (nameCol >= 0) {
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
      for (let i = 0; i < values.length; i++) {
        const name = String(values[i][nameCol] || "").trim();
        const category = categoryCol >= 0 ? String(values[i][categoryCol] || "").trim() : "";
        const visible = visibleCol >= 0 ? String(values[i][visibleCol] || "").trim() : "";
        if (name && visible !== "しない" && visible !== "非表示") {
          return {category: category || "技能講習", name: name};
        }
      }
    }
  }

  return {category: "技能講習", name: "玉掛け技能講習"};
}


function writeCompanyAutoTestRows_(marker) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const c = getCompanyAutoTestContext_(marker);
  const created = [];

  setupSheetsIfMissing_(getInputSheetNames_().concat(["車検履歴"]));

  const rowsBySheet = {
    "出先予定": [
      {"日付": addDays_(today, 1), "行き先": marker + "_出先", "用件": "自動テスト用件", "担当": c.assignee1, "社用車": c.vehicle, "状態": "予定", "電話対応": "", "備考": marker}
    ],
    "工事予定": [
      {"工事名": marker + "_工事", "現場": marker + "_現場", "依頼主": marker + "_依頼主", "連絡先": "000-0000-0000", "契約日": today, "契約金額": 123456, "税": "税込", "開始日": addDays_(today, 2), "終了日": addDays_(today, 5), "状態": "予定", "担当": c.assignee1, "電話対応": "未対応", "備考": marker}
    ],
    "会議予定": [
      {"日付": addDays_(today, 1), "開始時刻": "10:00", "終了時刻": "10:30", "会議名": marker + "_会議", "内容": "自動テスト会議", "担当": c.assignee2, "状態": "予定", "資料": "", "備考": marker}
    ],
    "行事予定": [
      {"日付": addDays_(today, 1), "終了日": addDays_(today, 2), "開始時刻": "09:00", "終了時刻": "12:00", "行事名": marker + "_行事", "担当": c.assignee2, "状態": "予定", "備考": marker}
    ],
    "作業状況": [
      {"現場": marker + "_作業現場", "作業内容": "自動テスト作業", "状態": "施工中", "担当": c.assignee1, "写真": "", "備考": marker}
    ],
    "電話履歴": [
      {"日付": today, "時間": "09:15", "相手": marker + "_電話相手", "折り返し電話番号": "000-0000-0000", "内容": "自動テスト折返し", "電話対応": "折返し", "担当": c.assignee3, "対応メモ": marker, "備考": marker}
    ],
    "車検管理": [
      {"車両名": c.vehicleName, "車番": c.vehicleNumber, "車検期限": addDays_(today, 10), "車検予定日": addDays_(today, 1), "新車検期限": addDays_(today, 370), "保険期限": addDays_(today, 30), "車検状態": "更新済", "担当": c.assignee1, "写真": "", "備考": marker}
    ],
    "社用車予約": [
      {"日付": addDays_(today, 1), "開始時刻": "08:30", "終了時刻": "17:00", "社用車": c.vehicle, "利用者": c.assignee1, "行き先": marker + "_行き先", "用途": "自動テスト", "状態": "予定", "備考": marker}
    ],
    "日報": [
      {"日付": today, "入力者": c.assignee1, "担当": c.assignee1, "現場": marker + "_日報現場", "作業内容": "自動テスト作業内容", "進捗": "予定通り", "問題点": "なし", "明日の予定": "自動テスト予定", "他現場状況": "", "写真": "", "日報文章": "", "状態": "下書き", "備考": marker}
    ],
    "日報レシート管理": [
      {"日付": today, "担当": c.assignee1, "現場": marker + "_日報現場", "支払先": marker + "_支払先", "内容": "自動テスト金額", "区分": "その他", "金額": 1111, "支払方法": "現金", "レシート写真": "", "経理確認": "未確認", "精算状態": "未精算", "備考": marker}
    ],
    "運転免許管理": [
      {"所有者": c.assignee1, "普通": true, "準中型": false, "中型": false, "大型": false, "大型特殊": false, "けん引": false, "二種": false, "免許証交付日": addDays_(today, -1000), "免許証有効期限": addDays_(today, 5), "コピー有無": "未確認", "状態": "更新予定", "備考": marker}
    ],
    "資格管理": [
      {"所有者": c.assignee1, "区分": c.qualificationCategory, "資格名": c.qualificationName, "取得区分": "保有", "保有状況": "保有", "取得日": addDays_(today, -500), "更新期限": addDays_(today, 5), "コピー有無": "未確認", "状態": "更新予定", "備考": marker}
    ],
    "備品修理管理": [
      {"購入日": addDays_(today, -100), "修理依頼日": today, "返却予定日": addDays_(today, 3), "修理依頼者": c.assignee2, "修理業者": c.repairVendor, "備品名": marker + "_備品", "内容": "自動テスト修理", "返却済": "未", "状態": "修理中", "備考": marker}
    ],
    "お知らせ": [
      {"投稿日": today, "タイトル": marker + "_お知らせ", "内容": "自動テストお知らせ", "投稿者": c.assignee1, "重要度": "高", "備考": marker}
    ],
    "個人ToDo": [
      {"登録日": today, "担当": c.assignee1, "内容": marker + "_ToDo", "期限": addDays_(today, 1), "状態": "未着手", "備考": marker}
    ],
    "フィードバック": [
      {"日付": today, "記入者": c.assignee1, "確認方法": "PC", "対象シート": "工事予定", "気になった内容": marker + "_フィードバック", "区分": "動作不備", "困り度": "低", "対応状況": "未対応", "対応方針": "自動テスト", "対応メモ": marker, "対応日": ""}
    ]
  };

  Object.keys(rowsBySheet).forEach(sheetName => {
    const count = appendCompanyAutoTestObjects_(sheetName, rowsBySheet[sheetName]);
    created.push({sheetName: sheetName, count: count});
  });

  const totalRows = created.reduce((sum, item) => sum + item.count, 0);
  return {marker: marker, created: created, totalRows: totalRows, context: c};
}


function appendCompanyAutoTestObjects_(sheetName, objects) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const headers = getSheetHeaders_()[sheetName];
    if (!headers) return 0;
    sheet = ss.insertSheet(sheetName);
    setupSheet_(sheet, headers);
  }

  const headers = getHeaders_(sheet);
  if (!headers.length || !objects || !objects.length) return 0;

  const personal = getPersonalMembers_();
  const rows = objects.map(obj => headers.map(header => {
    if (header === READ_HEADER || personal.includes(header) || LICENSE_TYPE_HEADERS.includes(header)) return obj[header] === true;
    return obj[header] !== undefined ? obj[header] : "";
  }));

  const startRow = Math.max(sheet.getLastRow() + 1, 2);
  const requiredRows = startRow + rows.length - 1;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
  ensureIdsForSheet_(sheet);
  updateNoticeColumnForSheet_(sheet);
  setCheckboxesForDataRows(sheet);
  // p8: 自動テスト投入時に毎シートの条件付き書式を再作成すると時間超過しやすいため、
  // 色ルールは既存のものを使い、必要な場合だけ別メニューで表示調整する。
  return rows.length;
}


function processCompanyAutoTestVehicleUpdate_(marker) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("車検管理");
  if (!sheet || sheet.getLastRow() < 2) return {added: 0, cleared: 0, reason: "車検管理なし"};

  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf("車検状態") + 1;
  const dueCol = headers.indexOf("車検期限") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const rowNumbers = findRowsContainingMarker_(sheet, marker).map(item => item.rowNumber);

  let added = 0;
  let cleared = 0;
  let lastReason = "";

  rowNumbers.forEach(row => {
    const result = processVehicleInspectionUpdatedRow_(sheet, row);
    if (result.added) {
      added++;
      if (statusCol > 0) {
        sheet.getRange(row, statusCol).clearContent();
        cleared++;
      }
    }
    if (result.reason) lastReason = result.reason;

    try {
      if (noticeCol > 0 && dueCol > 0) {
        const newDue = toDateOnly_(sheet.getRange(row, dueCol).getValue());
        sheet.getRange(row, noticeCol).setValue(newDue ? getNoticeText(newDue) : "");
      }
    } catch (e) {}
  });

  try { sortVehicleInspectionSheetByDue_(); } catch (e) {}
  return {added: added, cleared: cleared, reason: lastReason};
}


function buildCompanyAutoTestCheckRows_(marker, writeResult, vehicleResult, addedVendors, deletedBefore) {
  const rows = [];
  const inputNames = getInputSheetNames_();

  rows.push(autoTestRow_("情報", "テストマーカー", "OK", marker));
  rows.push(autoTestRow_("事前整理", "前回テストデータ削除", "情報", "削除行: " + deletedBefore));
  rows.push(autoTestRow_("自動入力", "入力シート投入", writeResult.totalRows >= inputNames.length ? "OK" : "要確認", "追加行: " + writeResult.totalRows));

  inputNames.forEach(name => {
    const count = countRowsContainingMarkerByName_(name, marker);
    rows.push(autoTestRow_("自動入力", name, count > 0 ? "OK" : "NG", "AUTO_TEST_ 行: " + count));
  });

  inputNames.forEach(name => {
    const idResult = checkCompanyAutoTestIds_(name, marker);
    if (idResult.checked > 0) {
      rows.push(autoTestRow_("ID", name, idResult.missing === 0 ? "OK" : "NG", "確認: " + idResult.checked + " / ID不足: " + idResult.missing));
    }
  });

  inputNames.forEach(name => {
    const checkResult = checkCompanyAutoTestCheckboxes_(name, marker);
    if (checkResult.checked > 0) {
      rows.push(autoTestRow_("既読チェック", name, checkResult.missing === 0 ? "OK" : "NG", "確認: " + checkResult.checked + " / 不足: " + checkResult.missing));
    }
  });

  const scheduleTargets = ["出先予定", "工事予定", "会議予定", "行事予定", "電話履歴", "車検管理", "社用車予約", "日報", "運転免許管理", "備品修理管理", "個人ToDo"];
  scheduleTargets.forEach(sourceName => {
    const count = countScheduleRowsForMarker_(sourceName, marker);
    rows.push(autoTestRow_("一覧反映", sourceName, count > 0 ? "OK" : "要確認", "一覧スケジュール内 AUTO_TEST_ 行: " + count));
  });

  const alertCount = countRowsContainingMarkerByName_("要確認一覧", marker);
  rows.push(autoTestRow_("要確認", "要確認一覧", alertCount > 0 ? "OK" : "情報", "AUTO_TEST_ 行: " + alertCount));

  rows.push(autoTestRow_("車検", "更新済→車検履歴", vehicleResult.added > 0 ? "OK" : "NG", "履歴追加: " + vehicleResult.added + " / 状態空欄戻し: " + vehicleResult.cleared + (vehicleResult.reason ? " / " + vehicleResult.reason : "")));

  const historyCount = countRowsContainingMarkerByName_("車検履歴", marker);
  rows.push(autoTestRow_("車検", "車検履歴のAUTO_TEST_行", historyCount > 0 ? "OK" : "NG", "履歴行: " + historyCount));

  const vendorCount = countRowsContainingMarkerByName_("設定", marker);
  rows.push(autoTestRow_("修理業者", "手入力→設定シート反映", vendorCount > 0 ? "OK" : "NG", "設定シート AUTO_TEST_ 修理業者: " + vendorCount + " / 追加処理: " + addedVendors));

  const diagnostic = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GPT診断用");
  rows.push(autoTestRow_("診断", "GPT診断用シート", diagnostic ? "OK" : "情報", diagnostic ? "作成済み" : "③GPT診断で作成"));

  return rows;
}


function autoTestRow_(category, target, result, detail) {
  return {
    "区分": category,
    "対象": target,
    "結果": result,
    "詳細": detail,
    "作成日時": new Date()
  };
}


function createCompanyAutoTestResultSheet_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = ["区分", "対象", "結果", "詳細", "作成日時"];
  let sheet = ss.getSheetByName("自動テスト結果");
  if (!sheet) sheet = ss.insertSheet("自動テスト結果");

  resetSheetLight_(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (rows && rows.length) {
    const values = rows.map(row => headers.map(header => row[header] !== undefined ? row[header] : ""));
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }

  formatSheetBase_(sheet, headers.length);
  try {
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 180);
    sheet.setColumnWidth(3, 90);
    sheet.setColumnWidth(4, 420);
    sheet.setColumnWidth(5, 160);
    createFilterSafely_(sheet, headers.length);
  } catch (e) {}
}


function countCompanyAutoTestNg_(rows) {
  return (rows || []).filter(row => String(row["結果"] || "") === "NG").length;
}


function checkCompanyAutoTestIds_(sheetName, marker) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const idHeader = SHEET_ID_HEADERS[sheetName];
  if (!sheet || !idHeader || sheet.getLastRow() < 2) return {checked: 0, missing: 0};

  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf(idHeader);
  if (idCol < 0) return {checked: 0, missing: 0};

  const rows = findRowsContainingMarker_(sheet, marker);
  let missing = 0;
  rows.forEach(item => {
    if (!String(item.values[idCol] || "").trim()) missing++;
  });
  return {checked: rows.length, missing: missing};
}


function checkCompanyAutoTestCheckboxes_(sheetName, marker) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return {checked: 0, missing: 0};

  const headers = getHeaders_(sheet);
  const readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol <= 0) return {checked: 0, missing: 0};

  const rows = findRowsContainingMarker_(sheet, marker);
  let missing = 0;
  rows.forEach(item => {
    const rule = sheet.getRange(item.rowNumber, readCol).getDataValidation();
    if (!rule) missing++;
  });
  return {checked: rows.length, missing: missing};
}


function countScheduleRowsForMarker_(sourceName, marker) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("一覧スケジュール");
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const headers = getHeaders_(sheet);
  const sourceCol = headers.indexOf("元シート");
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.filter(row => {
    const sourceOk = sourceCol < 0 || String(row[sourceCol] || "") === sourceName;
    return sourceOk && rowValuesContainMarker_(row, marker);
  }).length;
}


function countRowsContainingMarkerByName_(sheetName, marker) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return 0;
  return findRowsContainingMarker_(sheet, marker).length;
}


function findRowsContainingMarker_(sheet, marker) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const result = [];
  values.forEach((row, index) => {
    if (rowValuesContainMarker_(row, marker)) {
      result.push({rowNumber: index + 2, values: row});
    }
  });
  return result;
}


function rowValuesContainMarker_(row, marker) {
  const prefix = AUTO_TEST_MARKER_PREFIX;
  return (row || []).some(value => {
    if (value === null || value === undefined) return false;
    const text = String(value);
    if (marker && text.indexOf(marker) >= 0) return true;
    return text.indexOf(prefix) >= 0;
  });
}


function deleteCompanyAutoTestData_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = uniqueList_(getInputSheetNames_().concat([
    "車検履歴",
    "一覧スケジュール",
    "要確認一覧",
    "担当別未読",
    "既読率集計",
    "ダッシュボード",
    "設定",
    "自動テスト結果"
  ]));

  const summary = {
    deletedRows: 0,
    deletedBySheet: {},
    remainingAutoTestRows: 0,
    checkedSheets: sheetNames.length
  };

  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

    let deletedInSheet = 0;
    const headers = getHeaders_(sheet);
    for (let row = sheet.getLastRow(); row >= 2; row--) {
      const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (isCompanyAutoTestRowForDeletion_(name, headers, values)) {
        sheet.deleteRow(row);
        deletedInSheet++;
        summary.deletedRows++;
      }
    }
    if (deletedInSheet > 0) summary.deletedBySheet[name] = deletedInSheet;
  });

  // 削除後にまだ AUTO_TEST_ が残っていないか軽く確認する。
  summary.remainingAutoTestRows = countCompanyAutoTestResidueRows_();

  return summary;
}


function isCompanyAutoTestRowForDeletion_(sheetName, headers, values) {
  if (rowValuesContainMarker_(values, "")) return true;

  const rowObj = {};
  (headers || []).forEach((header, index) => {
    rowObj[String(header || "").trim()] = values[index];
  });

  const textValues = (values || []).map(value => String(value || "").trim());
  const joined = textValues.join(" ");

  // p8/p9以前でフィードバックだけ残ることがあったため、過去版の自動テスト行も削除対象にする。
  // 実データの誤削除を避けるため、フィードバックは複数条件が揃う行だけ消す。
  if (sheetName === "フィードバック") {
    const targetSheet = String(rowObj["対象シート"] || "").trim();
    const category = String(rowObj["区分"] || "").trim();
    const policy = String(rowObj["対応方針"] || "").trim();
    const memo = String(rowObj["対応メモ"] || "").trim();
    const content = String(rowObj["気になった内容"] || "").trim();

    if (targetSheet === "自動テスト") return true;
    if (policy === "自動テスト" && (memo.indexOf("AUTO_TEST_") >= 0 || content.indexOf("AUTO_TEST_") >= 0)) return true;
    if (policy === "自動テスト" && category === "動作確認" && content.indexOf("フィードバック") >= 0) return true;
    if (policy === "自動テスト" && category === "動作不備" && content.indexOf("フィードバック") >= 0) return true;
  }

  // 自動テスト結果シートは、古い判定行を毎回掃除する。
  if (sheetName === "自動テスト結果") {
    const category = String(rowObj["区分"] || "").trim();
    const target = String(rowObj["対象"] || "").trim();
    if (["自動入力", "一覧反映", "既読チェック", "ID", "車検", "修理業者", "要確認", "診断", "次の操作", "事前整理", "情報", "削除"].indexOf(category) >= 0) return true;
    if (joined.indexOf("自動テスト") >= 0 || joined.indexOf("AUTO_TEST_") >= 0) return true;
    if (target.indexOf("自動テスト") >= 0 || target.indexOf("AUTO_TEST_") >= 0) return true;
  }

  // 設定シートは AUTO_TEST_ 修理業者などだけを消す。普通の「自動テスト」という語は消さない。
  if (sheetName === "設定") {
    return joined.indexOf("AUTO_TEST_") >= 0;
  }

  return false;
}


function countCompanyAutoTestResidueRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = uniqueList_(getInputSheetNames_().concat([
    "車検履歴",
    "一覧スケジュール",
    "要確認一覧",
    "担当別未読",
    "既読率集計",
    "ダッシュボード",
    "設定",
    "自動テスト結果"
  ]));

  let count = 0;
  sheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;
    const headers = getHeaders_(sheet);
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    values.forEach(row => {
      if (isCompanyAutoTestRowForDeletion_(name, headers, row)) count++;
    });
  });
  return count;
}


function setupCompanyQualificationPack() {
  // p38:
  // 「資格管理だけ整える（軽量）」でも起動時間超過するため、ここは超軽量にする。
  // acquireCompanyPackLock_() は最大3分待つので、前回処理やonChangeと重なると待機だけで時間を使う。
  // このメニューでは短時間ロックにし、重い処理は分割メニューへ逃がす。
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    toast_("資格管理パックは他の処理中です。30秒ほど待ってから再実行してください。");
    return;
  }

  try {
    const result = repairQualificationManagementLight_();
    toast_(
      "資格管理パック（超軽量）が完了しました。"
      + " ヘッダー:" + result.headerFixed
      + " / マスタ確認:" + result.masterChecked
      + " / 通常プルダウン:" + result.fixedRows + "行"
      + " / 資格名:" + result.nameRows + "行"
      + " / ID非表示:" + result.hiddenIdCols + "列"
      + "。ソート・FALSE掃除・チェックボックス復旧・社員別一覧は必要時だけ別メニューで実行してください。"
    );
  } catch (err) {
    toast_("資格管理パック（超軽量）でエラー: " + err.message);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function repairQualificationManagementLight_() {
  // p38:
  // 起動時間超過対策として、日常の資格管理パックでは「壊れやすい最小部分」だけ直す。
  // ここでは実行しないもの:
  // - 資格管理の全行ソート
  // - FALSE掃除
  // - 既読/個人確認チェックボックス全再設定
  // - 社員別/資格別一覧作成
  // - 所有者グループ作成
  // - 表示書式の全調整
  // これらは必要時だけ保守メニューから個別実行する。
  clearSettingsCache_();
  clearQualificationValidationListsCache_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("資格管理");
  if (!sheet) sheet = ss.insertSheet("資格管理");

  const headerResult = ensureQualificationManagementHeadersUltraLight_(sheet);
  const masterResult = ensureQualificationMasterMinimumForUltraLight_();
  const dropdownResult = refreshQualificationManagementDropdownsUltraLight_(sheet);

  let hiddenIdCols = 0;
  try { hiddenIdCols = hideSystemColumnsForSheet_(sheet); } catch (e) {}
  try { hideQualificationMasterIdColumnUltraLight_(); } catch (e) {}

  return {
    headerFixed: headerResult.fixed || 0,
    masterChecked: masterResult.checked || 0,
    masterAdded: masterResult.added || 0,
    masterUpdated: masterResult.updated || 0,
    fixedRows: dropdownResult.fixedRows || 0,
    fixedCols: dropdownResult.fixedCols || 0,
    nameRows: dropdownResult.nameRows || 0,
    cleared: 0,
    sortedRows: 0,
    falseCleaned: 0,
    checkboxCount: 0,
    hiddenIdCols: hiddenIdCols
  };
}

function ensureQualificationManagementHeadersUltraLight_(sheet) {
  // p38:
  // 旧 repairQualificationManagementHeadersSafe_() はシート全体を書き直すため重い。
  // 普段は1行目の不足見出しだけ補う。列ズレが深刻な場合だけ、従来の重い修復メニューを使う。
  if (!sheet) return {fixed: 0};
  const desiredHeaders = getSheetHeaders_()["資格管理"] || [];
  if (!desiredHeaders.length) return {fixed: 0};

  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    setupSheet_(sheet, desiredHeaders);
    return {fixed: 1};
  }

  ensureColumns_(sheet, desiredHeaders.length);

  const current = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), desiredHeaders.length))
    .getValues()[0]
    .map(v => String(v || "").trim());

  let fixed = 0;
  const row = current.slice(0, desiredHeaders.length);
  desiredHeaders.forEach((header, i) => {
    const now = String(row[i] || "").trim();
    if (!now) {
      row[i] = header;
      fixed++;
      return;
    }
    if (now === "保有区分" && header === "取得区分") {
      row[i] = "取得区分";
      fixed++;
      return;
    }
    if (now === "有効期限" && header === "更新期限") {
      row[i] = "更新期限";
      fixed++;
      return;
    }
    if (now === "証明書コピー" && header === "コピー有無") {
      row[i] = "コピー有無";
      fixed++;
      return;
    }
  });

  // F列「取得日」抜けの保険。ここだけは安全に直接戻す。
  if (desiredHeaders[5] === "取得日" && String(row[5] || "").trim() !== "取得日") {
    if (!String(row[5] || "").trim()) {
      row[5] = "取得日";
      fixed++;
    }
  }

  if (fixed > 0) {
    sheet.getRange(1, 1, 1, desiredHeaders.length).setValues([row]);
  }

  return {fixed: fixed};
}

function ensureQualificationMasterMinimumForUltraLight_() {
  // p38:
  // 軽量パックで資格マスタ全件を毎回setValues/sortしない。
  // 資格マスタが空・必須列不足のときだけ、既存の安全補完を呼ぶ。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("資格マスタ");
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 3) {
    const result = ensureQualificationPaperMasterRowsForLightPack_();
    return {checked: 1, added: result.added || 0, updated: result.updated || 0};
  }

  const headers = getHeaders_(sheet);
  const required = ["表示順", "区分", "資格名", "一覧表示"];
  const missing = required.some(h => headers.indexOf(h) < 0);
  if (missing) {
    const result = ensureQualificationPaperMasterRowsForLightPack_();
    return {checked: 1, added: result.added || 0, updated: result.updated || 0};
  }

  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
  return {checked: 1, added: 0, updated: 0};
}

function getQualificationManagementDropdownTargetRowCountUltraLight_(sheet) {
  // p38:
  // 500行固定だと少人数運用では重い。入力済み最終行＋80行、最低200行までに絞る。
  if (!sheet) return 1;
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const targetLastRow = Math.min(maxRows, Math.max(lastRow + 80, 201));
  return Math.max(targetLastRow - 1, 1);
}

function refreshQualificationManagementDropdownsUltraLight_(sheet) {
  // p38:
  // 値の整合チェックや不一致クリアはしない。入力規則だけ高速に戻す。
  if (!sheet || sheet.getName() !== "資格管理") return {fixedRows: 0, fixedCols: 0, nameRows: 0};

  const rowCount = getQualificationManagementDropdownTargetRowCountUltraLight_(sheet);
  const fixed = setQualificationManagementFixedDropdownsForRange_(sheet, 2, rowCount);
  const names = applyQualificationManagementNameValidationsNoClean_(sheet, rowCount);

  return {
    fixedRows: fixed.rows || 0,
    fixedCols: fixed.cols || 0,
    nameRows: names.rows || 0
  };
}

function applyQualificationManagementNameValidationsNoClean_(sheet, rowCount) {
  // p38:
  // C列「資格名」の入力規則をB列「区分」ごとに貼るだけ。
  // 既存値のクリアやsetValuesを行わないので、途中停止しても値が壊れにくく軽い。
  if (!sheet || sheet.getName() !== "資格管理") return {rows: 0};

  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  if (categoryCol <= 0 || nameCol <= 0) return {rows: 0};

  const count = Math.max(1, Math.min(Number(rowCount || 1), sheet.getMaxRows() - 1));
  const categories = sheet.getRange(2, categoryCol, count, 1)
    .getValues()
    .map(r => normalizeQualificationCategoryValue_(r[0]));

  const ruleCache = {};
  let runStart = 0;
  let currentHeader = null;

  const applyRun_ = (startIndex, endIndex, helperHeader) => {
    if (startIndex > endIndex) return;
    if (!ruleCache[helperHeader]) {
      const categoryForRule = helperHeader === "全資格" ? "" : helperHeader;
      ruleCache[helperHeader] = buildQualificationNameValidationRuleForCategory_(categoryForRule, false);
    }
    sheet.getRange(2 + startIndex, nameCol, endIndex - startIndex + 1, 1)
      .setDataValidation(ruleCache[helperHeader]);
  };

  categories.forEach((category, i) => {
    const helperHeader = getQualificationHelperHeaderForCategory_(category);
    if (i === 0) {
      currentHeader = helperHeader;
      runStart = 0;
      return;
    }
    if (helperHeader !== currentHeader) {
      applyRun_(runStart, i - 1, currentHeader || "全資格");
      currentHeader = helperHeader;
      runStart = i;
    }
  });

  if (categories.length > 0) {
    applyRun_(runStart, categories.length - 1, currentHeader || "全資格");
  }

  return {rows: count};
}

function hideQualificationMasterIdColumnUltraLight_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格マスタ");
  if (!sheet) return 0;
  return hideSystemColumnsForSheet_(sheet);
}


function onChange(e) {
  // p38:
  // OTHER まで拾うと、入力規則貼り直し自体でも再実行されて重くなることがある。
  // 行/列の構造変更だけを対象にする。
  let lock = null;
  try {
    const changeType = e && e.changeType ? String(e.changeType) : "";
    const targetTypes = ["INSERT_ROW", "REMOVE_ROW", "INSERT_COLUMN", "REMOVE_COLUMN"];
    if (!targetTypes.includes(changeType)) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    if (!sheet || sheet.getName() !== "資格管理") return;

    lock = LockService.getDocumentLock();
    if (!lock.tryLock(3000)) return;

    repairQualificationManagementDropdownsAfterStructureChange_(sheet, changeType || "CHANGE");
  } catch (err) {
    console.log("onChange qualification dropdown repair error: " + err.message);
  } finally {
    try { if (lock) lock.releaseLock(); } catch (e) {}
  }
}

function ensureQualificationManagementOnChangeTriggerQuiet_() {
  // メニュー実行時だけ呼ぶ。行挿入後の自動復旧に必要なインストール型onChangeを1個だけ維持する。
  // simple onOpen/onEditからは呼ばないこと。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  let existing = 0;
  const duplicates = [];

  triggers.forEach(trigger => {
    try {
      if (trigger.getHandlerFunction && trigger.getHandlerFunction() === "onChange") {
        existing++;
        if (existing > 1) duplicates.push(trigger);
      }
    } catch (e) {}
  });

  duplicates.forEach(trigger => {
    try { ScriptApp.deleteTrigger(trigger); } catch (e) {}
  });

  if (existing === 0) {
    ScriptApp.newTrigger("onChange")
      .forSpreadsheet(ss)
      .onChange()
      .create();
    return {created: 1, existing: 0, deletedDuplicates: duplicates.length};
  }

  return {created: 0, existing: existing, deletedDuplicates: duplicates.length};
}


function installQualificationManagementRowInsertTrigger() {
  const lock = acquireCompanyPackLock_("資格管理 行挿入検知トリガー設定");
  if (!lock) return;

  try {
    const result = ensureQualificationManagementOnChangeTriggerQuiet_();
    const msg = result.created
      ? "資格管理の行挿入検知トリガーを設定しました。"
      : "資格管理の行挿入検知トリガーは設定済みです。";
    toast_(msg + (result.deletedDuplicates ? " 重複トリガーを" + result.deletedDuplicates + "個整理しました。" : ""));
  } catch (err) {
    toast_("行挿入検知トリガー設定でエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}

function repairQualificationManagementDropdownsAfterRowInsert() {
  const lock = acquireCompanyPackLock_("資格管理 行挿入後プルダウン復旧");
  if (!lock) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
    if (!sheet) {
      toast_("資格管理シートが見つかりません。");
      return;
    }
    const result = repairQualificationManagementDropdownsAfterStructureChange_(sheet, "MANUAL");
    toast_(
      "資格管理の行挿入後プルダウンを復旧しました"
      + "（所有者など:" + (result.fixedRows || 0)
      + "行 / 資格名:" + (result.nameRows || 0)
      + "行 / チェック:" + (result.checkboxCount || 0) + "個）。"
    );
  } catch (err) {
    toast_("行挿入後プルダウン復旧でエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}

function repairQualificationManagementDropdownsAfterStructureChange_(sheet, reason) {
  // p38:
  // 行挿入後の復旧も超軽量にする。トリガー確認・ID採番・表示全調整・チェックボックス全復旧はしない。
  // 消えやすいプルダウンだけを戻す。
  if (!sheet || sheet.getName() !== "資格管理") return {fixedRows: 0, nameRows: 0, checkboxCount: 0};

  clearSettingsCache_();
  clearQualificationValidationListsCache_();

  let headerResult = {fixed: 0};
  try { headerResult = ensureQualificationManagementHeadersUltraLight_(sheet); } catch (e) {}

  const dropdowns = refreshQualificationManagementDropdownsUltraLight_(sheet);
  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
  try { SpreadsheetApp.flush(); } catch (e) {}

  console.log("qualification dropdown ultra light repair after structure change: " + (reason || "") + " / nameRows=" + (dropdowns.nameRows || 0));
  return {
    fixedRows: dropdowns.fixedRows || 0,
    fixedCols: dropdowns.fixedCols || 0,
    nameRows: dropdowns.nameRows || 0,
    cleared: 0,
    checkboxCount: 0,
    headerFixed: headerResult.fixed || 0
  };
}

function setupCompanyQualificationListPack() {
  const lock = acquireCompanyPackLock_("資格一覧だけ更新");
  if (!lock) return;

  try {
    createEmployeeQualificationListSheet_();
    createQualificationHolderListSheet_();
    createQualificationDuplicateCheckSheet_();
    toast_("社員別資格一覧・資格別保有者一覧・資格重複チェックを更新しました。");
  } catch (err) {
    toast_("資格一覧更新でエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}

function setupCompanyOperationPack() {
  try {
    repairVehicleHistoryAndVehicleNameDropdowns_();
    repairConstructionContractAmountAndTax_();
    rebuildRepairVendorMasterFromEquipmentSheet_();
    repairRepairVendorDropdown_();
    repairEventScheduleColumns_();
    refreshInputSheets();
    toast_("車検・工事・備品・行事パックが完了しました。各シートを確認してください。");
  } catch (err) {
    toast_("車検・工事・備品・行事パックでエラー: " + err.message);
    throw err;
  }
}

function setupCompanyHiddenSheetsPack() {
  try {
    hideSupportSheets();
    toast_("裏方シートを整理しました。設定・資格マスタ・ダッシュボードは表示のままです。");
  } catch (err) {
    toast_("裏方シート整理でエラー: " + err.message);
    throw err;
  }
}

function repairEventScheduleColumns_() {
  ensureSheetHeadersPreserveData_("行事予定");
  hideEventContentColumnIfExists_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("行事予定");
  if (sheet) {
    applyDataValidationByHeaders_(sheet, getHeaders_(sheet));
    applyColorRules(sheet);
    formatSheetByName_("行事予定");
  }
}


function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  if (row <= 1) return;
  const sheetName = sheet.getName();

  // 設定シートを編集した場合、担当者・既読確認者・社用車プルダウンのキャッシュだけ消す。
  // 入力シートへの反映は「入力シートだけ設定反映」を実行。
  if (sheetName === "設定") {
    clearSettingsCache_();
    return;
  }

  // 資格かんたん登録は入力補助シート。
  // 通常のID/通知/既読処理には流さず、「追加する資格」だけ軽量処理する。
  if (sheetName === "資格かんたん登録") {
    try {
      handleQualificationEasyPasteAppend_(e);
    } catch (err) {
      console.log("qualification easy paste onEdit error: " + err.message);
    }
    return;
  }

  // 資格管理は正式データ。
  // p31:
  // 区分/資格名以外（取得区分・保有状況・取得日など）を編集した時も、
  // C列「資格名」の入力規則を周辺行へ貼り直して、プルダウンが消えたままにならないようにする。
  // ここで return し、汎用onEditへ流さないことで、資格管理専用の軽い補修だけに限定する。
  if (sheetName === "資格管理") {
    try {
      ensureRowsAfterEdit_(sheet, row);
      ensureIdForEditedRow_(sheet, row);

      const handled = handleQualificationManagementEdit_(e);
      if (!handled) {
        normalizeClearLabelForEditedCell_(e);
        updateNoticeForEditedRow_(sheet, row);
      }

      // 入力中は自動ソートしない。C列の候補だけ、編集行＋下の予備行へ維持する。
      // p43: 更新期限・コピー有無・状態だけ、資格更新履歴へ軽量記録する。
      handleQualificationUpdateHistoryEdit_(e);
      repairQualificationManagementDropdownsAroundEditedRow_(sheet, row);
      setCheckboxesForEditedRow_(sheet, row);
      try { hideSystemColumnsForSheet_(sheet); } catch (hideErr) {}
      markSummaryDirty_();
    } catch (err) {
      console.log("qualification management onEdit error: " + err.message);
    }
    return;
  }

  if (isAutoOutputSheet_(sheetName) || SUPPORT_SHEETS.includes(sheetName) || isUpdateHistorySheet_(sheetName)) return;

  try {
    ensureRowsAfterEdit_(sheet, row);
    ensureIdForEditedRow_(sheet, row);

    // 電話履歴の「電話対応」は専用処理で先に完結させる。
    // ここで return しないと、共通の通知判定で空欄行が「期限切れ」に戻ることがある。
    // 車検管理の「車検状態=更新済」は専用処理で先に完結させる。
    // 共通処理に流すと、環境によって更新済が残ることがあるため、ここで履歴追加・期限更新・空欄戻しまで行う。
    if (handleConstructionContactEdit_(e)) {
      setCheckboxesForEditedRow_(sheet, row);
      markSummaryDirty_();
      return;
    }

    if (handleRepairVendorEdit_(e)) {
      setCheckboxesForEditedRow_(sheet, row);
      markSummaryDirty_();
      return;
    }

    if (handleVehicleInspectionStatusEdit_(e)) {
      setCheckboxesForEditedRow_(sheet, row);
      markSummaryDirty_();
      return;
    }

    if (handlePhoneHistoryEdit_(e)) {
      setCheckboxesForEditedRow_(sheet, row);
      markSummaryDirty_();
      return;
    }

    // p43: 運転免許管理の有効期限・コピー有無・状態だけ、更新履歴へ軽量記録する。
    handleDriverLicenseUpdateHistoryEdit_(e);

    normalizeClearLabelForEditedCell_(e);
    processVehicleInspectionUpdatedRowsOnEdit_(e);
    updateNoticeForEditedRow_(sheet, row);
    setCheckboxesForEditedRow_(sheet, row);
    markSummaryDirty_();
  } catch (err) {
    console.log("onEdit error: " + err.message);
  }
}


function autoRefreshPhoneHistoryOnComplete_(e) {
  handlePhoneHistoryEdit_(e);
}

function handlePhoneHistoryEdit_(e) {
  if (!e || !e.range) return false;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "電話履歴") return false;
  if (e.range.getRow() <= 1) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  if (editedHeader !== "電話対応") return false;

  const row = e.range.getRow();

  // e.value は「空欄に戻す」のまま残ることがあるため、先に明示処理する。
  // これにより、完了日時・通知が古い値で残る問題を防ぐ。
  let newValue = String(e.value || e.range.getValue() || "").trim();
  if (newValue === CLEAR_LABEL) {
    newValue = "";
    e.range.clearContent();
  } else {
    newValue = String(e.range.getValue() || newValue || "").trim();
    if (newValue === CLEAR_LABEL) {
      newValue = "";
      e.range.clearContent();
    }
  }

  applyPhoneHistoryStatusForRow_(sheet, row, newValue);
  normalizePhoneHistoryCompletionColumns_(sheet);

  // 電話履歴は、完了以外を上、完了を下にする。
  // 未対応/対応中/折返し/空欄の細かい状態順ではなく、各グループ内は日時の新しい順にする。
  sortPhoneHistorySheet_();

  createScheduleList();
  createAlertList();
  clearSummaryDirty_();

  if (newValue === "完了") {
    toast_("電話履歴を完了として反映し、完了行を下へ整理しました");
  } else if (newValue === "") {
    toast_("電話対応を空欄に戻し、未完了側へ戻しました");
  } else {
    toast_("電話履歴を更新し、未完了側を日時順に整理しました");
  }

  return true;
}

function applyPhoneHistoryStatusForRow_(sheet, row, phoneStatus) {
  const headers = getHeaders_(sheet);
  const completedCol = headers.indexOf("対応完了日時") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const dateCol = headers.indexOf("日付") + 1;
  const statusCol = headers.indexOf("電話対応") + 1;
  if (row <= 1) return;

  const status = String(phoneStatus || "").trim();

  if (completedCol > 0) {
    const completedCell = sheet.getRange(row, completedCol);
    completedCell.setNumberFormat("yyyy/mm/dd hh:mm");

    if (status === "完了") {
      const current = completedCell.getValue();
      if (!current || isDateOnlyLike_(current)) completedCell.setValue(new Date());
      completedCell.setNumberFormat("yyyy/mm/dd hh:mm");
    } else {
      completedCell.clearContent();
      completedCell.setNumberFormat("yyyy/mm/dd hh:mm");
    }
  }

  if (noticeCol > 0) {
    const dateValue = dateCol > 0 ? sheet.getRange(row, dateCol).getValue() : "";
    sheet.getRange(row, noticeCol).setValue(getPhoneNoticeText_(dateValue, status));
  }

  if (statusCol > 0 && status === "") {
    sheet.getRange(row, statusCol).clearContent();
  }
}

function updatePhoneCompletedAtForRow_(sheet, row, phoneStatus) {
  applyPhoneHistoryStatusForRow_(sheet, row, phoneStatus);
}

function normalizePhoneHistoryCompletionColumns_(sheet) {
  if (!sheet || sheet.getName() !== "電話履歴") return;

  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf("電話対応") + 1;
  const completedCol = headers.indexOf("対応完了日時") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const dateCol = headers.indexOf("日付") + 1;
  if (statusCol <= 0 || sheet.getLastRow() < 2) {
    formatPhoneCompletedAtColumn_(sheet);
    return;
  }

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  const completedValues = [];
  const noticeValues = [];

  values.forEach(rowValues => {
    const status = String(rowValues[statusCol - 1] || "").trim();
    const dateValue = dateCol > 0 ? rowValues[dateCol - 1] : "";
    const currentCompleted = completedCol > 0 ? rowValues[completedCol - 1] : "";

    if (status === "完了") {
      completedValues.push([currentCompleted || new Date()]);
    } else {
      completedValues.push([""]);
    }

    noticeValues.push([getPhoneNoticeText_(dateValue, status)]);
  });

  if (completedCol > 0) {
    sheet.getRange(2, completedCol, rowCount, 1)
      .setValues(completedValues)
      .setNumberFormat("yyyy/mm/dd hh:mm")
      .clearDataValidations();
  }

  if (noticeCol > 0) {
    sheet.getRange(2, noticeCol, rowCount, 1).setValues(noticeValues);
  }

  formatPhoneCompletedAtColumn_(sheet);
}

function getPhoneNoticeText_(dateValue, phoneStatus) {
  const status = String(phoneStatus || "").trim();

  // 空欄へ戻した電話は「期限切れ」に戻さない。
  // 未対応にしたい場合は、明示的に「未対応」を選ぶ運用にする。
  if (!status) return "";
  if (status === "完了") return "完了";
  return getNoticeText(dateValue, status);
}

function isDateOnlyLike_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return false;
  return value.getHours() === 0
    && value.getMinutes() === 0
    && value.getSeconds() === 0
    && value.getMilliseconds() === 0;
}

function formatPhoneCompletedAtColumn_(sheet) {
  if (!sheet || sheet.getName() !== "電話履歴") return;

  const headers = getHeaders_(sheet);
  const completedCol = headers.indexOf("対応完了日時") + 1;
  if (completedCol <= 0) return;

  const rowCount = Math.max(getApplyRowCount_(sheet), 300);
  sheet.getRange(2, completedCol, rowCount, 1)
    .setNumberFormat("yyyy/mm/dd hh:mm")
    .clearDataValidations();

  try {
    sheet.setColumnWidth(completedCol, 145);
  } catch (e) {}
}

function sortPhoneHistorySheet() {
  sortPhoneHistorySheet_();
  toast_("電話履歴を未完了上・完了下・日時の新しい順に整理しました");
}

function sortPhoneHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("電話履歴");
  if (!sheet || sheet.getLastRow() < 3) return;

  normalizePhoneHistoryCompletionColumns_(sheet);

  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = headers.length;
  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, 1, rowCount, lastCol).getValues();

  const dataRows = values.filter(row => isBusinessDataRow_(headers, row, "電話履歴"));
  const emptyRows = values.filter(row => !isBusinessDataRow_(headers, row, "電話履歴"));

  dataRows.sort((a, b) => {
    const ao = getPhoneSortOrder_(a, headers);
    const bo = getPhoneSortOrder_(b, headers);
    if (ao !== bo) return ao - bo;

    const at = getPhoneDateTimeValue_(a, headers);
    const bt = getPhoneDateTimeValue_(b, headers);
    if (at !== bt) return bt - at; // 新しい順

    const ai = String(a[headers.indexOf("相手")] || "");
    const bi = String(b[headers.indexOf("相手")] || "");
    return ai.localeCompare(bi, "ja");
  });

  const sorted = dataRows.concat(emptyRows);
  sheet.getRange(2, 1, sorted.length, lastCol).setValues(sorted);
  setCheckboxesForDataRows(sheet);
  formatPhoneCompletedAtColumn_(sheet);
}

function getPhoneSortOrder_(row, headers) {
  const statusCol = headers.indexOf("電話対応");
  const status = statusCol >= 0 ? String(row[statusCol] || "").trim() : "";

  // v9.7.20:
  // 完了だけ下へ。それ以外（未対応/対応中/折返し/空欄/その他）はすべて上。
  // 同じグループ内は getPhoneDateTimeValue_() で日付・時間の新しい順にする。
  return status === "完了" ? 2 : 1;
}

function getPhoneDateTimeValue_(row, headers) {
  const dateCol = headers.indexOf("日付");
  const timeCol = headers.indexOf("時間");

  const d = dateCol >= 0 ? toDateOnly_(row[dateCol]) : null;
  if (!d) return 0;

  let base = d.getTime();
  const t = timeCol >= 0 ? row[timeCol] : null;

  if (t instanceof Date && !isNaN(t.getTime())) {
    base += t.getHours() * 60 * 60 * 1000;
    base += t.getMinutes() * 60 * 1000;
    base += t.getSeconds() * 1000;
  }

  return base;
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
    hideSystemColumnsForAllInputSheets_({silent: true});
    clearSummaryDirty_();
    toast_("日常更新が完了しました（一覧＋要確認・ID列非表示）");
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


function repairConstructionScheduleDropdownAndColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工事予定");
  if (!sheet) {
    toast_("工事予定シートが見つかりません");
    return;
  }

  const headers = getHeaders_(sheet);
  if (!headers.length) {
    toast_("工事予定シートのヘッダーが見つかりません");
    return;
  }

  const statusCol = headers.indexOf("状態") + 1;
  if (statusCol > 0) {
    setDropdown(sheet, statusCol, getStatusListForSheet_("工事予定"));
  }

  const phoneCol = headers.indexOf("電話対応") + 1;
  if (phoneCol > 0) {
    setDropdown(sheet, phoneCol, [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"]);
  }

  const staffCol = headers.indexOf("担当") + 1;
  if (staffCol > 0) {
    setDropdown(sheet, staffCol, [CLEAR_LABEL, ...getStaffMembers_()]);
  }

  updateNoticeColumnForSheet_(sheet);
  applyColorRules(sheet);
  createFilterSafely_(sheet, headers.length);

  repairConstructionContractAmountAndTax_();
  toast_("工事予定の未契約/請求済みプルダウン・契約金額円表示・税プルダウン・連絡先文字列固定・通知完了・カラーを修復しました");
}

function refreshInputSheets() {
  // p41:
  // 旧「入力シートだけ設定反映」は、全入力シートに対して
  // - 個人確認列の再構成
  // - 入力規則の広範囲再設定
  // - 通知再計算
  // - ID採番
  // - チェックボックス全復旧
  // - 条件付き書式再作成
  // を一度に実行していたため、実行時間上限に当たりやすかった。
  // 通常メニューでは「入力しやすさを戻す」ための最小処理だけに絞る。
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(8000)) {
    toast_("入力シート設定反映は他の処理中です。30秒ほど待ってから再実行してください。");
    return;
  }

  try {
    ensureSettingsSheet_();
    ensureDefaultSettingRows_("社用車", DEFAULT_COMPANY_VEHICLES, "社用車プルダウン用");
    ensureDefaultSettingRows_("車両名", DEFAULT_COMPANY_VEHICLES, "車検管理用");
    clearSettingsCache_();
    clearQualificationValidationListsCache_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetCount = 0;
    let validationCols = 0;
    let hiddenCols = 0;

    getInputSheetNames_().forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastColumn() < 1) return;

      const headers = getHeaders_(sheet);
      if (!headers.length) return;

      const result = applyInputSheetSettingsUltraLight_(sheet, headers);
      validationCols += result.validationCols || 0;
      hiddenCols += result.hiddenCols || 0;
      sheetCount++;
    });

    toast_(
      "入力シートだけ設定反映（超軽量）が完了しました。"
      + " 対象:" + sheetCount + "シート"
      + " / 入力規則:" + validationCols + "列"
      + " / ID非表示:" + hiddenCols + "列。"
      + " 通知・ID採番・チェックボックス・色は必要時だけ完全版/個別メニューで実行してください。"
    );
  } catch (err) {
    toast_("入力シート設定反映（超軽量）でエラー: " + err.message);
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function refreshInputSheetsFullMaintenance() {
  // p41:
  // 旧版と同じ完全メンテナンス。重いため、普段は使わない。
  // 個人確認列再構成、通知再計算、ID採番、チェックボックス全復旧、条件付き書式まで必要な時だけ実行する。

  ensureSettingsSheet_();
  ensureDefaultSettingRows_("社用車", DEFAULT_COMPANY_VEHICLES, "社用車プルダウン用");
  ensureDefaultSettingRows_("車両名", DEFAULT_COMPANY_VEHICLES, "車検管理用");
  clearSettingsCache_();

  // 設定シートの「既読確認者」を正として、既存シートに残った旧個人確認列を整理する。
  // これを入れないと、山田/高橋/鈴木などの旧列がヘッダーに残り続ける。
  rebuildPersonalCheckColumnsFromSettings_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureV9714MinorLayoutUpdates_();

  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;

    const headers = getHeaders_(sheet);
    if (!headers.length) return;

    // 入力規則を再反映する。
    // 状態プルダウンの追加（例：工事予定の未契約）や社用車設定変更はここで反映。
    applyDataValidationByHeaders_(sheet, headers);

    if (sheet.getLastRow() >= 2) {
      if (name === "車検管理") normalizeVehicleInspectionUpdatedRows_(sheet);
      updateNoticeColumnForSheet_(sheet);
      ensureIdsForSheet_(sheet);
      setCheckboxesForDataRows(sheet);
      if (name === "電話履歴") normalizePhoneHistoryCompletionColumns_(sheet);
    }

    // 条件付き書式も再反映する。
    // 通知「完了」や状態「未契約」「完了」の色はここで反映。
    applyColorRules(sheet);

    // p20: 設定反映や表示調整後にID列・カレンダーID列が再表示されないよう、標準で隠す。
    hideSystemColumnsForSheet_(sheet);
  });

  toast_("入力シート設定反映（完全・重い）を実行しました（プルダウン・通知・色・ID列非表示を更新）");
}


function getInputSheetSettingsTargetRowCountUltraLight_(sheet) {
  // p41:
  // 全シートに300～500行分を毎回貼ると重いので、通常メニューでは
  // 入力済み最終行＋50行、最低100行までに抑える。
  if (!sheet) return 1;
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const targetLastRow = Math.min(maxRows, Math.max(lastRow + 50, 101));
  return Math.max(targetLastRow - 1, 1);
}

function applyInputSheetSettingsUltraLight_(sheet, headers) {
  // p41:
  // 通常の入力設定反映では、入力規則・最低限の表示形式・ID列非表示だけ行う。
  // 通知、ID採番、チェックボックス全復旧、条件付き書式は重いためここでは行わない。
  const sheetName = sheet.getName();
  if (isAutoOutputSheet_(sheetName)) return {validationCols: 0, hiddenCols: 0};

  if (sheetName === "資格管理") {
    const headerResult = ensureQualificationManagementHeadersUltraLight_(sheet);
    const dropdownResult = refreshQualificationManagementDropdownsUltraLight_(sheet);
    const hidden = hideSystemColumnsForSheet_(sheet) || 0;
    return {
      validationCols: (dropdownResult.fixedCols || 0) + 1 + (headerResult.fixed ? 1 : 0),
      hiddenCols: hidden
    };
  }

  const rowCount = getInputSheetSettingsTargetRowCountUltraLight_(sheet);
  let validationCols = 0;

  headers.forEach((header, index) => {
    const col = index + 1;

    if (header === "対応完了日時") {
      sheet.getRange(2, col, rowCount, 1).setNumberFormat("yyyy/mm/dd hh:mm");
      return;
    }

    if (DATE_HEADERS.includes(header)) {
      sheet.getRange(2, col, rowCount, 1).setNumberFormat("yyyy/mm/dd");
    }

    if (TIME_HEADERS.includes(header)) {
      sheet.getRange(2, col, rowCount, 1).setNumberFormat("hh:mm");
    }

    // p39/p41:
    // 車番・折り返し電話番号・工事予定の連絡先は先頭0を守るため文字列固定。
    if (VEHICLE_NUMBER_HEADERS.includes(header) || header === "連絡先") {
      sheet.getRange(2, col, rowCount, 1).setNumberFormat("@");
    }

    if (sheetName === "工事予定" && header === "契約金額") {
      sheet.getRange(2, col, rowCount, 1).setNumberFormat("¥#,##0");
      return;
    }

    if (NUMBER_HEADERS.includes(header)) {
      const format = String(header || "") === "契約金額" || String(header || "") === "金額" ? "¥#,##0" : "#,##0";
      sheet.getRange(2, col, rowCount, 1).setNumberFormat(format);
    }

    if (sheetName === "備品修理管理" && header === "修理業者") {
      setDropdownForRowsUltraLight_(sheet, col, [CLEAR_LABEL, ...getRepairVendors_()], rowCount, true);
      validationCols++;
      return;
    }

    const list = getDropdownListForHeader_(sheetName, header);
    if (list) {
      setDropdownForRowsUltraLight_(sheet, col, list, rowCount, false);
      validationCols++;
    }
  });

  const hiddenCols = hideSystemColumnsForSheet_(sheet) || 0;
  return {validationCols: validationCols, hiddenCols: hiddenCols};
}

function setDropdownForRowsUltraLight_(sheet, col, list, rowCount, allowInvalid) {
  if (!sheet || col <= 0 || !list || !list.length) return;
  const count = Math.max(1, Math.min(Number(rowCount || 1), sheet.getMaxRows() - 1));
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  sheet.getRange(2, col, count, 1).setDataValidation(rule);
}

function getSheetHeaders_() {
  const personal = getPersonalMembers_();
  return {
    "出先予定": ["日付", "行き先", "用件", "担当", "社用車", "状態", "通知", "電話対応", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "出先ID"],
    "工事予定": ["工事名", "現場", "依頼主", "連絡先", "契約日", "契約金額", "税", "開始日", "終了日", "状態", "担当", "通知", "電話対応", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "工事ID"],
    "会議予定": ["日付", "開始時刻", "終了時刻", "会議名", "内容", "担当", "状態", "通知", "資料", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "会議ID"],
    "行事予定": ["日付", "終了日", "開始時刻", "終了時刻", "行事名", "担当", "状態", "通知", "備考", READ_HEADER, ...personal, CALENDAR_ID_HEADER, "行事ID"],
    "作業状況": ["現場", "作業内容", "状態", "担当", "写真", "通知", READ_HEADER, ...personal, "備考", "作業ID"],
    "電話履歴": ["日付", "時間", "相手", "折り返し電話番号", "内容", "電話対応", "担当", "対応完了日時", "対応メモ", "通知", READ_HEADER, ...personal, "備考", "電話ID"],
    "車検管理": ["車両名", "車番", "車検期限", "車検予定日", "新車検期限", "通知", "保険期限", "車検状態", "担当", "写真", READ_HEADER, ...personal, "備考", "車両ID"],
    "車検履歴": ["更新日", "車両名", "車番", "旧車検期限", "新車検期限", "車検予定日", "担当", "備考", "履歴ID"],
    "社用車予約": ["日付", "開始時刻", "終了時刻", "社用車", "利用者", "行き先", "用途", "状態", "通知", READ_HEADER, ...personal, "備考", CALENDAR_ID_HEADER, "予約ID"],
    "日報": ["日付", "入力者", "担当", "現場", "作業内容", "進捗", "問題点", "明日の予定", "他現場状況", "写真", "日報文章", "状態", "備考", READ_HEADER, ...personal, "PDFリンク", "日報ID"],
    "日報レシート管理": ["日付", "担当", "現場", "支払先", "内容", "区分", "金額", "支払方法", "レシート写真", "経理確認", "精算状態", "通知", "備考", "レシートID"],
    "運転免許管理": ["所有者", ...LICENSE_TYPE_HEADERS, DRIVER_LICENSE_ISSUE_DATE_HEADER, DRIVER_LICENSE_EXPIRY_DATE_HEADER, "コピー有無", "状態", "通知", READ_HEADER, ...personal, "備考", "免許ID"],
    "運転免許明細": ["所有者", "免許種類", "取得日", "保有状況", "限定条件", "備考", "免許明細ID"],
    "運転免許更新履歴": ["記録日時", "所有者", "変更項目", "旧値", "新値", "旧有効期限", "新有効期限", "免許証交付日", "コピー有無", "状態", "備考", "元免許ID", "免許更新履歴ID"],
    "資格管理": ["所有者", "区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER, ...personal, "備考", "資格ID"],
    "資格更新履歴": ["記録日時", "所有者", "区分", "資格名", "変更項目", "旧値", "新値", "旧更新期限", "新更新期限", "コピー有無", "状態", "備考", "元資格ID", "資格更新履歴ID"],
    "資格マスタ": ["表示順", "区分", "資格名", "期限管理", "有効年数", "通知日前", "一覧表示", "備考", "資格マスタID"],
    "資格プルダウン補助": ["区分", "全資格", "資格", "技能講習", "特別教育", "安全教育", "免許", "その他", "取得区分", "保有状況", "コピー有無", "状態"],
    "資格表記ゆれマスタ": ["入力名", "正式資格名", "区分", "取得区分", "備考"],
    "資格かんたん登録": ["所有者", "区分", "追加する資格", "資格名まとめ", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "備考", "登録結果"],
    "資格一括登録": ["所有者", "区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "備考", "登録結果"],
    "資格重複チェック": ["判定", "所有者", "区分", "資格名", "取得区分", "保有状況", "件数", "元行", "備考"],
    "社員別資格一覧": ["区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"],
    "資格別保有者一覧": ["資格名", "区分", "所有者", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"],
    "備品修理管理": ["購入日", "修理依頼日", "返却予定日", "修理依頼者", "修理業者", "備品名", "内容", "返却済", "状態", "通知", READ_HEADER, ...personal, "備考", "修理ID"],
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
  return ["一覧スケジュール", "要確認一覧", "担当別未読", "既読率集計", "ダッシュボード", "過去一覧", "帳簿PDF履歴", "帳簿出力設定", "SQL設計", "移行対応表", "SQLサンプル集", "手順シート", "社内管理説明", "機能チェック", "GPT診断用", "資格マスタ", "資格表記ゆれマスタ", "資格プルダウン補助", "資格かんたん登録", "資格一括登録", "資格重複チェック", "社員別資格一覧", "資格別保有者一覧"].includes(sheetName);
}

function getStatusListForSheet_(sheetName) {
  const map = {
    "出先予定": [CLEAR_LABEL, "予定", "移動中", "完了", "延期", "中止"],
    "工事予定": [CLEAR_LABEL, "予定", "未契約", "着工前", "施工中", "完了", "請求済み", "延期", "中止"],
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
  const carList = [CLEAR_LABEL, ...getCompanyVehicles_()];
  const vehicleNameList = [CLEAR_LABEL, ...getVehicleNames_()];
  const repairVendorList = [CLEAR_LABEL, ...getRepairVendors_()];
  const phoneList = [CLEAR_LABEL, "未対応", "対応中", "折返し", "完了"];
  const importanceList = [CLEAR_LABEL, "高", "中", "低"];
  const copyList = [CLEAR_LABEL, "有", "無", "未確認"];
  const returnedList = [CLEAR_LABEL, "未", "済"];
  const receiptCategoryList = [CLEAR_LABEL, "燃料", "駐車場", "高速代", "消耗品", "資材", "工具", "修理", "その他"];
  const qualificationCategoryList = [CLEAR_LABEL, "資格", "技能講習", "特別教育", "安全教育", "免許", "その他"];
  const qualificationPossessionList = [CLEAR_LABEL, "技士", "技士補", "保有", "修了", "免許"];
  const qualificationHoldingStatusList = [CLEAR_LABEL, "保有", "未確認", "失効", "不要"];
  const paymentList = [CLEAR_LABEL, "現金", "会社カード", "立替", "請求書", "その他"];
  const accountingList = [CLEAR_LABEL, "未確認", "確認中", "確認済", "差戻し"];
  const settlementList = [CLEAR_LABEL, "未精算", "精算中", "精算済", "対象外"];
  const feedbackCategoryList = [CLEAR_LABEL, "入力しづらい", "項目不足", "表示不備", "動作不備", "要望", "その他"];
  const difficultyList = [CLEAR_LABEL, "低", "中", "高", "至急"];
  const checkMethodList = [CLEAR_LABEL, "PC", "AppSheet", "紙", "口頭", "その他"];

  if (["担当", "入力者", "利用者", "投稿者", "記入者", "所有者", "修理依頼者"].includes(header)) return staffList;
  if (sheetName === "工事予定" && header === "税") return [CLEAR_LABEL, "税込", "税抜", "不明"];
  if (header === "車両名" && sheetName === "車検管理") return vehicleNameList;
  if (sheetName === DRIVER_LICENSE_DETAIL_SHEET_NAME && header === "免許種類") return [CLEAR_LABEL, ...LICENSE_TYPE_HEADERS];
  if (sheetName === DRIVER_LICENSE_DETAIL_SHEET_NAME && header === "保有状況") return [CLEAR_LABEL, "保有", "未確認", "失効", "不要"];
  if (header === "社用車") return carList;
  if (header === "修理業者") return repairVendorList;
  if (header === "電話対応") return phoneList;
  if (header === "重要度") return importanceList;
  if (header === "コピー有無") return copyList;
  if (header === "返却済") return returnedList;
  if (header === "区分" && sheetName === "日報レシート管理") return receiptCategoryList;
  if (header === "区分" && sheetName === "資格管理") return qualificationCategoryList;
  if (header === "資格名" && sheetName === "資格管理") return [CLEAR_LABEL, ...getQualificationMasterNameList_()];
  if (header === "取得区分" && sheetName === "資格管理") return qualificationPossessionList;
  if (header === "保有状況" && sheetName === "資格管理") return qualificationHoldingStatusList;
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

  // 既存データを消さない。ヘッダーだけ不足時に修復する。
  if (current[0] !== "設定種別" || current[1] !== "値") {
    const existingLastRow = sheet.getLastRow();
    const existingLastCol = Math.max(sheet.getLastColumn(), 3);
    let existing = [];
    if (existingLastRow >= 1) {
      existing = sheet.getRange(1, 1, existingLastRow, existingLastCol).getValues();
    }
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 3).setValues([headers]);

    // 旧ヘッダーが無い状態で値が入っていた場合だけ、2行目以降へ戻す。
    const rows = [];
    existing.forEach((row, index) => {
      if (index === 0 && (String(row[0] || "").trim() === "設定種別" || String(row[1] || "").trim() === "値")) return;
      const type = String(row[0] || "").trim();
      const value = String(row[1] || "").trim();
      const note = String(row[2] || "").trim();
      if (type || value || note) rows.push([type, value, note]);
    });
    if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  // サンプル担当者は自動補充しない。担当者・既読確認者は設定シートを正とする。
  // 社用車・車両名だけは、該当種別が1件も無い場合に限り初期候補を入れる。
  ensureDefaultSettingRows_("担当者", DEFAULT_STAFF_MEMBERS, "担当プルダウン用");
  ensureDefaultSettingRows_("既読確認者", DEFAULT_PERSONAL_MEMBERS, "個人確認列用");
  ensureDefaultSettingRows_("社用車", DEFAULT_COMPANY_VEHICLES, "社用車プルダウン用");
  ensureDefaultSettingRows_("車両名", DEFAULT_COMPANY_VEHICLES, "車検管理用");

  formatSheetBase_(sheet, headers.length);
  return sheet;
}

function ensureDefaultSettingRows_(settingType, defaults, note) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  if (!sheet) return;

  const list = (defaults || []).map(v => String(v || "").trim()).filter(Boolean);
  if (!list.length) return;

  const lastRow = sheet.getLastRow();

  // その設定種別が既に1件でもある場合は、初期値を追加しない。
  // これにより、会社用の担当者・車両名を入れた後にサンプル候補が戻る事故を防ぐ。
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const hasExistingType = values.some(r => {
      const type = String(r[0] || "").trim();
      const value = String(r[1] || "").trim();
      return type === settingType && value;
    });
    if (hasExistingType) return;
  }

  const rows = list.map(value => [settingType, value, note]);
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

function removeSampleStaffSettings() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const result = removeSampleStaffSettings_();
    toast_("設定シートからサンプル担当者を削除しました（担当者 " + result.staff + "件、既読確認者 " + result.personal + "件）。必要な会社用担当者を設定シートに入力してください。");
  } finally {
    lock.releaseLock();
  }
}

function removeSampleStaffSettings_() {
  const sheet = ensureSettingsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {staff: 0, personal: 0};

  const sampleStaff = {};
  SAMPLE_STAFF_MEMBERS_FOR_CLEANUP.forEach(v => sampleStaff[normalizeTextForKey_(v)] = true);
  const samplePersonal = {};
  SAMPLE_PERSONAL_MEMBERS_FOR_CLEANUP.forEach(v => samplePersonal[normalizeTextForKey_(v)] = true);

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const keep = [];
  let staff = 0;
  let personal = 0;

  values.forEach(row => {
    const type = String(row[0] || "").trim();
    const value = String(row[1] || "").trim();
    const key = normalizeTextForKey_(value);
    if (type === "担当者" && sampleStaff[key]) {
      staff++;
      return;
    }
    if (type === "既読確認者" && samplePersonal[key]) {
      personal++;
      return;
    }
    keep.push(row);
  });

  sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, 3).setValues(keep);
  clearSettingsCache_();
  return {staff: staff, personal: personal};
}


function cleanupCompanySampleStaffAndPersonalColumns() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const result = removeSampleStaffSettings_();
    const colResult = rebuildPersonalCheckColumnsFromSettings_();
    toast_(
      "設定シートと既存シートのサンプル整理が完了しました（設定: 担当者 "
      + result.staff + "件、既読確認者 " + result.personal
      + "件 / 個人確認列: 削除 " + colResult.deleted
      + "列、追加 " + colResult.inserted + "列）。"
    );
  } finally {
    lock.releaseLock();
  }
}

function rebuildPersonalCheckColumnsFromSettings() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const result = rebuildPersonalCheckColumnsFromSettings_();
    toast_("個人確認列を設定シート基準で整理しました（削除 " + result.deleted + "列、追加 " + result.inserted + "列）。");
  } finally {
    lock.releaseLock();
  }
}

function rebuildPersonalCheckColumnsFromSettings_() {
  clearSettingsCache_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = getPersonalMembers_();
  let deleted = 0;
  let inserted = 0;

  getInputSheetNames_().forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastColumn() < 1) return;

    const headers = getHeaders_(sheet);
    if (!headers.length || headers.indexOf(READ_HEADER) < 0) return;

    deleted += deleteObsoletePersonalCheckColumns_(sheet, members);
    inserted += ensurePersonalCheckColumns_(sheet, members);

    clearColumnGroupsForSheet_(sheet);
    applyPersonalCheckColumnLayout_(sheet);
    setCheckboxesForDataRows(sheet);
    rebuildPersonalCheckGroupForSheet_(sheet);
  });

  return {deleted: deleted, inserted: inserted};
}

function deleteObsoletePersonalCheckColumns_(sheet, members) {
  const memberSet = {};
  (members || []).forEach(v => memberSet[normalizeTextForKey_(v)] = true);

  let headers = getHeaders_(sheet);
  let readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol <= 0) return 0;

  const sheetName = sheet.getName();
  const anchors = getPersonalCheckAnchorHeadersForSheet_(sheetName);
  let nextAnchorCol = 0;
  anchors.forEach(anchor => {
    const col = headers.indexOf(anchor) + 1;
    if (col > readCol && (!nextAnchorCol || col < nextAnchorCol)) nextAnchorCol = col;
  });

  const sampleSet = {};
  SAMPLE_PERSONAL_MEMBERS_FOR_CLEANUP.forEach(v => sampleSet[normalizeTextForKey_(v)] = true);

  const deleteCols = [];
  if (nextAnchorCol > readCol + 1) {
    for (let col = readCol + 1; col < nextAnchorCol; col++) {
      const header = String(headers[col - 1] || "").trim();
      if (!header) continue;
      const key = normalizeTextForKey_(header);
      if (!memberSet[key]) deleteCols.push(col);
    }
  } else {
    // アンカーが見つからない異常時は、山田/高橋/鈴木など旧サンプル列だけを消す。
    headers.forEach((header, index) => {
      const key = normalizeTextForKey_(header);
      if (sampleSet[key] && !memberSet[key]) deleteCols.push(index + 1);
    });
  }

  let count = 0;
  deleteCols.sort((a, b) => b - a).forEach(col => {
    try {
      sheet.deleteColumn(col);
      count++;
    } catch (e) {
      console.log(sheet.getName() + " の旧個人確認列削除をスキップ: col=" + col + " " + e.message);
    }
  });
  return count;
}

function ensurePersonalCheckColumns_(sheet, members) {
  if (!members || !members.length) return 0;

  let inserted = 0;
  let headers = getHeaders_(sheet);
  let readCol = headers.indexOf(READ_HEADER) + 1;
  if (readCol <= 0) return 0;

  let insertAfter = readCol;
  members.forEach(member => {
    const name = String(member || "").trim();
    if (!name) return;

    headers = getHeaders_(sheet);
    const existingCol = headers.indexOf(name) + 1;
    if (existingCol > 0) {
      insertAfter = existingCol;
      return;
    }

    try {
      sheet.insertColumnAfter(insertAfter);
      sheet.getRange(1, insertAfter + 1).setValue(name);
      inserted++;
      insertAfter++;
    } catch (e) {
      console.log(sheet.getName() + " の個人確認列追加をスキップ: " + name + " " + e.message);
    }
  });

  return inserted;
}

function getPersonalCheckAnchorHeadersForSheet_(sheetName) {
  const idHeader = SHEET_ID_HEADERS[sheetName] || "";
  return uniqueList_(["備考", "PDFリンク", CALENDAR_ID_HEADER, idHeader].filter(Boolean));
}

function clearColumnGroupsForSheet_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return;
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (e) {}
  for (let i = 0; i < 10; i++) {
    try {
      sheet.getRange(1, 1, 1, sheet.getLastColumn()).shiftColumnGroupDepth(-1);
    } catch (e) {
      break;
    }
  }
}

function repairCompanyVehicleSettings() {
  const sheet = ensureSettingsSheet_();
  ensureDefaultSettingRows_("社用車", DEFAULT_COMPANY_VEHICLES, "社用車プルダウン用");
    ensureDefaultSettingRows_("車両名", DEFAULT_COMPANY_VEHICLES, "車検管理用");
  clearSettingsCache_();

  formatSheetBase_(sheet, 3);
  try {
    const lastRow = Math.max(sheet.getLastRow(), 2);
    sheet.getRange(1, 1, lastRow, 3).setWrap(true).setVerticalAlignment("middle");
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 180);
    createFilterSafely_(sheet, 3);
  } catch (e) {}

  toast_("設定シートに社用車一覧を追加/修復しました。続けて『入力シートだけ設定反映』を実行してください。");
}


let SETTINGS_CACHE_ = null;
let QUALIFICATION_DROPDOWN_HELPER_READY_ = false;
let QUALIFICATION_VALIDATION_LISTS_CACHE_ = null;
function clearSettingsCache_() {
  SETTINGS_CACHE_ = null;
  QUALIFICATION_VALIDATION_LISTS_CACHE_ = null;
}
function clearQualificationValidationListsCache_() {
  QUALIFICATION_VALIDATION_LISTS_CACHE_ = null;
}
function getSettingsCache_() {
  if (SETTINGS_CACHE_) return SETTINGS_CACHE_;

  const cache = {
    "担当者": DEFAULT_STAFF_MEMBERS.slice(),
    "既読確認者": DEFAULT_PERSONAL_MEMBERS.slice(),
    "社用車": DEFAULT_COMPANY_VEHICLES.slice(),
    "車両名": DEFAULT_COMPANY_VEHICLES.slice(),
    "修理業者": []
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  if (!sheet || sheet.getLastRow() < 2) {
    SETTINGS_CACHE_ = cache;
    return cache;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const staff = [];
  const personal = [];
  const vehicles = [];
  const vehicleNames = [];
  const repairVendors = [];

  values.forEach(r => {
    const type = String(r[0] || "").trim();
    const value = String(r[1] || "").trim();
    if (!value) return;

    if (type === "担当者") staff.push(value);
    if (type === "既読確認者") personal.push(value);

    // 「社用車」と「社用車一覧」の両方を許可。既存設定シートの表記ゆれ対策。
    if (type === "社用車" || type === "社用車一覧") vehicles.push(value);
    if (type === "車両名" || type === "車両名一覧") vehicleNames.push(value);
    if (type === "修理業者" || type === "修理業者一覧") repairVendors.push(value);
  });

  if (staff.length) cache["担当者"] = uniqueList_(staff);
  if (personal.length) cache["既読確認者"] = uniqueList_(personal);
  if (vehicles.length) cache["社用車"] = uniqueList_(vehicles);
  if (vehicleNames.length) cache["車両名"] = uniqueList_(vehicleNames);
  else if (vehicles.length) cache["車両名"] = uniqueList_(vehicles);
  if (repairVendors.length) cache["修理業者"] = uniqueList_(repairVendors);

  SETTINGS_CACHE_ = cache;
  return cache;
}
function getStaffMembers_() { return getSettingsCache_()["担当者"].slice(); }
function getPersonalMembers_() { return getSettingsCache_()["既読確認者"].slice(); }
function getCompanyVehicles_() { return getSettingsCache_()["社用車"].slice(); }
function getVehicleNames_() { return getSettingsCache_()["車両名"].slice(); }
function getRepairVendors_() { return getSettingsCache_()["修理業者"].slice(); }

function uniqueList_(list) {
  const seen = {};
  const result = [];
  (list || []).forEach(v => {
    const value = String(v || "").trim();
    if (!value) return;
    if (seen[value]) return;
    seen[value] = true;
    result.push(value);
  });
  return result;
}

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
  // p20: 内部ID列・カレンダーID列は作成直後から通常表示では隠す。
  hideSystemColumnsForSheet_(sheet);
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

  // v9.7.48m:
  // 車検履歴は履歴シートだが、日付列と担当列の入力規則だけは明示的に修復する。
  // 旧版で「担当」列だった場所に「車検予定日」を追加したため、
  // 日付列に担当者プルダウンが残る事故をここで防ぐ。
  if (sheetName === "車検履歴") {
    applyVehicleInspectionHistoryValidations_(sheet);
    return;
  }

  headers.forEach((header, index) => {
    const col = index + 1;

    // 対応完了日時は日付だけでなく時刻も残すため、日付列とは別扱いにする。
    if (header === "対応完了日時") {
      setDateTimeColumn(sheet, col);
      return;
    }

    if (DATE_HEADERS.includes(header)) setDateColumn(sheet, col);
    if (TIME_HEADERS.includes(header)) setTimeColumn(sheet, col);
    if (VEHICLE_NUMBER_HEADERS.includes(header)) setTextColumn(sheet, col);

    // v9.7.48p:
    // 工事予定の契約金額は円表示、税込/税抜/不明は隣の「税」列で選択する。
    // 金額欄は数値を推奨するが、未定などの文字入力を完全には弾かない。
    if (sheetName === "工事予定" && header === "契約金額") {
      setConstructionAmountCurrencyColumn_(sheet, col);
      return;
    }

    if (NUMBER_HEADERS.includes(header)) setNumberColumn(sheet, col, header);

    // v9.7.48h:
    // 資格管理のプルダウンは、資格マスタから作った「資格プルダウン補助」シートの範囲参照で設定する。
    // 値リスト直書きより安定し、候補が増えてもプルダウンが消えにくい。
    if (sheetName === "資格管理" && isQualificationManagementDropdownHeader_(header)) {
      applyQualificationManagementColumnValidation_(sheet, header, col);
      return;
    }

    if (sheetName === "備品修理管理" && header === "修理業者") {
      setDropdownAllowInvalid(sheet, col, [CLEAR_LABEL, ...getRepairVendors_()]);
      return;
    }

    const list = getDropdownListForHeader_(sheetName, header);
    if (list) setDropdown(sheet, col, list);
  });

  // p24:
  // 通常の「入力シートだけ設定反映」でも、資格管理C列はB列区分ごとの候補へ戻す。
  // これを入れないと、資格名列が全資格候補になり、区分ソートが崩れて見える。
  if (sheetName === "資格管理") {
    try {
      applyQualificationManagementNameValidationsStrictByRow_(sheet);
    } catch (e) {
      console.log("資格管理の行別資格名プルダウン再設定をスキップ: " + e.message);
    }
  }

  setCheckboxesForDataRows(sheet);
}

function setDateColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  sheet.getRange(2, col, rowCount).setNumberFormat("yyyy/mm/dd").setDataValidation(rule);
}

function setDateTimeColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount)
    .setNumberFormat("yyyy/mm/dd hh:mm")
    .clearDataValidations();
}

function setTimeColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount).setNumberFormat("hh:mm");
}

function setTextColumn(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount).setNumberFormat("@").clearDataValidations();
}

function setTextColumnAllCurrentRows_(sheet, col) {
  // p39: 連絡先など、先頭0を守りたい列は現在の最大行まで文字列固定にする。
  // getApplyRowCount_ だけだと、行追加後や下の方の入力で 0128 の 0 が落ちることがあるため。
  if (!sheet || col <= 0) return 0;
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, col, rows, 1).setNumberFormat("@").clearDataValidations();
  return rows;
}

function setConstructionAmountCurrencyColumn_(sheet, col) {
  const rowCount = getApplyRowCount_(sheet);
  sheet.getRange(2, col, rowCount)
    .setNumberFormat("¥#,##0")
    .clearDataValidations();
}

// 旧関数名互換。v9.7.48p以降は文字列固定ではなく円表示にする。
function setConstructionAmountTextColumn_(sheet, col) {
  setConstructionAmountCurrencyColumn_(sheet, col);
}

function setNumberColumn(sheet, col, header) {
  const rowCount = getApplyRowCount_(sheet);
  const format = String(header || "") === "契約金額" || String(header || "") === "金額" ? "¥#,##0" : "#,##0";
  sheet.getRange(2, col, rowCount).setNumberFormat(format);
}

function setDropdown(sheet, col, list) {
  const rowCount = getApplyRowCount_(sheet);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
  sheet.getRange(2, col, rowCount).setDataValidation(rule);
}

function setDropdownAllowInvalid(sheet, col, list) {
  const rowCount = getApplyRowCount_(sheet);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
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
  // 工事予定は専用の通知判定を使う。
  // ここで getNoticeText(開始日, 状態) を使うと、工事予定シート本体と一覧スケジュールで通知がズレる。
  collectRowsAsObjects(ss, "工事予定", r => r["開始日"], r => scheduleRow_(r["開始日"], "工事予定", joinText(r["工事名"], joinText(r["現場"], r["依頼主"])), r["担当"], r["状態"], getConstructionNoticeText_(r), r["電話対応"], r, "工事予定"), rows);
  collectRowsAsObjects(ss, "会議予定", r => r["日付"], r => scheduleRow_(r["日付"], "会議", joinText(r["会議名"], r["内容"]), r["担当"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "会議予定"), rows);
  collectRowsAsObjects(ss, "行事予定", r => r["日付"], r => scheduleRow_(r["日付"], "行事", buildMultiDayEventContent_(r["行事名"], r["内容"], r["日付"], r["終了日"]), r["担当"], r["状態"], getPeriodNoticeText_(r["日付"], r["終了日"], r["状態"]), "", r, "行事予定"), rows);
  collectRowsAsObjects(ss, "電話履歴", r => r["日付"] || r["相手"], r => scheduleRow_(r["日付"] || new Date(), "電話履歴", joinText(r["相手"], joinText(r["折り返し電話番号"], r["内容"])), r["担当"], r["電話対応"], getPhoneNoticeText_(r["日付"], r["電話対応"]), r["電話対応"], r, "電話履歴"), rows);
  collectRowsAsObjects(ss, "車検管理", r => r["車検期限"] || r["車検予定日"], r => scheduleRow_(r["車検予定日"] || r["車検期限"], "車検", joinText(r["車両名"], r["車番"]), "", r["車検状態"], getNoticeText(r["車検予定日"] || r["車検期限"], r["車検状態"]), "", r, "車検管理"), rows);
  collectRowsAsObjects(ss, "社用車予約", r => r["日付"] || r["社用車"], r => scheduleRow_(r["日付"] || new Date(), "社用車予約", joinText(r["社用車"], joinText(r["行き先"], r["用途"])), r["利用者"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "社用車予約"), rows);
  collectRowsAsObjects(ss, "日報", r => r["日付"] || r["現場"], r => scheduleRow_(r["日付"] || new Date(), "日報", joinText(r["現場"], r["作業内容"]), r["担当"] || r["入力者"], r["状態"], getNoticeText(r["日付"], r["状態"]), "", r, "日報"), rows);
  collectRowsAsObjects(ss, "運転免許管理", r => getDriverLicenseExpiryDate_(r) || r["所有者"], r => scheduleRow_(getDriverLicenseExpiryDate_(r) || new Date(), "運転免許", joinText(r["所有者"], buildLicenseTypeText_(r)), r["所有者"], r["状態"], getNoticeText(getDriverLicenseExpiryDate_(r), r["状態"]), "", r, "運転免許管理"), rows);
  // p13: 資格管理は件数が多く、一覧スケジュールを埋めるため通常表示しない。
  // 更新期限が近い資格だけは createAlertList() 側で要確認一覧へ直接出す。
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
  appendQualificationAlerts_(ss, rows);

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

// p13: 資格管理は一覧スケジュールには出さず、期限系だけ要確認一覧へ直接出す。
function appendQualificationAlerts_(ss, rows) {
  const sheet = ss.getSheetByName("資格管理");
  if (!sheet || sheet.getLastRow() < 2) return;

  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach(v => {
    const r = objectFromRow(headers, v);
    if (!isBusinessDataRow_(headers, v, "資格管理")) return;

    const status = String(r["状態"] || "").trim();
    if (["有効", "更新済"].includes(status)) return;

    const notice = getNoticeText(r["更新期限"], status);
    if (!["期限切れ", "今日", "3日以内", "7日以内"].includes(notice)) return;

    rows.push({
      "日付": r["更新期限"] || new Date(),
      "種類": "資格",
      "内容": joinText(r["所有者"], joinText(r["資格名"], joinText(r["区分"], joinText(r["取得区分"], r["保有状況"])))),
      "担当": r["所有者"],
      "状態": status,
      "通知": notice,
      "電話対応": "",
      "備考": r["備考"] || "",
      "元シート": "資格管理"
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
  if (sheetName === "運転免許管理") return getNoticeText(getDriverLicenseExpiryDate_(row), row["状態"]);
  if (sheetName === "資格管理") return getNoticeText(row["更新期限"], row["状態"]);
  if (sheetName === "備品修理管理") return getNoticeText(row["返却予定日"] || row["修理依頼日"], row["状態"]);
  if (sheetName === "個人ToDo") return getNoticeText(row["期限"], row["状態"]);
  if (sheetName === "工事予定") return getConstructionNoticeText_(row);
  if (sheetName === "行事予定") return getPeriodNoticeText_(row["日付"], row["終了日"], row["状態"]);
  if (sheetName === "電話履歴") return getPhoneNoticeText_(row["日付"], row["電話対応"]);
  if (sheetName === "会議予定" || sheetName === "出先予定" || sheetName === "社用車予約" || sheetName === "日報") return getNoticeText(row["日付"], row["状態"] || row["電話対応"]);
  return "";
}



function getConstructionNoticeText_(row) {
  const status = String(row["状態"] || "").trim();

  // 工事予定では、開始日を過ぎた行を「期限切れ」にしない。
  // 予定の開始日を過ぎたものは、一覧・要確認上では完了扱いに寄せる。
  // これにより、同じ工事予定で「完了」になったり「期限切れ」になったりするブレを防ぐ。
  if (["施工中", "完了", "請求済み", "更新済", "確認済"].includes(status)) return "完了";
  if (["中止"].includes(status)) return "";

  const start = toDateOnly_(row["開始日"]);
  if (!start) return "";

  const today = toDateOnly_(new Date());
  const diff = Math.floor((start.getTime() - today.getTime()) / (24 * 60 * 60 * 1000));

  if (diff < 0) return "完了";
  if (diff === 0) return "今日";
  if (diff <= 3) return "3日以内";
  if (diff <= 7) return "7日以内";
  if (diff <= 30) return "30日以内";
  return "";
}


function getPeriodNoticeText_(startValue, endValue, status) {
  const s = String(status || "").trim();
  if (["完了", "請求済み", "更新済", "確認済", "返却済", "精算済"].includes(s)) return "完了";
  if (["中止", "修理不可"].includes(s)) return "";

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

  const result = processVehicleInspectionUpdatedRows_(sheet, 2, sheet.getLastRow(), false);
  toast_("車検更新済を処理しました（履歴追加: " + result.added + " / 状態空欄戻し: " + result.cleared + "）");
}


function handleVehicleInspectionHistoryOnEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "車検管理") return;

  // 複数セル貼り付けでは旧値が安全に取れないため、自動処理は単一セル編集だけ対象。
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= 1) return;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[col - 1];

  // 車検期限を変更した時点では、履歴へ即追加せず「旧期限」を一時保存する。
  // その後、車検状態を「更新済」にした時に、履歴追加＋状態空欄戻しを同時に行う。
  if (editedHeader !== "車検期限") return;

  const oldDue = parseDateForVehicleHistory_(e.oldValue);
  const newDue = toDateOnly_(e.range.getValue());
  if (!oldDue || !newDue) return;
  if (oldDue.getTime() === newDue.getTime()) return;

  rememberVehicleInspectionDueChange_(sheet, row, oldDue, newDue);
}



function processVehicleInspectionUpdatedRowsOnEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "車検管理") return;

  // 車検管理では、どの列を編集しても「更新済」行があれば処理する。
  // これにより、車検状態を選んだ瞬間だけでなく、既に更新済が残っている行も次の編集で処理できる。
  const editedRow = e.range.getRow();
  if (editedRow <= 1) return;

  // 車検期限を直接変更した場合は、一応旧期限を保存する。
  // 基本運用は「車検期限=旧期限、車検予定日=新期限、車検状態=更新済」。
  try {
    handleVehicleInspectionHistoryOnEdit_(e);
  } catch (err) {
    console.log("車検期限旧値の一時保存をスキップ: " + err.message);
  }

  processVehicleInspectionUpdatedRows_(sheet, editedRow, editedRow, true);
}

function handleVehicleInspectionStatusEdit_(e) {
  if (!e || !e.range) return false;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "車検管理") return false;

  // 複数セル貼り付けは自動処理せず、保守メニューでまとめて処理する。
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const row = e.range.getRow();
  if (row <= 1) return false;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];

  // 車検期限の直接変更時は、旧期限を一時保存だけ行う。
  // その後、車検状態を更新済にしたときに履歴化する。
  try {
    handleVehicleInspectionHistoryOnEdit_(e);
  } catch (err) {
    console.log("車検期限旧値の一時保存をスキップ: " + err.message);
  }

  const statusCol = headers.indexOf("車検状態") + 1;
  const dueCol = headers.indexOf("車検期限") + 1;
  const plannedCol = headers.indexOf("車検予定日") + 1;
  const newDueCol = headers.indexOf("新車検期限") + 1;
  const noticeCol = headers.indexOf("通知") + 1;

  if (statusCol <= 0 || dueCol <= 0 || plannedCol <= 0) return false;

  // 基本は「車検状態」を更新済にした瞬間だけ自動処理する。
  // ただし既に更新済が残っている行は、車検予定日や車検期限を編集したときにも処理できるようにする。
  const status = String(sheet.getRange(row, statusCol).getValue() || "").trim();
  if (status !== "更新済") return false;
  if (!["車検状態", "車検予定日", "車検期限", "新車検期限"].includes(editedHeader)) return false;

  const result = processVehicleInspectionUpdatedRow_(sheet, row);

  // 履歴追加できた場合だけ、更新済を空欄へ戻す。
  // 新車検期限が未入力の場合に勝手に空欄へ戻すと、処理されたように見えるため残す。
  if (result.added) {
    sheet.getRange(row, statusCol).clearContent();
  }

  // 通知を新しい車検期限で再計算する。
  try {
    const newDue = toDateOnly_(sheet.getRange(row, dueCol).getValue());
    if (noticeCol > 0) sheet.getRange(row, noticeCol).setValue(newDue ? getNoticeText(newDue) : "");
  } catch (err) {}

  // 車検予定日が残っていると「処理されていない」ように見えるため、履歴追加できた場合は必ず消す。
  // processVehicleInspectionUpdatedRow_ 側でも消すが、ここでも保険として消す。
  if (result.added) {
    try { sheet.getRange(row, plannedCol).clearContent(); } catch (err) {}
  }

  sortVehicleInspectionSheetByDue_();
  createScheduleList();
  createAlertList();
  clearSummaryDirty_();

  if (result.added) {
    toast_("【車検更新済処理】履歴追加・期限更新・状態空欄戻しを行いました");
  } else {
    toast_("【車検更新済処理】履歴は未追加です（" + (result.reason || "条件不一致") + "）");
  }

  return true;
}

function processVehicleInspectionUpdatedRows_(sheet, startRow, endRow, fromEdit) {
  if (!sheet || sheet.getName() !== "車検管理") return { added: 0, cleared: 0 };

  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf("車検状態") + 1;
  const dueCol = headers.indexOf("車検期限") + 1;
  const noticeCol = headers.indexOf("通知") + 1;

  if (statusCol <= 0 || sheet.getLastRow() < 2) return { added: 0, cleared: 0 };

  const first = Math.max(Number(startRow || 2), 2);
  const last = Math.min(Number(endRow || sheet.getLastRow()), sheet.getLastRow());

  let addedCount = 0;
  let clearedCount = 0;
  let lastReason = "";

  for (let row = first; row <= last; row++) {
    const status = String(sheet.getRange(row, statusCol).getValue() || "").trim();
    if (status !== "更新済") continue;

    const result = processVehicleInspectionUpdatedRow_(sheet, row);
    if (result.added) addedCount++;
    if (result.reason) lastReason = result.reason;

    if (result.added) {
      sheet.getRange(row, statusCol).clearContent();
      clearedCount++;
    }

    try {
      if (noticeCol > 0 && dueCol > 0) {
        const newDue = toDateOnly_(sheet.getRange(row, dueCol).getValue());
        sheet.getRange(row, noticeCol).setValue(newDue ? getNoticeText(newDue) : "");
      }
    } catch (e) {}
  }

  if (clearedCount > 0) {
    sortVehicleInspectionSheetByDue_();

    if (fromEdit) {
      if (addedCount > 0) {
        toast_("【車検更新済処理】履歴追加: " + addedCount + " / 状態空欄戻し: " + clearedCount);
      } else {
        toast_("【車検更新済処理】状態は空欄に戻しました。履歴は未追加です（" + (lastReason || "条件不一致") + "）");
      }
    }
  }

  return { added: addedCount, cleared: clearedCount };
}

function handleVehicleInspectionUpdatedAndHistoryOnEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "車検管理") return;

  // 先に車検期限変更を一時保存
  handleVehicleInspectionHistoryOnEdit_(e);

  // 次に「更新済」なら履歴追加＋空欄戻し
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= 1) return;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[col - 1];
  if (editedHeader !== "車検状態") return;

  const newStatus = String(e.range.getValue() || "").trim();
  if (newStatus !== "更新済") return;

  normalizeVehicleInspectionUpdatedRow_(sheet, row);
}

function getVehicleInspectionChangeKey_(vehicleSheet, row) {
  const headers = getHeaders_(vehicleSheet);
  const values = vehicleSheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = objectFromRow(headers, values);

  const vehicleId = obj["車両ID"] || "";
  const vehicleName = obj["車両名"] || "";
  const vehicleNumber = obj["車番"] || "";

  const base = vehicleId || (String(vehicleName) + "__" + String(vehicleNumber)) || ("row_" + row);
  return "VEHICLE_DUE_CHANGE__" + base;
}

function rememberVehicleInspectionDueChange_(vehicleSheet, row, oldDue, newDue) {
  if (!vehicleSheet || vehicleSheet.getName() !== "車検管理") return;

  const headers = getHeaders_(vehicleSheet);
  const values = vehicleSheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = objectFromRow(headers, values);

  const payload = {
    row: row,
    vehicleId: obj["車両ID"] || "",
    vehicleName: obj["車両名"] || "",
    vehicleNumber: obj["車番"] || "",
    oldDue: formatDateForVehicleHistoryKey_(oldDue),
    newDue: formatDateForVehicleHistoryKey_(newDue),
    rememberedAt: new Date().toISOString()
  };

  const key = getVehicleInspectionChangeKey_(vehicleSheet, row);
  PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(payload));
}



function applyVehicleInspectionUpdateForActiveRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || sheet.getName() !== "車検管理") {
    toast_("車検管理シートで対象行を選択してから実行してください");
    return;
  }

  const range = sheet.getActiveRange();
  if (!range) {
    toast_("対象行を選択してください");
    return;
  }

  const row = range.getRow();
  if (row <= 1) {
    toast_("ヘッダー行ではなく、車検管理のデータ行を選択してください");
    return;
  }

  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf("車検状態") + 1;
  const dueCol = headers.indexOf("車検期限") + 1;
  const noticeCol = headers.indexOf("通知") + 1;

  // 選択行は、車検状態が空欄でも強制的に処理できる。
  const result = processVehicleInspectionUpdatedRow_(sheet, row);

  if (result.added && statusCol > 0) sheet.getRange(row, statusCol).clearContent();

  try {
    if (noticeCol > 0 && dueCol > 0) {
      const newDue = toDateOnly_(sheet.getRange(row, dueCol).getValue());
      sheet.getRange(row, noticeCol).setValue(newDue ? getNoticeText(newDue) : "");
    }
  } catch (e) {}

  sortVehicleInspectionSheetByDue_();

  if (result.added) {
    toast_("【車検履歴処理】車検履歴へ追加し、車検期限を更新しました");
  } else {
    toast_("【車検履歴処理】履歴は未追加です（" + (result.reason || "条件不一致") + "）");
  }
}

function processVehicleInspectionUpdatedRow_(vehicleSheet, row) {
  if (!vehicleSheet || vehicleSheet.getName() !== "車検管理" || row <= 1) {
    return { added: false, reason: "車検管理のデータ行ではありません" };
  }

  const headers = getHeaders_(vehicleSheet);
  const values = vehicleSheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = objectFromRow(headers, values);

  const dueCol = headers.indexOf("車検期限") + 1;
  const plannedCol = headers.indexOf("車検予定日") + 1;
  const newDueCol = headers.indexOf("新車検期限") + 1;

  const oldDueCurrent = toDateOnly_(obj["車検期限"]);
  const scheduledDate = toDateOnly_(obj["車検予定日"]);
  const explicitNewDue = toDateOnly_(obj["新車検期限"]);

  // v9.7.48l:
  // 車検予定日は「車検に出す予定日」であり、新車検期限ではない。
  // 更新済処理は、原則として「新車検期限」に入力された日付を次回期限として扱う。
  // 処理成功時だけ、車検期限を新車検期限へ更新し、車検予定日・新車検期限・車検状態を空欄へ戻す。
  if (explicitNewDue) {
    if (!oldDueCurrent) {
      return { added: false, reason: "旧車検期限が空欄のため履歴化できません" };
    }

    const added = appendVehicleInspectionHistoryFromRow_(vehicleSheet, row, oldDueCurrent, explicitNewDue, "更新済", scheduledDate);

    if (added) {
      if (dueCol > 0) vehicleSheet.getRange(row, dueCol).setValue(explicitNewDue).setNumberFormat("yyyy/mm/dd");
      if (plannedCol > 0) vehicleSheet.getRange(row, plannedCol).clearContent();
      if (newDueCol > 0) vehicleSheet.getRange(row, newDueCol).clearContent();
      clearRememberedVehicleInspectionDueChange_(vehicleSheet, row);
    }

    return { added: added, reason: added ? "" : "同じ履歴が既にある可能性があります" };
  }

  // 旧方式互換:
  // 車検期限セルを直接新期限へ変更してから更新済にした場合は、oldValue/newValueの一時保存から履歴化する。
  const remembered = getRememberedVehicleInspectionDueChange_(vehicleSheet, row);
  if (remembered && remembered.oldDue && remembered.newDue) {
    const oldDue = remembered.oldDue;
    const newDue = toDateOnly_(obj["車検期限"]) || remembered.newDue;
    const added = appendVehicleInspectionHistoryFromRow_(vehicleSheet, row, oldDue, newDue, "更新済", scheduledDate);
    clearRememberedVehicleInspectionDueChange_(vehicleSheet, row);

    if (added) {
      if (plannedCol > 0) vehicleSheet.getRange(row, plannedCol).clearContent();
      if (newDueCol > 0) vehicleSheet.getRange(row, newDueCol).clearContent();
    }

    return { added: added, reason: added ? "" : "同じ履歴が既にある可能性があります" };
  }

  if (scheduledDate) {
    return { added: false, reason: "車検予定日は予定日です。次回の車検期限は『新車検期限』へ入力してから更新済にしてください" };
  }

  return { added: false, reason: "新車検期限が未入力です。次回の車検期限を入力してから更新済にしてください" };
}

function getRememberedVehicleInspectionDueChange_(vehicleSheet, row) {
  const key = getVehicleInspectionChangeKey_(vehicleSheet, row);
  const raw = PropertiesService.getDocumentProperties().getProperty(key);
  if (!raw) return null;

  try {
    const payload = JSON.parse(raw);
    return {
      oldDue: parseDateForVehicleHistory_(payload.oldDue),
      newDue: parseDateForVehicleHistory_(payload.newDue),
      raw: payload
    };
  } catch (e) {
    return null;
  }
}

function clearRememberedVehicleInspectionDueChange_(vehicleSheet, row) {
  const key = getVehicleInspectionChangeKey_(vehicleSheet, row);
  PropertiesService.getDocumentProperties().deleteProperty(key);
}

function sortVehicleInspectionSheetByDue_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("車検管理");
  if (!sheet || sheet.getLastRow() < 3) return;

  const headers = getHeaders_(sheet);
  const dueColIndex = headers.indexOf("車検期限");
  const nameColIndex = headers.indexOf("車両名");
  if (dueColIndex < 0) return;

  try {
    const rowCount = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();

    const dataRows = values.filter(row => isBusinessDataRow_(headers, row, "車検管理"));
    const emptyRows = values.filter(row => !isBusinessDataRow_(headers, row, "車検管理"));

    dataRows.sort((a, b) => {
      const ad = toDateOnly_(a[dueColIndex]);
      const bd = toDateOnly_(b[dueColIndex]);

      // 期限ありを上、期限なしを下へ。
      if (ad && bd) {
        const diff = ad.getTime() - bd.getTime();
        if (diff !== 0) return diff;
      } else if (ad && !bd) {
        return -1;
      } else if (!ad && bd) {
        return 1;
      }

      const an = nameColIndex >= 0 ? String(a[nameColIndex] || "") : "";
      const bn = nameColIndex >= 0 ? String(b[nameColIndex] || "") : "";
      return an.localeCompare(bn, "ja");
    });

    const sorted = dataRows.concat(emptyRows);
    sheet.getRange(2, 1, sorted.length, headers.length).setValues(sorted);
    setCheckboxesForDataRows(sheet);
  } catch (e) {
    console.log("車検管理の並び替えをスキップ: " + e.message);
  }
}

function appendVehicleInspectionHistoryFromRememberedChange_(vehicleSheet, row, note) {
  if (!vehicleSheet || vehicleSheet.getName() !== "車検管理" || row <= 1) return false;

  const key = getVehicleInspectionChangeKey_(vehicleSheet, row);
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(key);
  if (!raw) return false;

  let payload;
  try {
    payload = JSON.parse(raw);
  } catch (e) {
    props.deleteProperty(key);
    return false;
  }

  const oldDue = parseDateForVehicleHistory_(payload.oldDue);
  const rememberedNewDue = parseDateForVehicleHistory_(payload.newDue);

  const headers = getHeaders_(vehicleSheet);
  const rowObj = objectFromRow(headers, vehicleSheet.getRange(row, 1, 1, headers.length).getValues()[0]);
  const currentNewDue = toDateOnly_(rowObj["車検期限"]);

  if (!oldDue || !rememberedNewDue || !currentNewDue) return false;

  // 一時保存後にさらに車検期限が変わった場合は、現在の車検期限を優先する。
  const newDue = currentNewDue;

  const historySheet = ensureVehicleInspectionHistorySheet_();
  const added = appendVehicleInspectionHistoryFromRow_(vehicleSheet, row, oldDue, newDue, note || "更新済");
  if (added || isDuplicateVehicleInspectionHistory_(historySheet, rowObj["車両名"], rowObj["車番"], oldDue, newDue)) {
    props.deleteProperty(key);
  }
  return added;
}

function formatDateForVehicleHistoryKey_(value) {
  const d = toDateOnly_(value);
  if (!d) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

function appendVehicleInspectionHistoryFromRow_(vehicleSheet, row, oldDue, newDue, note, scheduledDate) {
  if (!vehicleSheet || vehicleSheet.getName() !== "車検管理" || row <= 1) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ensureVehicleInspectionHistorySheet_();
  const vehicleHeaders = getHeaders_(vehicleSheet);
  const values = vehicleSheet.getRange(row, 1, 1, vehicleHeaders.length).getValues()[0];
  const obj = objectFromRow(vehicleHeaders, values);

  const vehicleName = obj["車両名"] || "";
  const vehicleNumber = obj["車番"] || "";
  if (!vehicleName && !vehicleNumber) return false;

  if (isDuplicateVehicleInspectionHistory_(historySheet, vehicleName, vehicleNumber, oldDue, newDue)) {
    return false;
  }

  const historyHeaders = getHeaders_(historySheet);
  const historyObj = {
    "更新日": new Date(),
    "車両名": vehicleName,
    "車番": vehicleNumber,
    "旧車検期限": oldDue,
    "新車検期限": newDue,
    "車検予定日": scheduledDate || obj["車検予定日"] || "",
    "担当": getVehicleInspectionStaffForHistory_(obj),
    "備考": note || obj["備考"] || "",
    "履歴ID": buildRecordId_("車検履歴")
  };

  const writeRow = historyHeaders.map(header => historyObj[header] !== undefined ? historyObj[header] : "");
  const nextRow = Math.max(historySheet.getLastRow() + 1, 2);
  historySheet.getRange(nextRow, 1, 1, historyHeaders.length).setValues([writeRow]);

  sortVehicleInspectionHistory_();
  return true;
}

function ensureVehicleInspectionHistorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["車検履歴"];
  let sheet = ss.getSheetByName("車検履歴");
  if (!sheet) sheet = ss.insertSheet("車検履歴");

  const currentHeaders = getHeaders_(sheet);
  const needsHeader = currentHeaders.length < headers.length || headers.some((h, i) => currentHeaders[i] !== h);
  if (needsHeader) {
    ensureSheetHeadersPreserveData_("車検履歴");
  } else {
    applyVehicleInspectionHistoryValidations_(sheet);
    formatSheetBase_(sheet, headers.length);
    createFilterSafely_(sheet, headers.length);
  }

  return sheet;
}

function isDuplicateVehicleInspectionHistory_(historySheet, vehicleName, vehicleNumber, oldDue, newDue) {
  if (!historySheet || historySheet.getLastRow() < 2) return false;

  const headers = getHeaders_(historySheet);
  const nameCol = headers.indexOf("車両名");
  const numberCol = headers.indexOf("車番");
  const oldCol = headers.indexOf("旧車検期限");
  const newCol = headers.indexOf("新車検期限");
  if ([nameCol, numberCol, oldCol, newCol].some(i => i < 0)) return false;

  const values = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, headers.length).getValues();
  const oldTime = toDateOnly_(oldDue).getTime();
  const newTime = toDateOnly_(newDue).getTime();

  return values.some(row => {
    const rowOld = toDateOnly_(row[oldCol]);
    const rowNew = toDateOnly_(row[newCol]);
    return String(row[nameCol] || "") === String(vehicleName || "")
      && String(row[numberCol] || "") === String(vehicleNumber || "")
      && rowOld && rowNew
      && rowOld.getTime() === oldTime
      && rowNew.getTime() === newTime;
  });
}

function sortVehicleInspectionHistory_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("車検履歴");
  if (!sheet || sheet.getLastRow() < 3) return;

  const headers = getHeaders_(sheet);
  const updateCol = headers.indexOf("更新日") + 1;
  if (updateCol <= 0) return;

  try {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length)
      .sort([{ column: updateCol, ascending: false }]);
  } catch (e) {
    console.log("車検履歴の並び替えをスキップ: " + e.message);
  }
}

function parseDateForVehicleHistory_(value) {
  if (!value) return null;

  if (value instanceof Date) return toDateOnly_(value);

  const text = String(value).trim();
  if (!text) return null;

  // Google Sheetsの表示値が yyyy/mm/dd, yyyy-mm-dd, yyyy年m月d日 どれでも読めるようにする。
  const normalized = text
    .replace(/[年月]/g, "/")
    .replace(/日/g, "")
    .replace(/-/g, "/");

  const d = new Date(normalized);
  if (isNaN(d.getTime())) return null;
  return toDateOnly_(d);
}

function getCurrentUserLabel_() {
  // v9.7.48m:
  // 履歴や帳票へGmailアドレスを出さない。
  // 個人情報混入防止のため、ユーザー名が取れない場合は固定ラベルにする。
  return "システム更新";
}

function getVehicleInspectionStaffForHistory_(rowObj) {
  const staff = String((rowObj && rowObj["担当"]) || "").trim();
  if (!staff) return "";
  if (isEmailLike_(staff)) return "";
  if (staff === CLEAR_LABEL) return "";
  return staff;
}

function isEmailLike_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function addVehicleInspectionHistoryForActiveRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (!sheet || sheet.getName() !== "車検管理") {
    toast_("車検管理シートで履歴にしたい行を選択してから実行してください");
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= 1) {
    toast_("車検管理のデータ行を選択してください");
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "選択行の車検履歴を手動追加",
    "旧車検期限を入力してください。例: 2026/06/10",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const oldDue = parseDateForVehicleHistory_(res.getResponseText());
  if (!oldDue) {
    toast_("旧車検期限を日付として読み取れませんでした");
    return;
  }

  const headers = getHeaders_(sheet);
  const obj = objectFromRow(headers, sheet.getRange(row, 1, 1, headers.length).getValues()[0]);
  const newDue = toDateOnly_(obj["車検期限"]);
  if (!newDue) {
    toast_("選択行の車検期限が空欄、または日付ではありません");
    return;
  }

  const added = appendVehicleInspectionHistoryFromRow_(sheet, row, oldDue, newDue, "手動補完");
  toast_(added ? "車検履歴を追加しました" : "同じ車検履歴が既にあるため追加しませんでした");
}

function normalizeVehicleInspectionUpdatedRow_(sheet, row) {
  if (!sheet || sheet.getName() !== "車検管理" || row <= 1) return;

  const headers = getHeaders_(sheet);
  const statusCol = headers.indexOf("車検状態") + 1;
  const dueCol = headers.indexOf("車検期限") + 1;
  const plannedCol = headers.indexOf("車検予定日") + 1;
  const newDueCol = headers.indexOf("新車検期限") + 1;
  const noticeCol = headers.indexOf("通知") + 1;

  if (statusCol <= 0 || dueCol <= 0) return;

  const values = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const obj = objectFromRow(headers, values);
  const status = String(obj["車検状態"] || "").trim();

  if (status !== "更新済") return;

  // 「更新済」は保存する状態ではなく、
  // 車検履歴追加 → 車検期限更新 → 車検状態空欄戻し の実行ボタンとして扱う。
  const result = processVehicleInspectionUpdatedRow_(sheet, row);

  // 履歴追加できた場合だけ更新済を空欄に戻す。
  // 新車検期限未入力などで履歴化できない場合は、利用者が直せるように残す。
  if (result.added) sheet.getRange(row, statusCol).clearContent();

  // 通知を更新
  try {
    if (noticeCol > 0) {
      const newDue = toDateOnly_(sheet.getRange(row, dueCol).getValue());
      sheet.getRange(row, noticeCol).setValue(newDue ? getNoticeText(newDue) : "");
    }
  } catch (e) {}

  // 車検期限順に並べ直し。期限が先のものを上、遠いものを下へ。
  sortVehicleInspectionSheetByDue_();

  if (result.added) {
    toast_("【車検履歴処理】車検履歴へ追加し、車検期限更新＋状態空欄戻しを行いました");
  } else if (result.reason) {
    toast_("【車検履歴処理】履歴は未追加です（" + result.reason + "）");
  } else {
    toast_("【車検履歴処理】履歴は未追加です");
  }
}

function normalizeVehicleInspectionUpdatedRows_(sheet) {
  if (!sheet || sheet.getName() !== "車検管理" || sheet.getLastRow() < 2) return;
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    normalizeVehicleInspectionUpdatedRow_(sheet, row);
  }
}

function ensureV9714MinorLayoutUpdates_() {
  ensureSheetHeadersPreserveData_("行事予定");
  hideEventContentColumnIfExists_();
  ensureV9715FeedbackColumnUpdates_();
}

function repairFeedbackColumnUpdates() {
  ensureV9715FeedbackColumnUpdates_();
  refreshInputSheets();
  toast_("フィードバック反映列を修復しました（資格区分・工事契約日/金額・備品順番・電話番号）");
}

function ensureV9715FeedbackColumnUpdates_() {
  ["工事予定", "電話履歴", "備品修理管理", "資格管理"].forEach(name => {
    ensureSheetHeadersPreserveData_(name);
  });
}

function ensureSheetHeadersPreserveData_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const expectedHeaders = getSheetHeaders_()[sheetName];
  if (!sheet || !expectedHeaders || !expectedHeaders.length) return;

  const currentHeaders = getHeaders_(sheet);
  if (currentHeaders.length && arraysEqual_(currentHeaders.slice(0, expectedHeaders.length), expectedHeaders) && currentHeaders.length === expectedHeaders.length) {
    applyDataValidationByHeaders_(sheet, expectedHeaders);
    formatSheetByName_(sheetName);
    return;
  }

  const lastRow = sheet.getLastRow();
  const oldLastCol = Math.max(sheet.getLastColumn(), currentHeaders.length, expectedHeaders.length);
  const dataRowCount = Math.max(lastRow - 1, 0);
  const objects = [];

  if (currentHeaders.length && dataRowCount > 0) {
    const values = sheet.getRange(2, 1, dataRowCount, Math.max(currentHeaders.length, 1)).getValues();
    values.forEach(rowValues => objects.push(objectFromRow(currentHeaders, rowValues)));
  }

  removeFilter_(sheet);
  showAllColumns_(sheet);
  ensureColumns_(sheet, expectedHeaders.length);

  if (sheet.getMaxColumns() > expectedHeaders.length) {
    sheet.deleteColumns(expectedHeaders.length + 1, sheet.getMaxColumns() - expectedHeaders.length);
  }

  const clearRows = Math.max(sheet.getMaxRows(), 1);
  sheet.getRange(1, 1, clearRows, expectedHeaders.length).clearContent().clearDataValidations();

  sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);

  if (objects.length) {
    const personal = getPersonalMembers_();
    const rows = objects.map(obj => expectedHeaders.map(header => {
      if (header === READ_HEADER || personal.includes(header) || LICENSE_TYPE_HEADERS.includes(header)) return obj[header] === true;
      return getPreservedHeaderValue_(sheetName, header, obj);
    }));
    sheet.getRange(2, 1, rows.length, expectedHeaders.length).setValues(rows);
  }

  applyDataValidationByHeaders_(sheet, expectedHeaders);
  ensureIdsForSheet_(sheet);
  setCheckboxesForDataRows(sheet);
  updateNoticeColumnForSheet_(sheet);
  formatSheetByName_(sheetName);
  createFilterSafely_(sheet, expectedHeaders.length);
}


function getPreservedHeaderValue_(sheetName, header, obj) {
  if (!obj) return "";

  // v9.7.48p1:
  // 工事予定は旧「税区分」列を短い「税」列へ移行する。
  // 以前、契約金額セルに「1,100,000（税込）」のように入れていた場合も、
  // 税だけを分離できる範囲で拾う。
  if (sheetName === "工事予定" && header === "税") {
    return obj["税"] || obj["税区分"] || extractTaxLabelFromAmount_(obj["契約金額"]);
  }

  if (sheetName === "工事予定" && header === "契約金額") {
    return normalizeConstructionAmountValue_(obj["契約金額"]);
  }

  return obj[header] !== undefined ? obj[header] : "";
}

function extractTaxLabelFromAmount_(value) {
  const text = String(value || "").trim();
  if (!text) return "";
  if (/税\s*込|税込み|税込/.test(text)) return "税込";
  if (/税\s*抜|税抜き|税抜/.test(text)) return "税抜";
  if (/不明/.test(text)) return "不明";
  return "";
}

function normalizeConstructionAmountValue_(value) {
  if (value === null || value === undefined) return "";

  // すでに数値ならそのまま保持する。表示形式で円表示する。
  if (typeof value === "number") return value;

  const text = String(value).trim();
  if (!text) return "";

  // 「¥1,200,000（税込）」のような値は税表記だけ外す。
  // 純粋な数値化までは強制しない。未定/概算などを壊さないため。
  return text
    .replace(/[（(]\s*税込み?\s*[）)]/g, "")
    .replace(/[（(]\s*税抜き?\s*[）)]/g, "")
    .replace(/[（(]\s*不明\s*[）)]/g, "")
    .replace(/\s*税込み?\s*$/g, "")
    .replace(/\s*税抜き?\s*$/g, "")
    .replace(/\s*不明\s*$/g, "")
    .trim();
}

function arraysEqual_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (String(a[i] || "") !== String(b[i] || "")) return false;
  }
  return true;
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
  const s = String(status || "").trim();

  // 完了系の行は、過去日でも「期限切れ」ではなく「完了」と表示する。
  // 要確認一覧には「完了」は出さないため、完了済みの期限切れ表示を避けられる。工事予定では施工中も通知欄は完了扱いにする。
  if (["完了", "請求済み", "更新済", "確認済", "返却済", "精算済"].includes(s)) return "完了";

  // 中止・修理不可は完了とは少し意味が違うため、通知欄は空欄にする。
  if (["中止", "修理不可"].includes(s)) return "";

  if (!dateValue) return "";
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


function cleanInputSheetsOrphanIdsLight_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    cleanedIds: 0,
    processedSheets: 0
  };

  getInputSheetNames_().forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const r = cleanOrphanIdsAndCheckboxesForSheet_(sheet, false);
    result.cleanedIds += r.cleanedIds || 0;
    if (r.processed) result.processedSheets++;
  });

  return result;
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
  const prefixMap = { "出先予定":"OUT", "工事予定":"CON", "会議予定":"MTG", "行事予定":"EVT", "作業状況":"WRK", "電話履歴":"TEL", "車検管理":"CAR", "車検履歴":"CARH", "社用車予約":"RES", "日報":"DR", "日報レシート管理":"RCT", "運転免許管理":"LIC", "運転免許明細":"LICD", "運転免許更新履歴":"LICH", "資格管理":"QAL", "資格更新履歴":"QAH", "資格マスタ":"QAM", "備品修理管理":"REP", "お知らせ":"NEWS", "個人ToDo":"TODO", "フィードバック":"FB", "帳簿PDF履歴":"LEDGER" };
  const prefix = prefixMap[sheetName] || "ID";
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss");
  return prefix + "-" + now + "-" + Utilities.getUuid().slice(0, 8);
}

function addSampleMainSchedules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSheetsIfMissing_(["出先予定", "工事予定", "会議予定", "行事予定", "電話履歴", "社用車予約", "お知らせ", "個人ToDo"]);
  const today = new Date();

  writeObjectsToSheet(ss.getSheetByName("出先予定"), [
    {"日付": addDays_(today, 0), "行き先": "サンプル市役所", "用件": "書類提出", "担当": "山田", "社用車": "サンプル作業車A", "状態": "予定", "電話対応": "", "備考": "SAMPLE"},
    {"日付": addDays_(today, 2), "行き先": "サンプル現場A", "用件": "進捗確認", "担当": "鈴木", "社用車": "サンプルトラックA", "状態": "予定", "電話対応": "", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("工事予定"), [
    {"工事名": "サンプル舗装工事", "現場": "サンプル町1丁目", "依頼主": "サンプル建設", "連絡先": "018-000-0000", "契約金額": 1200000, "税": "税込", "開始日": addDays_(today, 1), "終了日": addDays_(today, 10), "状態": "着工前", "担当": "高橋", "電話対応": "未対応", "備考": "SAMPLE"},
    {"工事名": "サンプル排水工事", "現場": "サンプル町2丁目", "依頼主": "サンプル商事", "連絡先": "018-111-1111", "契約金額": 850000, "税": "税抜", "開始日": addDays_(today, -2), "終了日": addDays_(today, 5), "状態": "施工中", "担当": "田中", "電話対応": "対応中", "備考": "SAMPLE"}
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
    {"日付": addDays_(today, 1), "開始時刻": "08:30", "終了時刻": "17:00", "社用車": "サンプル作業車A", "利用者": "山田", "行き先": "サンプル現場", "用途": "資材運搬", "状態": "予定", "備考": "SAMPLE"}
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
    {"所有者": "山田", "普通": true, "準中型": true, "中型": false, "大型": false, "大型特殊": false, "けん引": false, "二種": false, "免許証交付日": addDays_(today, -1000), "免許証有効期限": addDays_(today, 7), "コピー有無": "有", "状態": "更新予定", "備考": "SAMPLE"},
    {"所有者": "鈴木", "普通": true, "準中型": false, "中型": true, "大型": false, "大型特殊": false, "けん引": false, "二種": false, "免許証交付日": addDays_(today, -1200), "免許証有効期限": addDays_(today, 120), "コピー有無": "有", "状態": "有効", "備考": "SAMPLE"}
  ]);

  writeObjectsToSheet(ss.getSheetByName("資格管理"), [
    {"所有者": "高橋", "資格名": "玉掛け技能講習", "区分": "技能講習", "取得区分": "修了", "保有状況": "保有", "取得日": addDays_(today, -600), "更新期限": addDays_(today, 5), "コピー有無": "未確認", "状態": "更新予定", "備考": "SAMPLE"},
    {"所有者": "田中", "資格名": "小型移動式クレーン", "区分": "技能講習", "取得区分": "修了", "保有状況": "保有", "取得日": addDays_(today, -900), "更新期限": addDays_(today, 200), "コピー有無": "有", "状態": "有効", "備考": "SAMPLE"}
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

  // v9.7.48i:
  // GPT診断前に、診断対象の自動出力系ヘッダーだけ軽量修復する。
  // これにより、旧「保有区分」や社員別資格一覧の所有者列なし仕様を、診断表側で誤検知しにくくする。
  let preRepairSummary = repairGptDiagnosticTargetSheets_();

  // v9.7.48p3:
  // GPT診断用シート作成時点でも、資格管理の既読・個人確認チェックを再確認する。
  // p2では修復メニューのトースト上は860個作成できていても、
  // 診断シート側に旧判定が残る/再修復前の状態を拾うケースがあったため、
  // 診断直前にもう一度だけ資格管理のチェックボックスを強制再設定する。
  try {
    const qResult = refreshQualificationManagementCheckboxes_();
    preRepairSummary += " / 資格管理既読チェック " + qResult.checkboxCount + "個";
  } catch (e) {
    preRepairSummary += " / 資格管理既読チェック再確認スキップ:" + e.message;
  }

  try {
    const hideResult = hideSystemColumnsForAllInputSheets_({silent: true});
    preRepairSummary += " / ID列非表示 " + hideResult.hiddenCount + "列";
  } catch (e) {
    preRepairSummary += " / ID列非表示スキップ:" + e.message;
  }

  const rows = [["区分", "対象", "確認内容", "結果", "詳細", "GPTに見てほしい点"]];
  const nowText = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  rows.push(["貼付用", "GPTへの依頼文", "このシートの内容をChatGPTへ貼る", "案内", buildGptDiagnosticPrompt_(), "この行と下の診断表を貼る"]);
  rows.push(["システム", "バージョン", "SYSTEM_VERSION", "情報", SYSTEM_VERSION, "v9.7系の最新版か確認"]);
  rows.push(["システム", "作成日時", "診断作成日時", "情報", nowText, "診断時点の状態"]);
  rows.push(["システム", "担当者", "設定シートから読み込み", "情報", getStaffMembers_().join(" / "), "実名なら社外共有前に匿名化"]);
  rows.push(["システム", "社用車", "設定シートから読み込み", "情報", getCompanyVehicles_().join(" / "), "実車名なら社外共有前にサンプル名へ置換"]);
  rows.push(["システム", "既読確認者", "個人確認列", "情報", getPersonalMembers_().join(" / "), "個人確認列の数と表示幅"]);
  rows.push(["システム", "診断前軽量修復", "資格系ヘッダー・一覧ヘッダー", "情報", preRepairSummary, "GPT診断用シート作成前に、旧列名・旧見出しの誤検知を減らすための軽量修復を実行"]);

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
  appendGptConstructionDiagnostics_(ss, rows);
  appendGptRepairVendorDiagnostics_(ss, rows);
  appendGptVehicleNumberDiagnostics_(ss, rows);
  appendGptVehicleHistoryDiagnostics_(ss, rows);
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
    "下の診断表を見て、シート構成、ヘッダー、一覧反映、要確認一覧、既読チェックボックス、資格管理プルダウン、修理業者マスタ、修理業者の候補外手入力許可、カラーリング/条件付き書式、表示設定、getLastRow/IDだけ空行、車番日付化、車検履歴の担当メールアドレス混入、車検履歴の日付列入力規則、車検履歴の手動再同期仕様、工事予定の契約金額円表示・税プルダウン・連絡先文字列固定、AppSheet設定、メニュー構成、保守用メニューの隔離、車検更新済処理で問題がありそうな点を指摘してください。",
    "実データは社外共有できないため、件数・ヘッダー・診断結果だけで判断してください。",
    "社員別資格一覧は、社員名を見出し行に出すため、明細ヘッダーに所有者列がない仕様です。これをNG扱いしないでください。",
    "工事予定は、契約金額を円表示し、税込/税抜/不明は短い「税」列のプルダウンで管理する仕様です。旧「税区分」列が残っている場合は要確認です。",
    "車検履歴は完全自動相互同期ではなく、打ち間違い時だけ手動メニューで車両名・車番・担当を再同期する仕様です。",
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
    "車検期限変更時に自動追加",
    "OK",
    "車検状態を更新済にした時点で、車検履歴追加・車検期限更新・車検状態空欄戻しを同時に行います。車検予定日は車検に出す日、新車検期限は更新後の次回期限として別扱いします。",
    "更新済と空欄戻しは同じ処理です。車検期限と車検予定日が既に同じ場合は旧期限が分からないため『選択行の車検履歴を手動追加』で補完"
  ]);

  rows.push([
    "車検仕様",
    "車検状態更新時の自動クリア",
    "未来日に更新済なら状態だけ空欄へ戻す",
    "OK",
    "車検状態を更新済にした時点で、同じ行をスキャンし、車検期限=旧期限・新車検期限=次回期限・車検予定日=車検に出す予定日として履歴追加・期限更新・状態空欄戻しを同時に行います。自動で動かない場合は『選択行の車検履歴を処理』で同じ処理を手動実行できます。",
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
    ["会社用：ダッシュボードを作り直す", "repairDashboardSheet", "トップ"],
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
    ["全入力シートを標準表示に整える（少し大きめ）", "standardizeAllInputSheetsReadable", "保守用：表示調整"],
    ["免許・資格・備品だけコンパクト表示", "compactLicenseQualificationEquipmentSheets", "保守用：個別修復"],
    ["運転免許シートのズレ調整", "alignLicenseSheetLayout", "保守用：表示調整"],
    ["車検更新済を処理（履歴＋期限更新）", "clearUpdatedVehicleStatuses", "保守用：表示調整"],
    ["選択行の車検履歴を手動追加", "addVehicleInspectionHistoryForActiveRow", "保守用：電話・免許・その他"],
    ["選択行の車検履歴を車検管理から再同期", "syncSelectedVehicleHistoryRowFromVehicleManagement", "保守用：予定・車検・工事・備品"],
    ["行事予定に終了日列を追加/修復", "ensureV9714MinorLayoutUpdates_", "保守用：表示調整"],
    ["今のシートだけコンパクト表示", "compactActiveSheet", "保守用：表示調整"],
    ["今のシートだけ標準表示に戻す", "standardizeActiveSheetReadable", "保守用：表示調整"],
    ["今のシートの非表示列をすべて表示", "showAllColumnsForActiveSheet", "保守用：表示調整"],
    ["全入力シートのID列・カレンダーID列を非表示", "hideSystemColumnsForAllInputSheets", "保守用：表示調整"],
    ["今のシートだけ表示調整", "formatActiveSheetOnly", "保守用：表示調整"],
    ["今のシートだけID・既読整理", "cleanActiveSheetOrphanIdsAndCheckboxes", "保守用：表示調整"],
    ["今のシートだけ既読チェック再設定", "refreshActiveSheetCheckboxes", "保守用：表示調整"],
    ["表示幅・カラーリングを調整（入力シート・重め）", "formatInputSheetsForLongText", "保守用：表示調整"],
    ["既読表示・個人確認列を再調整（全体・重い）", "applyReadDisplayToAllSheets", "保守用：表示調整"],
    ["個人確認グループを作り直す（全体・重い）", "rebuildCheckGroups", "保守用：表示調整"],
    ["資格管理の既読チェックを軽量再設定", "refreshQualificationManagementCheckboxes", "保守用：資格管理"],
    ["資格診断ヘッダーを軽量修復", "repairQualificationDiagnosticHeaders", "保守用：資格管理"],
    ["全入力シートのID・既読チェックを整理（非常用・重い）", "cleanOrphanIdsAndCheckboxes", "保守用：表示調整"],
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

    const expected = getExpectedHeadersForDiagnostic_(name, headersMap[name]);
    const actual = getHeaders_(sheet);
    const normalizedActual = normalizeHeadersForDiagnostic_(name, actual);
    const missing = expected.filter(h => !normalizedActual.includes(h));
    const extras = getUnexpectedHeadersForDiagnostic_(name, actual, expected);
    const duplicated = findDuplicatedValues_(actual);
    const dataCount = Math.max(sheet.getLastRow() - 1, 0);

    rows.push(["シート構成", name, "存在・件数", "OK", "データ行数: " + dataCount + " / 列数: " + actual.length, "極端に多い場合は動作が重くなる可能性"]);

    const note = getHeaderDiagnosticNote_(name);
    rows.push([
      "ヘッダー",
      name,
      "必須ヘッダー",
      missing.length ? "NG" : "OK",
      missing.length ? "不足: " + missing.join(", ") : "必須ヘッダーOK" + (note ? " / " + note : ""),
      "列名変更後はAppSheet Regenerate Structureが必要"
    ]);

    if (extras.length) {
      rows.push(["ヘッダー", name, "想定外ヘッダー", "確認", extras.join(", "), "旧列名や手入力列が混ざっていないか確認"]);
    }
    if (duplicated.length) {
      rows.push(["ヘッダー", name, "重複ヘッダー", "NG", duplicated.join(", "), "AppSheetや集計で不具合になりやすい"]);
    }
  });
}

function getExpectedHeadersForDiagnostic_(sheetName, defaultExpected) {
  if (sheetName === "社員別資格一覧") {
    // v9.7.48f以降: 社員名は見出し行に出すため、明細ヘッダーに所有者列は置かない。
    return ["区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"];
  }
  if (sheetName === "資格別保有者一覧") {
    return ["資格名", "区分", "所有者", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"];
  }
  return defaultExpected || [];
}

function normalizeHeadersForDiagnostic_(sheetName, headers) {
  // 旧表示名のままでも意味が同じものは、診断では新ヘッダー名として扱う。
  return (headers || []).map(h => {
    const value = String(h || "").trim();
    if ((sheetName === "資格一括登録" || sheetName === "資格重複チェック" || sheetName === "資格別保有者一覧") && value === "保有区分") return "取得区分";
    if (sheetName === "資格別保有者一覧" && value === "資格名/見出し") return "資格名";
    return value;
  });
}

function getUnexpectedHeadersForDiagnostic_(sheetName, actualHeaders, expectedHeaders) {
  const expected = expectedHeaders || [];
  const normalized = normalizeHeadersForDiagnostic_(sheetName, actualHeaders || []);
  const extras = [];
  (actualHeaders || []).forEach((header, i) => {
    const raw = String(header || "").trim();
    if (!raw) return;
    const normalizedHeader = normalized[i];
    if (!expected.includes(normalizedHeader)) extras.push(raw);
  });

  // v9.7.48f以降の社員別資格一覧は、所有者列なしが正。
  // 古い診断表が「所有者不足」と見なさないようにする。
  return uniqueList_(extras);
}

function getHeaderDiagnosticNote_(sheetName) {
  if (sheetName === "社員別資格一覧") return "社員名は見出し行表示のため、所有者列なしが正";
  if (sheetName === "資格一括登録" || sheetName === "資格重複チェック") return "旧『保有区分』は取得区分＋保有状況へ移行対象";
  return "";
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
  const requiredTypes = ["車検", "運転免許", "備品修理", "工事予定", "電話履歴"];
  requiredTypes.forEach(type => {
    const count = typeCounts[type] || 0;
    rows.push(["一覧", type, "一覧反映", count > 0 ? "OK" : "確認", "件数: " + count, count > 0 ? "" : "サンプルがない/更新前/抽出条件違いの可能性"]);
  });

  const qualificationCount = typeCounts["資格"] || 0;
  rows.push(["一覧", "資格管理", "一覧スケジュール通常非表示", qualificationCount === 0 ? "OK" : "要確認", "資格件数: " + qualificationCount, qualificationCount === 0 ? "資格管理は一覧スケジュールには出さず、期限系だけ要確認一覧で確認する仕様" : "日常更新を実行して、一覧スケジュールから資格行を消してください"]);

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

    // v9.7.48p3:
    // 資格管理は、診断する直前にも専用修復を走らせる。
    // これにより、メニュー実行後にGPT診断を作ったのに
    // 「入力済み行チェック不足: 860」が残る誤診断を避ける。
    if (name === "資格管理") {
      try { refreshQualificationManagementCheckboxes_(); } catch (e) { console.log("資格管理チェック診断前再設定をスキップ: " + e.message); }
    }

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
    const advice = name === "資格管理"
      ? "資格管理だけなら『資格管理の既読チェックを軽量再設定』、全体なら『既読表示・個人確認列を再調整』を実行"
      : "空行にチェックが並ぶ場合は『既読表示・個人確認列を再調整』を実行";
    rows.push(["既読", name, "入力済み行だけチェックボックス", result, "空行チェック残り: " + blankRowsWithCheckbox + " / 入力済み行チェック不足: " + dataRowsWithoutCheckbox, advice]);
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
    ["通知", "完了", "淡い緑 #d9ead3", "完了済みの行は期限切れにしない"],
    ["状態", "未対応/未依頼/要確認/中止/失効", "赤〜赤紫系", "対応が必要なもの"],
    ["状態", "未契約", "淡い青灰系", "契約前の工事予定"],
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

function appendGptRepairVendorDiagnostics_(ss, rows) {
  const settings = ss.getSheetByName("設定");
  const equipment = ss.getSheetByName("備品修理管理");
  const vendors = getRepairVendors_();

  rows.push([
    "修理業者",
    "設定シート",
    "修理業者マスタ",
    vendors.length ? "OK" : "確認",
    "登録件数: " + vendors.length + (vendors.length ? " / " + vendors.join(" / ") : ""),
    "備品修理管理で使った業者が 設定種別=修理業者 として追加されているか確認"
  ]);

  if (!equipment) {
    rows.push(["修理業者", "備品修理管理", "修理業者列", "NG", "備品修理管理シートがありません", "シート構成を確認"]);
    return;
  }

  const headers = getHeaders_(equipment);
  const vendorCol = headers.indexOf("修理業者") + 1;
  if (vendorCol <= 0) {
    rows.push(["修理業者", "備品修理管理", "修理業者列", "NG", "修理業者列がありません", "ヘッダーを確認"]);
    return;
  }

  let hasValidation = false;
  let allowInvalid = false;
  try {
    const rule = equipment.getRange(2, vendorCol).getDataValidation();
    hasValidation = !!rule;
    allowInvalid = !!(rule && typeof rule.getAllowInvalid === "function" && rule.getAllowInvalid());
  } catch (e) {}

  rows.push([
    "修理業者",
    "備品修理管理",
    "プルダウン候補外入力",
    hasValidation && allowInvalid ? "OK" : "NG",
    "入力規則: " + (hasValidation ? "あり" : "なし") + " / 候補外入力許可: " + (allowInvalid ? "はい" : "いいえ"),
    "修理業者は手入力後にマスタ自動追加するため、候補外入力許可が必要"
  ]);

  const used = [];
  if (equipment.getLastRow() >= 2) {
    equipment.getRange(2, vendorCol, equipment.getLastRow() - 1, 1).getValues().forEach(row => {
      const value = String(row[0] || "").trim();
      if (isValidRepairVendorName_(value)) used.push(value);
    });
  }

  const masterKeys = {};
  vendors.forEach(v => masterKeys[normalizeTextForKey_(v)] = true);
  const missing = uniqueList_(used).filter(v => !masterKeys[normalizeTextForKey_(v)]);

  rows.push([
    "修理業者",
    "備品修理管理",
    "使用済み業者のマスタ反映",
    missing.length ? "確認" : "OK",
    missing.length ? "未登録: " + missing.join(" / ") : "使用済み修理業者はマスタ反映済み",
    missing.length ? "『修理業者マスタ＋プルダウンを修復』を実行" : ""
  ]);
}

function appendGptConstructionDiagnostics_(ss, rows) {
  const sheet = ss.getSheetByName("工事予定");
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  const amountCol = headers.indexOf("契約金額") + 1;
  const taxCol = headers.indexOf("税") + 1;
  const oldTaxCol = headers.indexOf("税区分") + 1;
  const statusCol = headers.indexOf("状態") + 1;

  let statusHasBilling = false;
  try {
    if (statusCol > 0) {
      const rule = sheet.getRange(2, statusCol).getDataValidation();
      const values = rule && rule.getCriteriaValues ? rule.getCriteriaValues() : [];
      const list = values && values[0] ? values[0] : [];
      statusHasBilling = Array.isArray(list) && list.indexOf("請求済み") >= 0;
    }
  } catch (e) {}

  rows.push([
    "工事予定",
    "契約金額・税",
    "契約金額は円表示、税は税込/税抜/不明のプルダウン",
    amountCol > 0 && taxCol > 0 && oldTaxCol <= 0 ? "OK" : "要確認",
    "契約金額列: " + (amountCol > 0 ? "あり" : "なし") + " / 税列: " + (taxCol > 0 ? "あり" : "なし") + " / 旧税区分列: " + (oldTaxCol > 0 ? "あり" : "なし"),
    amountCol > 0 && taxCol > 0 && oldTaxCol <= 0 ? "" : "工事予定の契約金額・税・状態を修復を実行"
  ]);

  rows.push([
    "工事予定",
    "状態プルダウン",
    "請求済みがあるか",
    statusHasBilling ? "OK" : "要確認",
    statusHasBilling ? "請求済みあり" : "請求済み未確認",
    statusHasBilling ? "" : "工事予定の未契約・通知完了・色を修復、または契約金額・状態修復を実行"
  ]);
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


function appendGptVehicleHistoryDiagnostics_(ss, rows) {
  const sheet = ss.getSheetByName("車検履歴");
  if (!sheet) {
    rows.push(["車検履歴", "車検履歴", "シート存在", "NG", "車検履歴シートがありません", "車検履歴列修復を実行"]);
    return;
  }

  const headers = getHeaders_(sheet);
  const dateTargets = ["更新日", "旧車検期限", "新車検期限", "車検予定日"];
  const invalidDateDropdowns = [];
  dateTargets.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) {
      invalidDateDropdowns.push(header + ": 列なし");
      return;
    }
    try {
      const rule = sheet.getRange(2, col).getDataValidation();
      const criteria = rule ? String(rule.getCriteriaType()) : "なし";
      if (rule && criteria.indexOf("DATE") < 0) invalidDateDropdowns.push(header + ": " + criteria);
    } catch (e) {}
  });

  rows.push([
    "車検履歴",
    "日付列入力規則",
    "更新日・旧車検期限・新車検期限・車検予定日",
    invalidDateDropdowns.length ? "NG" : "OK",
    invalidDateDropdowns.length ? invalidDateDropdowns.join(" / ") : "日付列として設定済み",
    invalidDateDropdowns.length ? "車検履歴の日付・担当入力規則を修復を実行" : ""
  ]);

  const staffCol = headers.indexOf("担当") + 1;
  const emailRows = [];
  if (staffCol > 0 && sheet.getLastRow() >= 2) {
    sheet.getRange(2, staffCol, sheet.getLastRow() - 1, 1).getValues().forEach((row, i) => {
      const value = String(row[0] || "").trim();
      if (isEmailLike_(value)) emailRows.push("行" + (i + 2));
    });
  }

  rows.push([
    "車検履歴",
    "担当メールアドレス混入",
    "担当列にGmail等が入っていないか",
    emailRows.length ? "NG" : "OK",
    emailRows.length ? emailRows.join(" / ") : "メールアドレスなし",
    emailRows.length ? "担当は車検管理の担当名を転記し、メールアドレスは空欄化" : ""
  ]);
  rows.push([
    "車検履歴",
    "手動再同期仕様",
    "車検管理の打ち間違い修正を履歴へ反映する方法",
    "案内",
    "過去履歴は自動相互同期しません。必要時だけ選択行の車検履歴を車検管理から再同期します。",
    "車両名・車番・担当のみ再同期し、旧/新車検期限や車検予定日は履歴保護のため変更しません"
  ]);


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
  rows.push(["AppSheet", "数値列", "Number/Price型候補", "情報", "工事予定: 契約金額、日報レシート管理: 金額", "工事予定の税はEnum/Text型、契約金額はNumber/Price型推奨"]);
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

  const rows = buildDashboardRows_();
  resetDashboardSheetForRewrite_(sheet, rows.length, rows[0].length);
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  formatDashboardSheet_(sheet, rows.length, rows[0].length);
}

function repairDashboardSheet() {
  try {
    createDashboard();
    toast_("ダッシュボードを作り直しました");
  } catch (err) {
    toast_("ダッシュボード作り直しでエラー: " + err.message);
    throw err;
  }
}

function buildDashboardRows_() {
  const scheduleCount = getSheetDataCount_("一覧スケジュール");
  const alertCount = getSheetDataCount_("要確認一覧");
  const unreadAssigneeCount = getSheetDataCount_("担当別未読");
  const readRateCount = getSheetDataCount_("既読率集計");

  return [
    ["項目", "件数/状態", "備考"],
    ["システム版", SYSTEM_VERSION, ""],
    ["一覧スケジュール件数", scheduleCount, "資格管理は通常一覧から除外"],
    ["今日の予定", countNotice_("今日"), "一覧スケジュール"],
    ["期限切れ", countNotice_("期限切れ"), "一覧スケジュール"],
    ["3日以内", countNotice_("3日以内"), "一覧スケジュール"],
    ["7日以内", countNotice_("7日以内"), "一覧スケジュール"],
    ["要確認件数", alertCount, "要確認一覧"],
    ["未確認レシート", countReceiptAlerts_(), "日報レシート管理"],
    ["担当別未読", unreadAssigneeCount, "担当別未読シート"],
    ["既読率集計", readRateCount, "既読率集計シート"],
    ["更新必要フラグ", isSummaryDirty_() ? "あり" : "なし", "onEdit後は日常更新推奨"],
    ["最終更新", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm"), "ダッシュボード作成時刻"]
  ];
}

function resetDashboardSheetForRewrite_(sheet, rowCount, colCount) {
  const keepRows = Math.max(rowCount + 2, 20);
  const keepCols = Math.max(colCount, 3);

  try { if (sheet.getFilter()) sheet.getFilter().remove(); } catch (e) {}
  try { sheet.setConditionalFormatRules([]); } catch (e) {}
  try { sheet.getDataRange().breakApart(); } catch (e) {}
  try { sheet.showRows(1, sheet.getMaxRows()); } catch (e) {}
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (e) {}
  try { clearAllRowGroupsAndShowRows_(sheet); } catch (e) {}

  if (sheet.getMaxRows() < keepRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), keepRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < keepCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), keepCols - sheet.getMaxColumns());
  }

  try {
    if (sheet.getMaxRows() > keepRows) {
      sheet.deleteRows(keepRows + 1, sheet.getMaxRows() - keepRows);
    }
  } catch (e) {}
  try {
    if (sheet.getMaxColumns() > keepCols) {
      sheet.deleteColumns(keepCols + 1, sheet.getMaxColumns() - keepCols);
    }
  } catch (e) {}

  sheet.clear();
  sheet.clearFormats();
  try { sheet.clearNotes(); } catch (e) {}
  try { sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations(); } catch (e) {}
}

function formatDashboardSheet_(sheet, rowCount, colCount) {
  try { sheet.setFrozenRows(1); } catch (e) {}
  try { sheet.setFrozenColumns(0); } catch (e) {}

  sheet.getRange(1, 1, 1, colCount)
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  if (rowCount > 1) {
    sheet.getRange(2, 1, rowCount - 1, colCount)
      .setVerticalAlignment("middle")
      .setWrap(true);
    sheet.getRange(2, 2, rowCount - 1, 1).setHorizontalAlignment("center");
  }

  try {
    sheet.getRange(1, 1, rowCount, colCount).setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);
  } catch (e) {}

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 330);
  try { sheet.setRowHeights(1, Math.max(rowCount, 1), 32); } catch (e) {}
  try { sheet.setRowHeight(1, 36); } catch (e) {}
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


function standardizeAllInputSheetsReadable() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    "全入力シートを標準表示に整える（少し大きめ）",
    "全入力シートを読みやすさ優先の少し大きめ表示に整えます。ID列・カレンダーID列などは非表示にしますが、データや列は削除しません。続行しますか？",
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
    applyDisplayLayoutForSheet_(sheet, "standard");
    count++;
  });

  hideSystemColumnsForAllInputSheets_({silent: true});
  toast_("全入力シートを標準表示に整えました（" + count + "シート・ID列非表示）" + (skipped.length ? " / 未作成: " + skipped.join(", ") : ""));
}

// 旧メニュー名・旧診断との互換用。全体調整は「少し大きめ標準表示」に寄せる。
function compactAllInputSheets() {
  standardizeAllInputSheetsReadable();
}

function standardizeActiveSheetReadable() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || isAutoOutputSheet_(sheet.getName()) || SUPPORT_SHEETS.includes(sheet.getName())) {
    toast_("入力シートを開いてから実行してください");
    return;
  }

  ensureV9714MinorLayoutUpdates_();
  applyDisplayLayoutForSheet_(sheet, "standard");
  hideSystemColumnsForSheet_(sheet);
  toast_("今のシートを標準表示に戻しました（ID列非表示）: " + sheet.getName());
}

function compactActiveSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet || isAutoOutputSheet_(sheet.getName()) || SUPPORT_SHEETS.includes(sheet.getName())) {
    toast_("入力シートを開いてから実行してください");
    return;
  }

  ensureV9714MinorLayoutUpdates_();
  applyDisplayLayoutForSheet_(sheet, "compact");
  hideSystemColumnsForSheet_(sheet);
  toast_("今のシートだけコンパクト表示にしました（ID列非表示）: " + sheet.getName());
}

function applyCompactLayoutForSheet_(sheet) {
  applyDisplayLayoutForSheet_(sheet, "compact");
}

function applyDisplayLayoutForSheet_(sheet, mode) {
  const name = sheet.getName();
  const layoutMode = mode === "compact" ? "compact" : "standard";

  // 資格管理は資格IDが長く、折り返しで行高が肥大化しやすい。
  // 免許・備品の汎用表示調整とは分けて、専用の軽量コンパクト表示を使う。
  if (name === "資格管理") {
    applyQualificationManagementCompactDisplay_(sheet);
    return;
  }

  if (["運転免許管理", "備品修理管理"].includes(name)) {
    applyLicenseEquipmentDisplayLayout_(sheet, layoutMode);
    return;
  }

  applyGeneralDisplayLayout_(sheet, layoutMode);
}

function applyCompactGeneralSheetLayout_(sheet) {
  applyGeneralDisplayLayout_(sheet, "compact");
}

function applyGeneralDisplayLayout_(sheet, mode) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const sheetName = sheet.getName();
  const headers = getHeaders_(sheet);
  const minDisplayRows = sheetName === "運転免許管理" ? DRIVER_LICENSE_DISPLAY_MIN_ROWS : 2;
  const lastRow = Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow(), minDisplayRows));
  const lastCol = headers.length;
  const personal = getPersonalMembers_();
  const layoutMode = mode === "compact" ? "compact" : "standard";

  sheet.setFrozenColumns(0);
  showAllColumns_(sheet);

  headers.forEach((header, index) => {
    const col = index + 1;
    if (!header) return;

    const width = getDisplayColumnWidth_(sheetName, header, personal, layoutMode);
    sheet.setColumnWidth(col, width);

    const shouldVertical = header === READ_HEADER || personal.includes(header);
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
        .setVerticalAlignment("middle")
        .setWrap(true);
    }

    if (DATE_HEADERS.includes(header) || TIME_HEADERS.includes(header) || [
      "通知", "状態", "車検状態", "電話対応", "経理確認", "精算状態", "重要度", "対応状況", "区分", "コピー有無", "返却済"
    ].includes(header) || shouldVertical) {
      sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    } else {
      sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1)
        .setVerticalAlignment("middle");
    }
  });

  const idHeader = SHEET_ID_HEADERS[sheetName];
  if (idHeader) hideColumnByHeaderSafely_(sheet, idHeader);
  hideColumnByHeaderSafely_(sheet, CALENDAR_ID_HEADER);
  ["写真", "レシート写真", "PDFリンク"].forEach(header => hideColumnByHeaderSafely_(sheet, header));

  rebuildPersonalCheckGroupForSheet_(sheet);

  const maxRows = Math.min(sheet.getMaxRows(), Math.max(lastRow, 20));
  sheet.setRowHeight(1, layoutMode === "compact" ? 52 : 60);
  sheet.setRowHeights(2, Math.max(maxRows - 1, 1), layoutMode === "compact" ? 31 : 36);
  sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  applyColorRules(sheet);
  createFilterSafely_(sheet, lastCol);
}

function getCompactColumnWidth_(sheetName, header, personal) {
  return getDisplayColumnWidth_(sheetName, header, personal || getPersonalMembers_(), "compact");
}

function getDisplayColumnWidth_(sheetName, header, personal, mode) {
  const compact = mode === "compact";

  if (personal.includes(header)) return compact ? 36 : 42;
  if (header === READ_HEADER) return compact ? 36 : 42;

  if (DATE_HEADERS.includes(header)) return compact ? 102 : 118;
  if (TIME_HEADERS.includes(header)) return compact ? 68 : 78;

  if (header === "通知") return compact ? 94 : 112;
  if (["状態", "車検状態", "電話対応", "対応状況"].includes(header)) return compact ? 88 : 106;
  if (["経理確認", "精算状態"].includes(header)) return compact ? 88 : 106;
  if (["重要度", "区分", "コピー有無", "返却済"].includes(header)) return compact ? 74 : 90;

  if (["担当", "入力者", "利用者", "投稿者", "記入者", "所有者", "修理依頼者"].includes(header)) return compact ? 92 : 115;
  if (["社用車", "車両名"].includes(header)) return compact ? 110 : 145;
  if (header === "車番") return compact ? 108 : 140;

  if (["行き先", "現場", "相手", "支払先", "備品名", "資格名", "タイトル", "依頼主"].includes(header)) return compact ? 130 : 165;
  if (["工事名", "会議名", "行事名"].includes(header)) return compact ? 150 : 190;

  if (["用件", "内容", "作業内容", "日報文章", "問題点", "明日の予定", "他現場状況", "対応メモ", "気になった内容", "対応方針"].includes(header)) return compact ? 180 : 235;

  if (["連絡先", "契約金額", "支払方法", "金額"].includes(header)) return compact ? 96 : 120;
  if (["備考", "メモ"].includes(header)) return compact ? 155 : 205;

  if (header === CALENDAR_ID_HEADER || String(header).endsWith("ID") || header === "ToDo_ID") return 40;
  if (["写真", "レシート写真", "PDFリンク"].includes(header)) return 40;

  return compact ? 92 : 112;
}


function getDriverLicenseIssueDate_(row) {
  if (!row) return "";
  return row[DRIVER_LICENSE_ISSUE_DATE_HEADER] || row["取得日"] || row["交付日"] || "";
}

function getDriverLicenseExpiryDate_(row) {
  if (!row) return "";
  return row[DRIVER_LICENSE_EXPIRY_DATE_HEADER] || row["更新期限"] || row["有効期限"] || row["免許証更新期限"] || "";
}


function isUpdateHistorySheet_(sheetName) {
  return sheetName === DRIVER_LICENSE_UPDATE_HISTORY_SHEET_NAME || sheetName === QUALIFICATION_UPDATE_HISTORY_SHEET_NAME;
}

function createDriverLicenseUpdateHistorySheet() {
  const sheet = ensureDriverLicenseUpdateHistorySheet_();
  sortDriverLicenseUpdateHistorySheet_();
  toast_("運転免許更新履歴シートを作成/修復しました。普段は免許証有効期限・コピー有無・状態の変更だけ記録します。");
  return sheet;
}

function createQualificationUpdateHistorySheet() {
  const sheet = ensureQualificationUpdateHistorySheet_();
  sortQualificationUpdateHistorySheet_();
  toast_("資格更新履歴シートを作成/修復しました。普段は更新期限・コピー有無・状態の変更だけ記録します。");
  return sheet;
}

function ensureDriverLicenseUpdateHistorySheet_() {
  return ensureUpdateHistorySheet_(DRIVER_LICENSE_UPDATE_HISTORY_SHEET_NAME);
}

function ensureQualificationUpdateHistorySheet_() {
  return ensureUpdateHistorySheet_(QUALIFICATION_UPDATE_HISTORY_SHEET_NAME);
}

function ensureUpdateHistorySheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()[sheetName];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const currentHeaders = getHeaders_(sheet);
  const needsHeader = currentHeaders.length < headers.length || headers.some((h, i) => currentHeaders[i] !== h);
  if (needsHeader) {
    ensureSheetHeadersPreserveData_(sheetName);
  } else {
    formatSheetBase_(sheet, headers.length);
    createFilterSafely_(sheet, headers.length);
  }

  try {
    const recordCol = headers.indexOf("記録日時") + 1;
    if (recordCol > 0 && sheet.getMaxRows() > 1) {
      sheet.getRange(2, recordCol, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy/mm/dd hh:mm");
    }
    ["旧有効期限", "新有効期限", "免許証交付日", "旧更新期限", "新更新期限"].forEach(header => {
      const col = headers.indexOf(header) + 1;
      if (col > 0 && sheet.getMaxRows() > 1) {
        sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("yyyy/mm/dd");
      }
    });
    hideSystemColumnsForSheet_(sheet);
  } catch (e) {}

  return sheet;
}

function handleDriverLicenseUpdateHistoryEdit_(e) {
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "運転免許管理") return false;
  if (e.range.getRow() <= 1 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const header = headers[e.range.getColumn() - 1] || "";
  const targetHeaders = [DRIVER_LICENSE_EXPIRY_DATE_HEADER, "更新期限", "コピー有無", "状態"];
  if (!targetHeaders.includes(header)) return false;

  const oldValue = e.oldValue !== undefined ? String(e.oldValue || "").trim() : "";
  const newValue = String(e.range.getDisplayValue() || "").trim();
  if (oldValue === newValue) return false;

  const row = e.range.getRow();
  const valueRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const displayRow = sheet.getRange(row, 1, 1, headers.length).getDisplayValues()[0];
  const obj = objectFromRow(headers, valueRow);
  const disp = objectFromRow(headers, displayRow);
  const owner = String(disp["所有者"] || obj["所有者"] || "").trim();
  if (!owner) return false;

  const expiryHeader = headers.includes(DRIVER_LICENSE_EXPIRY_DATE_HEADER) ? DRIVER_LICENSE_EXPIRY_DATE_HEADER : "更新期限";
  const currentExpiry = disp[expiryHeader] || "";
  const oldExpiry = (header === DRIVER_LICENSE_EXPIRY_DATE_HEADER || header === "更新期限") ? oldValue : currentExpiry;
  const newExpiry = (header === DRIVER_LICENSE_EXPIRY_DATE_HEADER || header === "更新期限") ? newValue : currentExpiry;

  appendDriverLicenseUpdateHistory_({
    "記録日時": new Date(),
    "所有者": owner,
    "変更項目": header === "更新期限" ? DRIVER_LICENSE_EXPIRY_DATE_HEADER : header,
    "旧値": oldValue,
    "新値": newValue,
    "旧有効期限": oldExpiry,
    "新有効期限": newExpiry,
    "免許証交付日": disp[DRIVER_LICENSE_ISSUE_DATE_HEADER] || disp["取得日"] || "",
    "コピー有無": disp["コピー有無"] || "",
    "状態": disp["状態"] || "",
    "備考": disp["備考"] || "",
    "元免許ID": disp["免許ID"] || obj["免許ID"] || ""
  });
  return true;
}

function handleQualificationUpdateHistoryEdit_(e) {
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "資格管理") return false;
  if (e.range.getRow() <= 1 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const header = headers[e.range.getColumn() - 1] || "";
  const targetHeaders = ["更新期限", "コピー有無", "状態"];
  if (!targetHeaders.includes(header)) return false;

  const oldValue = e.oldValue !== undefined ? String(e.oldValue || "").trim() : "";
  const newValue = String(e.range.getDisplayValue() || "").trim();
  if (oldValue === newValue) return false;

  const row = e.range.getRow();
  const valueRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  const displayRow = sheet.getRange(row, 1, 1, headers.length).getDisplayValues()[0];
  const obj = objectFromRow(headers, valueRow);
  const disp = objectFromRow(headers, displayRow);
  const owner = String(disp["所有者"] || obj["所有者"] || "").trim();
  const qualificationName = String(disp["資格名"] || obj["資格名"] || "").trim();
  if (!owner || !qualificationName) return false;

  const currentDue = disp["更新期限"] || "";
  const oldDue = header === "更新期限" ? oldValue : currentDue;
  const newDue = header === "更新期限" ? newValue : currentDue;

  appendQualificationUpdateHistory_({
    "記録日時": new Date(),
    "所有者": owner,
    "区分": disp["区分"] || obj["区分"] || "",
    "資格名": qualificationName,
    "変更項目": header,
    "旧値": oldValue,
    "新値": newValue,
    "旧更新期限": oldDue,
    "新更新期限": newDue,
    "コピー有無": disp["コピー有無"] || "",
    "状態": disp["状態"] || "",
    "備考": disp["備考"] || "",
    "元資格ID": disp["資格ID"] || obj["資格ID"] || ""
  });
  return true;
}

function appendDriverLicenseUpdateHistory_(historyObj) {
  const sheet = ensureDriverLicenseUpdateHistorySheet_();
  historyObj["免許更新履歴ID"] = historyObj["免許更新履歴ID"] || buildRecordId_(DRIVER_LICENSE_UPDATE_HISTORY_SHEET_NAME);
  appendUpdateHistoryObject_(sheet, historyObj);
}

function appendQualificationUpdateHistory_(historyObj) {
  const sheet = ensureQualificationUpdateHistorySheet_();
  historyObj["資格更新履歴ID"] = historyObj["資格更新履歴ID"] || buildRecordId_(QUALIFICATION_UPDATE_HISTORY_SHEET_NAME);
  appendUpdateHistoryObject_(sheet, historyObj);
}

function appendUpdateHistoryObject_(sheet, obj) {
  if (!sheet || !obj) return;
  const headers = getHeaders_(sheet);
  const nextRow = Math.max(sheet.getLastRow() + 1, 2);
  const row = headers.map(header => obj[header] !== undefined ? obj[header] : "");
  if (sheet.getMaxRows() < nextRow) sheet.insertRowsAfter(sheet.getMaxRows(), nextRow - sheet.getMaxRows());
  sheet.getRange(nextRow, 1, 1, headers.length).setValues([row]);
  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
}

function sortDriverLicenseUpdateHistorySheet() {
  sortDriverLicenseUpdateHistorySheet_();
  toast_("運転免許更新履歴を新しい順に整理しました");
}

function sortQualificationUpdateHistorySheet() {
  sortQualificationUpdateHistorySheet_();
  toast_("資格更新履歴を新しい順に整理しました");
}

function sortDriverLicenseUpdateHistorySheet_() {
  sortUpdateHistorySheetByRecordedAt_(DRIVER_LICENSE_UPDATE_HISTORY_SHEET_NAME);
}

function sortQualificationUpdateHistorySheet_() {
  sortUpdateHistorySheetByRecordedAt_(QUALIFICATION_UPDATE_HISTORY_SHEET_NAME);
}

function sortUpdateHistorySheetByRecordedAt_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 3) return;
  const headers = getHeaders_(sheet);
  const dateCol = headers.indexOf("記録日時");
  if (dateCol < 0) return;

  const count = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, count, headers.length).getValues();
  const dataRows = values.filter(row => row.some(v => v !== "" && v !== null));
  const emptyRows = values.filter(row => !row.some(v => v !== "" && v !== null));

  dataRows.sort((a, b) => {
    const ad = toDateOnly_(a[dateCol]) || (a[dateCol] instanceof Date ? a[dateCol] : null);
    const bd = toDateOnly_(b[dateCol]) || (b[dateCol] instanceof Date ? b[dateCol] : null);
    const at = ad ? ad.getTime() : 0;
    const bt = bd ? bd.getTime() : 0;
    return bt - at;
  });

  const sorted = dataRows.concat(emptyRows);
  sheet.getRange(2, 1, sorted.length, headers.length).setValues(sorted);
  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
}

function repairLicenseAndQualificationUpdateHistorySheets() {
  ensureDriverLicenseUpdateHistorySheet_();
  ensureQualificationUpdateHistorySheet_();
  toast_("運転免許更新履歴・資格更新履歴を作成/修復しました");
}

function repairDriverLicenseSheetsLight() {
  const lock = acquireCompanyPackLock_("運転免許ヘッダー・明細修復");
  if (!lock) return;

  try {
    const licenseSheet = repairDriverLicenseManagementHeadersLight_();
    const detailSheet = ensureDriverLicenseDetailSheet_();
    applyDriverLicenseManagementLightFormat_();
    if (licenseSheet) hideSystemColumnsForSheet_(licenseSheet);
    if (detailSheet) hideSystemColumnsForSheet_(detailSheet);
    toast_("運転免許管理の見出し・プルダウン・明細シートを軽量修復しました");
  } catch (err) {
    toast_("運転免許ヘッダー・明細修復でエラー: " + err.message);
    throw err;
  } finally {
    releaseCompanyPackLock_(lock);
  }
}

function repairDriverLicenseManagementHeadersLight_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "運転免許管理";
  let sheet = ss.getSheetByName(sheetName);
  const desired = getSheetHeaders_()[sheetName];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheet_(sheet, desired);
    return sheet;
  }

  ensureColumns_(sheet, desired.length);
  const width = Math.max(sheet.getLastColumn(), desired.length);
  const current = sheet.getRange(1, 1, 1, width).getValues()[0].map(v => String(v || "").trim());
  const next = current.slice(0, width);

  desired.forEach((header, index) => {
    const now = String(next[index] || "").trim();
    if (!now) {
      next[index] = header;
      return;
    }
    if (header === DRIVER_LICENSE_ISSUE_DATE_HEADER && ["取得日", "交付日", "免許交付日"].includes(now)) {
      next[index] = DRIVER_LICENSE_ISSUE_DATE_HEADER;
      return;
    }
    if (header === DRIVER_LICENSE_EXPIRY_DATE_HEADER && ["更新期限", "有効期限", "免許証更新期限", "免許更新期限"].includes(now)) {
      next[index] = DRIVER_LICENSE_EXPIRY_DATE_HEADER;
      return;
    }
  });

  // 既存の旧見出しが別位置に残っている場合も、標準位置の見出しを優先して復旧する。
  desired.forEach((header, index) => {
    if (!String(next[index] || "").trim()) next[index] = header;
  });

  sheet.getRange(1, 1, 1, desired.length).setValues([next.slice(0, desired.length)]);
  applyDataValidationByHeaders_(sheet, desired);
  updateNoticeColumnForSheet_(sheet);
  ensureIdsForSheet_(sheet);
  setCheckboxesForDataRows(sheet);
  hideSystemColumnsForSheet_(sheet);
  return sheet;
}

function createDriverLicenseDetailSheet() {
  ensureDriverLicenseDetailSheet_();
  toast_("運転免許明細シートを作成/修復しました");
}

function ensureDriverLicenseDetailSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = DRIVER_LICENSE_DETAIL_SHEET_NAME;
  const headers = getSheetHeaders_()[sheetName];
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    setupSheet_(sheet, headers);
  } else {
    ensureColumns_(sheet, headers.length);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    applyDataValidationByHeaders_(sheet, headers);
    ensureIdsForSheet_(sheet);
    hideSystemColumnsForSheet_(sheet);
  }

  try {
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 115);
    sheet.setColumnWidth(2, 105);
    sheet.setColumnWidth(3, 105);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 140);
    sheet.setColumnWidth(6, 220);
    const displayRows = Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow(), DRIVER_LICENSE_DETAIL_DISPLAY_MIN_ROWS));
    sheet.setRowHeight(1, 52);
    if (displayRows >= 2) sheet.setRowHeights(2, displayRows - 1, 34);
    sheet.getRange(1, 1, displayRows, headers.length).setVerticalAlignment("middle").setWrap(true);
    createFilterSafely_(sheet, headers.length);
  } catch (e) {}
  return sheet;
}

function applyDriverLicenseManagementLightFormat_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("運転免許管理");
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  const lastCol = headers.length;
  const displayRows = Math.min(
    sheet.getMaxRows(),
    Math.max(sheet.getLastRow(), DRIVER_LICENSE_DISPLAY_MIN_ROWS)
  );
  const personal = getPersonalMembers_();

  try {
    sheet.setFrozenRows(1);
    const widthMap = {
      "所有者": 115,
      "普通": 56,
      "準中型": 64,
      "中型": 56,
      "大型": 56,
      "大型特殊": 72,
      "けん引": 60,
      "二種": 56,
      [DRIVER_LICENSE_ISSUE_DATE_HEADER]: 118,
      [DRIVER_LICENSE_EXPIRY_DATE_HEADER]: 118,
      "コピー有無": 82,
      "状態": 82,
      "通知": 92,
      READ_HEADER: 38,
      "備考": 180
    };

    headers.forEach((header, index) => {
      const col = index + 1;
      const w = widthMap[header] || (personal.includes(header) ? 38 : 96);
      if (w) sheet.setColumnWidth(col, w);

      const range = sheet.getRange(1, col, displayRows, 1);
      const shouldVertical = LICENSE_TYPE_HEADERS.includes(header) || header === READ_HEADER || personal.includes(header);
      if (shouldVertical) {
        range
          .setVerticalText(true)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setWrap(true);
      } else {
        range
          .setVerticalText(false)
          .setVerticalAlignment("middle")
          .setWrap(true);
        if (["所有者", DRIVER_LICENSE_ISSUE_DATE_HEADER, DRIVER_LICENSE_EXPIRY_DATE_HEADER, "コピー有無", "状態", "通知"].includes(header)) {
          sheet.getRange(2, col, Math.max(displayRows - 1, 1), 1)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        }
      }
    });

    sheet.setRowHeight(1, 80);
    if (displayRows >= 2) {
      sheet.setRowHeights(2, displayRows - 1, 34);
    }

    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    createFilterSafely_(sheet, headers.length);
    hideSystemColumnsForSheet_(sheet);
  } catch (e) {}
}

function getStaffSortOrderMap_() {
  const map = {};
  getStaffMembers_().forEach((name, index) => {
    const key = String(name || "").trim();
    if (key && map[key] === undefined) map[key] = index + 1;
  });
  return map;
}

function compareJapaneseText_(a, b) {
  const aa = String(a || "").trim();
  const bb = String(b || "").trim();
  if (!aa && !bb) return 0;
  if (!aa) return 1;
  if (!bb) return -1;
  try {
    return aa.localeCompare(bb, "ja", {numeric: true, sensitivity: "base"});
  } catch (e) {
    return aa > bb ? 1 : aa < bb ? -1 : 0;
  }
}

function compareStaffNameBySettingsOrder_(a, b) {
  const order = getStaffSortOrderMap_();
  const aa = String(a || "").trim();
  const bb = String(b || "").trim();
  const oa = order[aa] || 999999;
  const ob = order[bb] || 999999;
  if (oa !== ob) return oa - ob;
  return compareJapaneseText_(aa, bb);
}

function sortDriverLicenseManagementByOwnerKana() {
  repairDriverLicenseManagementHeadersLight_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("運転免許管理");
  if (!sheet || sheet.getLastRow() < 3) {
    toast_("運転免許管理は並び替えるデータがありません");
    return;
  }

  const headers = getHeaders_(sheet);
  const rowCount = sheet.getLastRow() - 1;
  const colCount = headers.length;
  const range = sheet.getRange(2, 1, rowCount, colCount);
  const rows = range.getValues();

  const indexed = rows.map((values, index) => ({values, index, obj: objectFromRow(headers, values)}));
  indexed.sort((a, b) => {
    const aHas = isBusinessDataRow_(headers, a.values, "運転免許管理");
    const bHas = isBusinessDataRow_(headers, b.values, "運転免許管理");
    if (aHas !== bHas) return aHas ? -1 : 1;
    if (!aHas && !bHas) return a.index - b.index;
    const ownerCompare = compareStaffNameBySettingsOrder_(a.obj["所有者"], b.obj["所有者"]);
    if (ownerCompare !== 0) return ownerCompare;
    const dueCompare = toTime_(getDriverLicenseExpiryDate_(a.obj)) - toTime_(getDriverLicenseExpiryDate_(b.obj));
    if (dueCompare !== 0) return dueCompare;
    return a.index - b.index;
  });

  range.setValues(indexed.map(item => item.values));
  ensureIdsForSheet_(sheet);
  updateNoticeColumnForSheet_(sheet);
  setCheckboxesForDataRows(sheet);
  applyDriverLicenseManagementLightFormat_();
  toast_("運転免許管理を所有者50音順に整理しました。設定シートの担当者順がある場合はその順を優先します。紙フォルダも同じ順にすると照合しやすいです。");
}

function sortDriverLicenseDetailByOwnerKana() {
  const sheet = ensureDriverLicenseDetailSheet_();
  if (!sheet || sheet.getLastRow() < 3) {
    toast_("運転免許明細は並び替えるデータがありません");
    return;
  }

  const headers = getHeaders_(sheet);
  const rowCount = sheet.getLastRow() - 1;
  const colCount = headers.length;
  const range = sheet.getRange(2, 1, rowCount, colCount);
  const rows = range.getValues();
  const typeOrder = {};
  LICENSE_TYPE_HEADERS.forEach((name, index) => typeOrder[name] = index + 1);

  const indexed = rows.map((values, index) => ({values, index, obj: objectFromRow(headers, values)}));
  indexed.sort((a, b) => {
    const aHas = isBusinessDataRow_(headers, a.values, DRIVER_LICENSE_DETAIL_SHEET_NAME);
    const bHas = isBusinessDataRow_(headers, b.values, DRIVER_LICENSE_DETAIL_SHEET_NAME);
    if (aHas !== bHas) return aHas ? -1 : 1;
    if (!aHas && !bHas) return a.index - b.index;
    const ownerCompare = compareStaffNameBySettingsOrder_(a.obj["所有者"], b.obj["所有者"]);
    if (ownerCompare !== 0) return ownerCompare;
    const ta = typeOrder[String(a.obj["免許種類"] || "").trim()] || 999999;
    const tb = typeOrder[String(b.obj["免許種類"] || "").trim()] || 999999;
    if (ta !== tb) return ta - tb;
    const dateCompare = toTime_(a.obj["取得日"]) - toTime_(b.obj["取得日"]);
    if (dateCompare !== 0) return dateCompare;
    return a.index - b.index;
  });

  range.setValues(indexed.map(item => item.values));
  ensureIdsForSheet_(sheet);
  hideSystemColumnsForSheet_(sheet);
  toast_("運転免許明細を所有者50音順＋免許種類順に整理しました");
}

function alignLicenseSheetLayout() {
  const sheet = repairDriverLicenseManagementHeadersLight_();
  if (!sheet) {
    toast_("運転免許管理シートが見つかりません");
    return;
  }
  applyDriverLicenseManagementLightFormat_();
  toast_("運転免許シートを、20行以降も含めて見やすい行高・列幅に整えました");
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
    applyDisplayLayoutForSheet_(sheet, "compact");
  });

  if (missing.length) {
    toast_("免許・資格・備品だけコンパクト表示を実行しました。一部シートなし: " + missing.join(", "));
  } else {
    toast_("免許・資格・備品だけコンパクト表示を実行しました");
  }
}

function applyCompactLicenseEquipmentLayout_(sheet) {
  applyLicenseEquipmentDisplayLayout_(sheet, "compact");
}

function applyLicenseEquipmentDisplayLayout_(sheet, mode) {
  if (!sheet || sheet.getLastColumn() < 1) return;

  const sheetName = sheet.getName();
  const headers = getHeaders_(sheet);
  const minDisplayRows = sheetName === "運転免許管理" ? DRIVER_LICENSE_DISPLAY_MIN_ROWS : 2;
  const lastRow = Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow(), minDisplayRows));
  const lastCol = headers.length;
  const personal = getPersonalMembers_();
  const compact = mode === "compact";

  sheet.setFrozenColumns(0);
  showAllColumns_(sheet);

  headers.forEach((header, index) => {
    const col = index + 1;
    if (!header) return;

    let width = compact ? 92 : 112;

    if (sheetName === "運転免許管理") {
      if (header === "所有者") width = compact ? 112 : 135;
      else if (LICENSE_TYPE_HEADERS.includes(header)) width = compact ? 52 : 58;
      else if (["取得日", "更新期限", DRIVER_LICENSE_ISSUE_DATE_HEADER, DRIVER_LICENSE_EXPIRY_DATE_HEADER].includes(header)) width = compact ? 106 : 120;
      else if (header === "コピー有無") width = compact ? 88 : 100;
      else if (header === "状態") width = compact ? 92 : 106;
      else if (header === "通知") width = compact ? 98 : 112;
      else if (header === READ_HEADER || personal.includes(header)) width = compact ? 36 : 42;
      else if (header === "備考") width = compact ? 160 : 205;
      else if (header === "免許ID") width = 40;
    } else if (sheetName === "資格管理") {
      if (header === "所有者") width = compact ? 105 : 125;
      else if (header === "資格名") width = compact ? 155 : 190;
      else if (header === "区分") width = compact ? 74 : 90;
      else if (["取得日", "更新期限"].includes(header)) width = compact ? 104 : 120;
      else if (header === "コピー有無") width = compact ? 88 : 100;
      else if (header === "状態") width = compact ? 92 : 106;
      else if (header === "通知") width = compact ? 98 : 112;
      else if (header === READ_HEADER || personal.includes(header)) width = compact ? 36 : 42;
      else if (header === "備考") width = compact ? 150 : 200;
      else if (header === "資格ID") width = 40;
    } else if (sheetName === "備品修理管理") {
      if (header === "購入日") width = compact ? 94 : 112;
      else if (header === "備品名") width = compact ? 115 : 145;
      else if (header === "修理業者") width = compact ? 110 : 140;
      else if (header === "内容") width = compact ? 150 : 200;
      else if (header === "修理依頼者") width = compact ? 90 : 115;
      else if (["修理依頼日", "返却予定日"].includes(header)) width = compact ? 94 : 112;
      else if (header === "返却済") width = compact ? 62 : 82;
      else if (header === "状態") width = compact ? 92 : 106;
      else if (header === "通知") width = compact ? 98 : 112;
      else if (header === READ_HEADER || personal.includes(header)) width = compact ? 36 : 42;
      else if (header === "備考") width = compact ? 145 : 190;
      else if (header === "修理ID") width = 40;
    }

    sheet.setColumnWidth(col, width);

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

  const idHeader = SHEET_ID_HEADERS[sheetName];
  if (idHeader) hideColumnByHeaderSafely_(sheet, idHeader);
  [CALENDAR_ID_HEADER, "写真", "レシート写真", "PDFリンク"].forEach(header => hideColumnByHeaderSafely_(sheet, header));

  rebuildPersonalCheckGroupForSheet_(sheet);

  const minRowHeightRows = sheetName === "運転免許管理" ? DRIVER_LICENSE_DISPLAY_MIN_ROWS : 20;
  const maxRows = Math.min(sheet.getMaxRows(), Math.max(lastRow, minRowHeightRows));
  sheet.setRowHeight(1, sheetName === "運転免許管理" ? (compact ? 72 : 80) : (compact ? 62 : 70));
  sheet.setRowHeights(2, Math.max(maxRows - 1, 1), compact ? 32 : 36);
  sheet.getRange(1, 1, 1, lastCol)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  applyColorRules(sheet);
  createFilterSafely_(sheet, lastCol);
}


/**
 * v9.7.48p14:
 * AppSheet用ID列・カレンダーID列は内部処理に必要なので削除しない。
 * ただし通常入力では邪魔になるため、表示調整や診断作成後に必ず非表示へ戻す。
 */
function hideSystemColumnsForAllInputSheets() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const result = hideSystemColumnsForAllInputSheets_({silent: true});
    toast_("ID列・カレンダーID列を非表示にしました（対象" + result.sheetCount + "シート / " + result.hiddenCount + "列）");
  } finally {
    lock.releaseLock();
  }
}

function hideSystemColumnsForAllInputSheets_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetNames = getSystemColumnHideTargetSheetNames_();
  let sheetCount = 0;
  let hiddenCount = 0;

  targetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheetCount++;
    hiddenCount += hideSystemColumnsForSheet_(sheet);
  });

  return {sheetCount: sheetCount, hiddenCount: hiddenCount};
}

function getSystemColumnHideTargetSheetNames_() {
  const names = [];
  const seen = {};

  function add(name) {
    if (!name || seen[name]) return;
    seen[name] = true;
    names.push(name);
  }

  getInputSheetNames_().forEach(add);
  Object.keys(SHEET_ID_HEADERS || {}).forEach(add);
  return names;
}

function hideSystemColumnsForSheet_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return 0;

  const sheetName = sheet.getName();
  const headers = getHeaders_(sheet);
  let count = 0;

  headers.forEach((header, index) => {
    if (!isSystemColumnHeaderToHide_(sheetName, header)) return;
    const col = index + 1;
    try {
      sheet.hideColumns(col);
      count++;
    } catch (e) {
      console.log("hideSystemColumnsForSheet_ skipped: " + sheetName + " / " + header + " / " + e.message);
    }
  });

  return count;
}

function isSystemColumnHeaderToHide_(sheetName, header) {
  const h = String(header || "").trim();
  if (!h) return false;
  if (h === CALENDAR_ID_HEADER) return true;
  if (h === "ToDo_ID") return true;

  const idHeader = SHEET_ID_HEADERS[sheetName];
  if (idHeader && h === idHeader) return true;

  // 予定・履歴・入力系の内部IDは基本的に末尾がID。
  // 「資格ID」「工事ID」「車両ID」「フィードバックID」などをまとめて隠す。
  if (/ID$/.test(h)) return true;

  return false;
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
      rows.push(["表示", name, "全入力シート標準表示", "NG", "シートがありません", "シート名を確認"]);
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
      "全入力シート標準表示",
      status,
      "列数: " + headers.length + " / 固定列: " + frozenCols + " / ID列非表示: " + (hiddenId ? "はい" : "いいえ") + " / カレンダーID非表示: " + (hiddenCalendar ? "はい" : "いいえ") + (notes.length ? " / " + notes.join(" / ") : ""),
      "全入力シートは『全入力シートを標準表示に整える（少し大きめ）』で読みやすさ優先。横スクロールを減らしたいシートだけ『今のシートだけコンパクト表示』を使う"
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
      "全体は標準表示、免許・資格・備品だけ詰めたい場合は『免許・資格・備品だけコンパクト表示』を実行"
    ]);
  });
}


function formatActiveSheetOnly() {
  formatSheetBase_(SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getLastColumn());
  formatSheetByName_(SpreadsheetApp.getActiveSheet().getName());
  hideSystemColumnsForSheet_(SpreadsheetApp.getActiveSheet());
  toast_("今のシートを表示調整しました（ID列非表示）");
}

function formatSheetByName_(sheetName) {
  if (sheetName === "ダッシュボード") {
    createDashboard();
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  if (!headers.length) return;
  formatSheetBase_(sheet, headers.length);
  const widthMap = {
    "内容": 260, "備考": 260, "対応メモ": 260, "日報文章": 420, "問題点": 260, "明日の予定": 260,
    "現場": 180, "工事名": 180, "行き先": 160, "車両名": 160, "車番": 160, "写真": 120,
    "通知": 120, "状態": 120, "車検状態": 120, "電話対応": 120, "開始時刻": 90, "終了時刻": 90, "時間": 90,
    "契約日": 110, "契約金額": 145, "税": 70, "連絡先": 130, "折り返し電話番号": 150, "対応完了日時": 150, "修理依頼者": 130, "修理業者": 160, "PDFリンク": 260, "フィードバックID": 180
  };
  headers.forEach((h, i) => sheet.setColumnWidth(i + 1, widthMap[h] || 115));
  if (sheet.getLastRow() > 1) sheet.setRowHeights(2, Math.min(sheet.getLastRow() - 1, 200), 34);
  try { sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length).setWrap(true).setVerticalAlignment("middle"); } catch (e) {}
  hideSystemColumnsForSheet_(sheet);
  applyColorRules(sheet);
  applyPersonalCheckColumnLayout_(sheet);
}

function formatSheetBase_(sheet, colCount) {
  if (!sheet) return;
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, colCount).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sheet.setRowHeight(1, 42);
  // p29: 通常の整形処理後も、AppSheet用ID列・カレンダーID列はデフォルト非表示に戻す。
  // ヘッダーがまだ無い作成途中のシートでは何もしない。
  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
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
  addTextRule_(rules, sheet, noticeCol, "完了", "#d9ead3", rowCount); // 淡い緑
  addTextRule_(rules, sheet, noticeCol, "重要", "#ead1dc", rowCount); // 淡い赤紫
  addTextRule_(rules, sheet, noticeCol, "未確認", "#fff7d6", rowCount); // さらに淡い黄
  addTextRule_(rules, sheet, noticeCol, "差戻し", "#ead1dc", rowCount); // 淡い赤紫
  addTextRule_(rules, sheet, noticeCol, "未精算", "#fde9d9", rowCount); // 淡い橙

  // 状態：進捗・完了・異常を淡色で分ける。
  addTextRule_(rules, sheet, statusCol, "予定", "#eeeeee", rowCount); // 淡い灰
  addTextRule_(rules, sheet, statusCol, "未契約", "#d9eaf7", rowCount); // 淡い青灰
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
  addTextRule_(rules, sheet, statusCol, "請求済み", "#cfe8e6", rowCount); // 請求済みは完了後の淡い青緑
  addTextRule_(rules, sheet, statusCol, "更新済", "#cfe8e6", rowCount); // 淡い青緑
  addTextRule_(rules, sheet, statusCol, "延期", "#f9cb9c", rowCount); // 橙
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

  // v9.7.48f: 資格管理・資格一覧系は、資格区分/取得区分/保有状況を淡色で見分けやすくする。
  applyQualificationSpecificColorRules_(rules, sheet, headers, rowCount);

  sheet.setConditionalFormatRules(rules);
}

function addTextRuleByHeader_(rules, sheet, headers, header, text, color, rowCount) {
  const col = headers.indexOf(header) + 1;
  addTextRule_(rules, sheet, col, text, color, rowCount);
}

function applyQualificationSpecificColorRules_(rules, sheet, headers, rowCount) {
  if (!sheet || !headers || !headers.length) return;
  const sheetName = sheet.getName();
  const isQualificationSheet = [
    "資格管理",
    "社員別資格一覧",
    "資格別保有者一覧",
    "資格重複チェック",
    "資格かんたん登録",
    "資格一括登録"
  ].includes(sheetName);
  if (!isQualificationSheet) return;

  // 区分：資格名の性質を素早く見分けるための淡色。
  addTextRuleByHeader_(rules, sheet, headers, "区分", "資格", "#d9ead3", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "区分", "免許", "#d9eaf7", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "区分", "技能講習", "#fff2cc", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "区分", "特別教育", "#e4dfec", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "区分", "安全教育", "#fce5cd", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "区分", "その他", "#eeeeee", rowCount);

  // 取得区分：施工管理技士系の技士/技士補と、講習系の修了を区別。
  addTextRuleByHeader_(rules, sheet, headers, "取得区分", "技士", "#cfe2f3", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "取得区分", "技士補", "#d9eaf7", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "取得区分", "保有", "#d9ead3", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "取得区分", "修了", "#fff2cc", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "取得区分", "免許", "#d0e0e3", rowCount);

  // 保有状況：ここを一番見落としにくくする。
  addTextRuleByHeader_(rules, sheet, headers, "保有状況", "保有", "#d9ead3", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "保有状況", "未確認", "#fff2cc", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "保有状況", "失効", "#d7ccc8", rowCount);
  addTextRuleByHeader_(rules, sheet, headers, "保有状況", "不要", "#eeeeee", rowCount);
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


/******************************
 * v9.7.28 資格管理 1シート集約・重複チェック
 ******************************/

const QUALIFICATION_OLD_INPUT_SHEET_NAMES = [
  "資格入力_資格",
  "資格入力_技能講習",
  "資格入力_特別教育",
  "資格入力_安全教育",
  "資格入力_免許",
  "資格入力_その他"
];


function rebuildQualificationMasterClean() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const count = rebuildQualificationMasterClean_();
    toast_("資格マスタを紙版標準へ戻しました（" + count + "件）。手動変更した区分も標準値で上書きされています。");
  } finally {
    lock.releaseLock();
  }
}

function rebuildQualificationMasterClean_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["資格マスタ"] || ["表示順", "区分", "資格名", "期限管理", "有効年数", "通知日前", "一覧表示", "備考", "資格マスタID"];
  let sheet = ss.getSheetByName("資格マスタ");
  if (!sheet) sheet = ss.insertSheet("資格マスタ");

  const existingMap = getExistingQualificationMasterMetaMap_(sheet);

  removeFilter_(sheet);
  showAllColumns_(sheet);
  sheet.clear();
  sheet.clearFormats();
  sheet.setConditionalFormatRules([]);

  const rows = getQualificationPaperMasterRows_().map(item => {
    const key = buildQualificationMasterMetaKey_(item.category, item.name);
    const old = existingMap[key] || {};
    const noticeDays = item.noticeDays !== undefined ? item.noticeDays : (old.noticeDays !== undefined ? old.noticeDays : (item.category === "免許" ? 90 : ""));
    const visible = old.visible !== undefined ? old.visible : true;
    return [
      item.order,
      item.category,
      item.name,
      item.expire || "期限なし",
      item.years || "",
      noticeDays,
      visible,
      item.note || "紙の資格一覧表より",
      old.id || buildRecordId_("資格マスタ")
    ];
  });

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  try {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#d9ead3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.getRange(1, 1, lastRow, headers.length)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment("middle");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 70);
    sheet.setColumnWidth(2, 95);
    sheet.setColumnWidth(3, 280);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 80);
    sheet.setColumnWidth(6, 85);
    sheet.setColumnWidth(7, 85);
    sheet.setColumnWidth(8, 210);
    sheet.setColumnWidth(9, 150);
    const noticeCol = headers.indexOf("通知日前") + 1;
    if (noticeCol > 0 && lastRow >= 2) sheet.getRange(2, noticeCol, lastRow - 1, 1).setNumberFormat("0").clearDataValidations();
    const visibleCol = headers.indexOf("一覧表示") + 1;
    if (visibleCol > 0 && lastRow >= 2) sheet.getRange(2, visibleCol, lastRow - 1, 1).insertCheckboxes();
    createFilterSafely_(sheet, headers.length);
  } catch (e) {}

  return rows.length;
}



function getQualificationPaperMasterRows_() {
  // p28:
  // 紙の資格一覧表の左から右の並びを「表示順」として持つ。
  // 区分はプルダウン分離用に残し、ソートは表示順を優先する。
  return [
    // 資格
    {order: 1, category: "資格", name: "監理技術者", note: "紙の資格一覧表より"},
    {order: 2, category: "資格", name: "1級土木施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 3, category: "資格", name: "2級土木施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 4, category: "資格", name: "1級管工事施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 5, category: "資格", name: "2級管工事施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 6, category: "資格", name: "1級建設機械施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 7, category: "資格", name: "2級建設機械施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 8, category: "資格", name: "1級舗装施工管理技術者"},
    {order: 9, category: "資格", name: "2級舗装施工管理技術者"},
    {order: 10, category: "資格", name: "1級建築施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 11, category: "資格", name: "2級建築施工管理技士", note: "●技士 / ○技士補対応"},
    {order: 12, category: "資格", name: "解体工事施工技士"},
    {order: 13, category: "資格", name: "排水設備工事責任技術者"},
    {order: 14, category: "資格", name: "給水装置工事主任技術者"},
    {order: 15, category: "資格", name: "建築物石綿含有建材調査者講習"},
    {order: 16, category: "資格", name: "工作物石綿含有建材調査者講習"},
    {order: 17, category: "資格", name: "測量士"},
    {order: 18, category: "資格", name: "1級配管技能士"},
    {order: 19, category: "資格", name: "2級配管技能士"},
    {order: 20, category: "資格", name: "2級建設業経理事務士"},
    {order: 21, category: "資格", name: "3級建設業経理事務士"},

    // 免許
    {order: 22, category: "免許", name: "大型自動車免許", expire: "期限あり", note: "運転免許管理と併用"},
    {order: 23, category: "免許", name: "大型特殊免許", expire: "期限あり", note: "運転免許管理と併用"},
    {order: 24, category: "免許", name: "中型免許", expire: "期限あり", note: "運転免許管理と併用"},
    {order: 25, category: "免許", name: "準中型免許", expire: "期限あり", note: "運転免許管理と併用"},

    // 技能講習・特別教育・安全教育（紙面の横並び順）
    {order: 26, category: "技能講習", name: "地山掘削作業主任者"},
    {order: 27, category: "技能講習", name: "土止め支保工作業主任者"},
    {order: 28, category: "技能講習", name: "足場の組立て等作業主任者"},
    {order: 29, category: "技能講習", name: "型枠支保工の組立て等作業主任者"},
    {order: 30, category: "技能講習", name: "コンクリート工作物の解体等作業主任者"},
    {order: 31, category: "技能講習", name: "有機溶剤作業主任者"},
    {order: 32, category: "技能講習", name: "酸素欠乏・硫化水素危険作業主任者"},
    {order: 33, category: "技能講習", name: "特定化学物質等作業主任者"},
    {order: 34, category: "技能講習", name: "石綿作業主任者"},
    {order: 35, category: "技能講習", name: "移動式クレーン"},
    {order: 36, category: "技能講習", name: "クレーン"},
    {order: 37, category: "技能講習", name: "玉掛け"},
    {order: 38, category: "技能講習", name: "車両系建設機械（整地）"},
    {order: 39, category: "技能講習", name: "車両系建設機械（解体3種）"},
    {order: 40, category: "特別教育", name: "車両系建設機械（締固）"},
    {order: 41, category: "技能講習", name: "不整地運搬車"},
    {order: 42, category: "特別教育", name: "高所作業車"},
    {order: 43, category: "技能講習", name: "フォークリフト"},
    {order: 44, category: "特別教育", name: "研削砥石"},
    {order: 45, category: "特別教育", name: "アーク溶接"},
    {order: 46, category: "技能講習", name: "ガス溶接"},
    {order: 47, category: "特別教育", name: "伐木等機械運転"},
    {order: 48, category: "特別教育", name: "走行集材機械運転"},
    {order: 49, category: "特別教育", name: "簡易架線集材等運転"},
    {order: 50, category: "特別教育", name: "ゴンドラ"},
    {order: 51, category: "特別教育", name: "チェーンソー作業者"},
    {order: 52, category: "特別教育", name: "立木の伐木作業者"},
    {order: 53, category: "特別教育", name: "フルハーネス型墜落制止用器具"},
    {order: 54, category: "特別教育", name: "石綿取扱作業従事者"},
    {order: 55, category: "特別教育", name: "テールゲートリフター"},
    {order: 56, category: "安全教育", name: "職長"},
    {order: 57, category: "安全教育", name: "刈払い機"},
    {order: 58, category: "安全教育", name: "振動工具"},
    {order: 59, category: "安全教育", name: "除雪講習"},

    // 紙面外だが既存運用で候補として残すもの。紙面順の後ろへ回す。
    {order: 60, category: "免許", name: "普通自動車免許", expire: "期限あり", note: "運転免許管理と併用"},
    {order: 61, category: "技能講習", name: "小型移動式クレーン"},
    {order: 62, category: "技能講習", name: "床上操作式クレーン"},
    {order: 63, category: "技能講習", name: "ローラーの運転"},
    {order: 64, category: "安全教育", name: "職長・安全衛生責任者教育"},
    {order: 65, category: "安全教育", name: "安全衛生責任者教育"}
  ];
}

function repairQualificationSingleSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const sheet = ensureQualificationManagementSheet_();
    applyQualificationSingleSheetLightSettings_(sheet);
    sortQualificationManagementByOwner_();
    createQualificationDuplicateCheckSheet_();
    toast_("資格管理を1シート形式に修復しました。重複チェックも作成しました。");
  } finally {
    lock.releaseLock();
  }
}

function sortQualificationManagementByOwner() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const count = sortQualificationManagementByOwner_();
    toast_("資格管理を所有者＋紙マスタ順に整理しました。空欄行は下へ移動し、資格名プルダウンも復旧しました（" + count + "件）");
  } finally {
    lock.releaseLock();
  }
}

function groupQualificationManagementByOwner() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const groupCount = groupQualificationManagementByOwner_();
    try {
      cleanupQualificationManagementFalseDisplay_();
      refreshQualificationManagementCheckboxes_();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
      if (sheet) {
        applyQualificationManagementCompactDisplay_(sheet);
        hideSystemColumnsForSheet_(sheet);
      }
    } catch (e) {
      console.log("資格管理グループ化後の軽量表示復旧をスキップ: " + e.message);
    }
    toast_("資格管理を所有者ごとにグループ化しました（" + groupCount + "名）");
  } finally {
    lock.releaseLock();
  }
}

function formatQualificationSheetsAndOwnerGroups() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const easy = ss.getSheetByName("資格かんたん登録");
    if (easy) {
      formatQualificationEasyPasteSheet_(easy);
      repairQualificationEasyPasteColumnValidations_(easy);
    }

    const management = ss.getSheetByName("資格管理");
    if (management) {
      applyQualificationManagementCompactDisplay_(management);
      groupQualificationManagementByOwner_();
    }

    const employeeCount = createEmployeeQualificationListSheet_();
    toast_("資格まわりを整えました。資格管理は所有者ごとにグループ化、社員別資格一覧は人ごと展開で更新しました（" + employeeCount + "件）");
  } finally {
    lock.releaseLock();
  }
}

function createQualificationDuplicateCheckSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const count = createQualificationDuplicateCheckSheet_();
    if (count > 0) {
      toast_("資格管理の重複候補を " + count + " 件出力しました");
    } else {
      toast_("資格管理の重複候補はありません");
    }
  } finally {
    lock.releaseLock();
  }
}

function hideOldQualificationInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;
  QUALIFICATION_OLD_INPUT_SHEET_NAMES.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    try {
      sheet.hideSheet();
      count++;
    } catch (e) {}
  });
  toast_("横長資格入力シートを非表示にしました（" + count + "シート）");
}

function deleteOldQualificationInputSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let count = 0;
  QUALIFICATION_OLD_INPUT_SHEET_NAMES.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    try {
      ss.deleteSheet(sheet);
      count++;
    } catch (e) {
      console.log("資格入力シート削除スキップ: " + name + " / " + e.message);
    }
  });
  toast_("横長資格入力シートを削除しました（" + count + "シート）。資格管理本体は削除していません。");
}


function migrateQualificationAcquisitionStatusColumns_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  let headers = getHeaders_(sheet);
  if (!headers.length) return;

  const oldIndex = headers.indexOf("保有区分");
  const acquisitionIndex = headers.indexOf("取得区分");
  if (oldIndex >= 0 && acquisitionIndex < 0) {
    sheet.getRange(1, oldIndex + 1).setValue("取得区分");
    headers = getHeaders_(sheet);
  }

  if (headers.indexOf("保有状況") < 0) {
    const acqCol = headers.indexOf("取得区分") + 1;
    const qualificationCol = headers.indexOf("資格名") + 1;
    const ownerCol = headers.indexOf("所有者") + 1;
    const insertAfter = acqCol > 0 ? acqCol : (qualificationCol > 0 ? qualificationCol : Math.min(headers.length, 3));
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue("保有状況");
    headers = getHeaders_(sheet);

    const statusCol = headers.indexOf("保有状況") + 1;
    if (statusCol > 0 && sheet.getLastRow() >= 2) {
      const lastRow = sheet.getLastRow();
      const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      const statusValues = data.map(row => {
        const owner = ownerCol > 0 ? row[ownerCol - 1] : "";
        const qualification = qualificationCol > 0 ? row[qualificationCol - 1] : "";
        return [(owner || qualification) ? "保有" : ""];
      });
      sheet.getRange(2, statusCol, statusValues.length, 1).setValues(statusValues);
    }
  }
}

function ensureQualificationManagementSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("資格管理");
  if (!sheet) sheet = ss.insertSheet("資格管理");
  migrateQualificationAcquisitionStatusColumns_(sheet);

  const desiredHeaders = getSheetHeaders_()["資格管理"];
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    setupSheet_(sheet, desiredHeaders);
    return sheet;
  }

  let headers = getHeaders_(sheet);
  if (!headers.length || !headers[0]) {
    setupSheet_(sheet, desiredHeaders);
    return sheet;
  }

  // 既存データを極力残したまま、不足列だけ追加する。
  // 取得区分は「資格名」の右へ追加する。
  if (headers.indexOf("取得区分") < 0) {
    const qualificationNameCol = headers.indexOf("資格名") + 1;
    const insertAfter = qualificationNameCol > 0 ? qualificationNameCol : Math.min(headers.length, 3);
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue("取得区分");
    headers = getHeaders_(sheet);
  }

  desiredHeaders.forEach(header => {
    headers = getHeaders_(sheet);
    if (headers.indexOf(header) >= 0) return;
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue(header);
  });

  reorderQualificationManagementColumns_(sheet, desiredHeaders);
  return sheet;
}

function reorderQualificationManagementColumns_(sheet, desiredHeaders) {
  let headers = getHeaders_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const currentValues = sheet.getRange(1, 1, lastRow, headers.length).getValues();

  const indexMap = {};
  headers.forEach((h, i) => { if (h) indexMap[h] = i; });

  const reordered = currentValues.map(row => {
    return desiredHeaders.map(header => {
      const oldIndex = indexMap[header];
      return oldIndex === undefined ? "" : row[oldIndex];
    });
  });

  // 旧列があっても資格管理は指定列だけに絞る。
  resetSheetLight_(sheet, desiredHeaders.length);
  sheet.getRange(1, 1, reordered.length, desiredHeaders.length).setValues(reordered);
}

function applyQualificationSingleSheetLightSettings_(sheet) {
  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  formatSheetBase_(sheet, headers.length);
  createFilterSafely_(sheet, headers.length);
  applyDataValidationByHeaders_(sheet, headers);
  applyColorRules(sheet);

  // v9.7.33:
  // 資格ID列の長い文字列が折り返されて行高が大きくなるため、
  // 資格管理だけは専用コンパクト表示でID非表示・折り返し停止・行高固定を行う。
  applyQualificationManagementCompactDisplay_(sheet);
}

function setColumnWidthIfExists_(sheet, headers, header, width) {
  const col = headers.indexOf(header) + 1;
  if (col > 0) sheet.setColumnWidth(col, width);
}


function compactQualificationManagementSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
  if (!sheet) {
    toast_("資格管理シートが見つかりません");
    return;
  }

  applyQualificationManagementCompactDisplay_(sheet);
  applyColorRules(sheet);
  toast_("資格管理をコンパクト表示にしました（資格ID非表示・行高固定）");
}

function formatQualificationManagementColors() {
  formatQualificationManagementColors_();
  toast_("資格管理・社員別資格一覧のカラーリングを再調整しました");
}

function formatQualificationManagementColors_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["資格管理", "社員別資格一覧", "資格別保有者一覧"].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    if (name === "資格管理") applyQualificationManagementCompactDisplay_(sheet);
    applyColorRules(sheet);
  });
}

function applyQualificationManagementCompactDisplay_(sheet) {
  if (!sheet || sheet.getName() !== "資格管理" || sheet.getLastColumn() < 1) return;

  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const lastCol = headers.length;
  const dataRows = Math.max(lastRow - 1, 1);
  const personal = getPersonalMembers_();

  try {
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(0);
    showAllColumns_(sheet);

    // v9.7.38:
    // 資格管理は正式データなので、入力しやすさより「一覧性」と「行高固定」を優先。
    // 資格IDは非表示、備考だけ広め、その他は中央揃えにする。
    const allRange = sheet.getRange(1, 1, lastRow, lastCol);
    allRange
      .setWrap(false)
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#d9ead3")
      .setWrap(false);

    const widths = {
      "所有者": 105,
      "区分": 110,
      "資格名": 300,
      "取得区分": 90,
      "保有状況": 90,
      "取得日": 105,
      "更新期限": 105,
      "コピー有無": 95,
      "状態": 105,
      "通知": 105,
      "備考": 230,
      "資格ID": 55
    };

    const centerHeaders = ["所有者", "区分", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", READ_HEADER];
    headers.forEach((header, index) => {
      if (!header) return;
      const col = index + 1;
      let width = widths[header] || 90;
      if (header === READ_HEADER || personal.includes(header)) width = 42;
      sheet.setColumnWidth(col, width);

      const range = sheet.getRange(1, col, lastRow, 1);
      if (header === READ_HEADER || personal.includes(header)) {
        // v9.7.48p4:
        // チェック列全体を縦書きにすると、FALSE が A L S のように見える。
        // ヘッダーだけ縦書き、データ行は通常表示にする。
        sheet.getRange(1, col)
          .setVerticalText(true)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setWrap(false);

        if (lastRow >= 2) {
          sheet.getRange(2, col, dataRows, 1)
            .setVerticalText(false)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setWrap(false);
        }
      } else if (centerHeaders.includes(header)) {
        range
          .setVerticalText(false)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setWrap(false);
      } else if (header === "資格名" || header === "備考") {
        range
          .setVerticalText(false)
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle")
          .setWrap(false);
      } else {
        range
          .setVerticalText(false)
          .setVerticalAlignment("middle")
          .setWrap(false);
      }
    });

    sheet.setRowHeight(1, 46);
    sheet.setRowHeights(2, dataRows, 34);

    hideColumnByHeaderSafely_(sheet, "資格ID");
    createFilterSafely_(sheet, lastCol);
  } catch (e) {
    console.log("applyQualificationManagementCompactDisplay_ error: " + e.message);
  }
}

function sortQualificationManagementByOwner_() {
  // p31:
  // 旧処理はデータ範囲全体の入力規則を一度消してから、後続処理で復旧する構成だったため、
  // 実行時間超過や途中停止が起きるとC列「資格名」のプルダウンが消えたままになることがあった。
  // ここでは、値の並び替えは軽量ソートへ一本化し、C列の入力規則だけを最後に必ず復旧する。
  const sheet = ensureQualificationManagementSheet_();
  if (!sheet) return 0;

  let count = 0;
  try {
    count = sortQualificationManagementRowsOnlyByOwnerMaster_(sheet);
  } catch (err) {
    console.log("sortQualificationManagementByOwner_ light sort error: " + err.message);
    throw err;
  }

  try { cleanupQualificationManagementFalseDisplayFast_(); } catch (e) {}
  try { refreshQualificationManagementCheckboxesFast_(); } catch (e) {}
  try { refreshQualificationManagementDropdownsLight_(sheet); } catch (e) {}
  try { applyQualificationManagementMinimalDisplay_(sheet); } catch (e) {}
  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}

  return count;
}

function groupQualificationManagementByOwner_() {
  const sheet = ensureQualificationManagementSheet_();

  // 先に並び替えてから、既存の行グループを完全解除する。
  // グループが残ったまま setValues/sort を行うと、別所有者の途中にグループが残ることがある。
  const count = sortQualificationManagementByOwner_();
  if (count < 2) {
    clearRowGroupsSafe_(sheet, 2, Math.max(sheet.getLastRow(), 2));
    return 0;
  }

  SpreadsheetApp.flush();

  const headers = getHeaders_(sheet);
  const ownerCol = headers.indexOf("所有者") + 1;
  if (ownerCol <= 0) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    clearRowGroupsSafe_(sheet, 2, Math.max(lastRow, 2));
    return 0;
  }

  // 資格管理だけを対象に、既存グループを完全に外してから作り直す。
  clearRowGroupsSafe_(sheet, 2, lastRow);

  try {
    sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  } catch (e) {}

  const ownerValues = sheet.getRange(2, ownerCol, lastRow - 1, 1).getValues();
  let groupCount = 0;
  let startRow = 2;
  let currentOwner = normalizeTextForKey_(ownerValues[0][0]);

  for (let i = 1; i <= ownerValues.length; i++) {
    const rowNo = i + 2;
    const owner = i < ownerValues.length ? normalizeTextForKey_(ownerValues[i][0]) : "__END__";

    if (owner !== currentOwner) {
      const endRow = rowNo - 1;

      // 所有者が空欄の行はグループ化しない。
      // 2行以上ある所有者だけグループ化する。
      // 先頭行は見出し兼1件目として残し、2行目以降だけ折りたたむ。
      if (currentOwner && endRow > startRow) {
        groupRangeRowsSafe_(sheet, startRow + 1, endRow);
        groupCount++;
      }

      startRow = rowNo;
      currentOwner = owner;
    }
  }

  applyQualificationManagementCompactDisplay_(sheet);
  return groupCount;
}

function groupRangeRowsSafe_(sheet, startRow, endRow) {
  if (!sheet || endRow < startRow) return;
  try {
    sheet.getRange(startRow, 1, endRow - startRow + 1, 1).shiftRowGroupDepth(1);
  } catch (e) {
    console.log("groupRangeRowsSafe_ error: " + e.message);
  }
}

function clearRowGroupsSafe_(sheet, startRow, endRow) {
  if (!sheet) return;
  const first = Math.max(1, startRow || 1);
  const last = Math.max(first, endRow || sheet.getLastRow() || sheet.getMaxRows());
  const rowCount = last - first + 1;

  try { sheet.expandAllRowGroups(); } catch (e) {}

  // 深さが残る場合に備えて複数回下げる。
  // 範囲全体で失敗した場合は、連続範囲を小さくして再試行する。
  for (let depth = 0; depth < 8; depth++) {
    try {
      sheet.getRange(first, 1, rowCount, 1).shiftRowGroupDepth(-1);
    } catch (e1) {
      try {
        for (let r = last; r >= first; r--) {
          const d = typeof sheet.getRowGroupDepth === "function" ? sheet.getRowGroupDepth(r) : 0;
          if (d > 0) sheet.getRange(r, 1, 1, 1).shiftRowGroupDepth(-1);
        }
      } catch (e2) {}
    }
  }
}

function createQualificationDuplicateCheckSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ensureQualificationManagementSheet_();
  const headers = getHeaders_(source);
  const outHeaders = ["判定", "所有者", "区分", "資格名", "取得区分", "保有状況", "件数", "元行", "備考"];
  let output = ss.getSheetByName("資格重複チェック");
  if (!output) output = ss.insertSheet("資格重複チェック");

  output.clear();
  output.clearFormats();
  output.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);

  const lastRow = source.getLastRow();
  if (lastRow < 2) {
    formatSheetBase_(output, outHeaders.length);
    return 0;
  }

  const ownerIndex = headers.indexOf("所有者");
  const categoryIndex = headers.indexOf("区分");
  const qualificationIndex = headers.indexOf("資格名");
  const possessionIndex = headers.indexOf("取得区分");
  const holdingStatusIndex = headers.indexOf("保有状況");
  const idIndex = headers.indexOf("資格ID");
  const values = source.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const map = {};
  const blankRows = [];

  values.forEach((row, index) => {
    const rowNo = index + 2;
    if (!isBusinessDataRow_(headers, row, "資格管理")) return;

    const owner = ownerIndex >= 0 ? row[ownerIndex] : "";
    const category = categoryIndex >= 0 ? row[categoryIndex] : "";
    const qualification = qualificationIndex >= 0 ? row[qualificationIndex] : "";
    const possession = possessionIndex >= 0 ? row[possessionIndex] : "";
    const holdingStatus = holdingStatusIndex >= 0 ? row[holdingStatusIndex] : "";
    const idValue = idIndex >= 0 ? row[idIndex] : "";

    if (!normalizeTextForKey_(owner) || !normalizeTextForKey_(qualification)) {
      blankRows.push(["要確認", owner, category, qualification, possession, holdingStatus, 1, rowNo, "所有者または資格名が空欄です" + (idValue ? " / ID:" + idValue : "")]);
      return;
    }

    // 重複判定は、同じ人・同じ区分・同じ資格名を基本にする。
    // 技士/技士補/保有が違っても同じ資格が複数ある場合は確認対象にする。
    const key = [owner, category, qualification].map(normalizeTextForKey_).join("||");
    if (!map[key]) {
      map[key] = { owner: owner, category: category, qualification: qualification, possessions: {}, holdingStatuses: {}, rows: [] };
    }
    map[key].possessions[String(possession || "").trim() || "空欄"] = true;
    map[key].holdingStatuses[String(holdingStatus || "").trim() || "空欄"] = true;
    map[key].rows.push(rowNo);
  });

  const out = [];
  Object.keys(map).forEach(key => {
    const item = map[key];
    if (item.rows.length <= 1) return;
    out.push([
      "重複候補",
      item.owner,
      item.category,
      item.qualification,
      Object.keys(item.possessions).join(" / "),
      Object.keys(item.holdingStatuses).join(" / "),
      item.rows.length,
      item.rows.join(", "),
      "同じ所有者・区分・資格名が複数行あります"
    ]);
  });

  const rows = out.concat(blankRows);
  if (rows.length) {
    output.getRange(2, 1, rows.length, outHeaders.length).setValues(rows);
  }

  formatSheetBase_(output, outHeaders.length);
  try {
    output.setColumnWidth(1, 90);
    output.setColumnWidth(2, 90);
    output.setColumnWidth(3, 100);
    output.setColumnWidth(4, 240);
    output.setColumnWidth(5, 110);
    output.setColumnWidth(6, 60);
    output.setColumnWidth(7, 120);
    output.setColumnWidth(8, 120);
    output.setColumnWidth(9, 260);
    createFilterSafely_(output, outHeaders.length);
  } catch (e) {}

  return rows.length;
}

function getQualificationCategoryOrder_(category) {
  const value = String(category || "").trim();
  const order = {
    "資格": 1,
    "免許": 2,
    "技能講習": 3,
    "特別教育": 4,
    "安全教育": 5,
    "その他": 9
  };
  return order[value] || 99;
}

function getQualificationMasterOrderMap_() {
  const map = {};
  getQualificationMasterItems_().forEach(item => {
    const category = String(item.category || "その他").trim();
    const name = cleanQualificationName_(item.name);
    if (!name) return;
    const order = Number(item.order) || 999999;
    const nameKey = normalizeTextForKey_(name);
    const categoryKey = normalizeTextForKey_(category);
    const categoryNameKey = categoryKey + "||" + nameKey;
    if (!map[categoryNameKey] || order < map[categoryNameKey]) map[categoryNameKey] = order;
    const nameOnlyKey = "||" + nameKey;
    if (!map[nameOnlyKey] || order < map[nameOnlyKey]) map[nameOnlyKey] = order;
  });
  return map;
}

function getQualificationMasterSortOrder_(category, qualificationName, masterOrderMap) {
  const nameKey = normalizeTextForKey_(cleanQualificationName_(qualificationName));
  if (!nameKey) return 999999;
  const categoryKey = normalizeTextForKey_(category);
  const map = masterOrderMap || getQualificationMasterOrderMap_();
  return map[categoryKey + "||" + nameKey] || map["||" + nameKey] || 999999;
}

function normalizeTextForKey_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/[　\s]+/g, " ");
}

function getDateTimeSortValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  const d = new Date(value);
  if (!isNaN(d.getTime())) return d.getTime();
  return 9999999999999;
}




/******************************
 * v9.7.34 資格かんたん登録シート方式
 * v9.7.36:
 * - 資格かんたん登録に「追加する資格」列を追加。
 * - プルダウンで資格を選ぶと「資格名まとめ」へ自動追記。
 * - 同じ資格は二重追記しない。正式登録時も重複追加しない。
 * 1行に複数資格を貼り付けて、資格管理へ1人1資格1行で展開する。
 ******************************/

function createQualificationEasyPasteSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    ensureQualificationEasyPasteSheet_();
    toast_("資格かんたん登録シートを作成/修復しました");
  } finally {
    lock.releaseLock();
  }
}

function repairQualificationEasyPasteDropdownAppendSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    // v9.7.46:
    // 既存の「保有区分」列が残っている場合も、ここで必ず「取得区分＋保有状況」へ移行する。
    // 入力規則のズレだけでなく、ヘッダー順のズレも軽量に修復する。
    const sheet = ensureQualificationEasyPasteSheet_();
    normalizeAndSortQualificationEasyPasteSheet_(sheet, false);
    toast_("資格かんたん登録の列・入力規則を修復しました（資格名まとめ列は自由入力です）");
  } finally {
    lock.releaseLock();
  }
}

function applyQualificationEasyPasteToManagement() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const result = applyQualificationEasyPasteToManagement_();
    toast_("資格かんたん登録を反映しました。追加 " + result.added + "件 / 重複スキップ " + result.skipped + "件 / 取得区分補正 " + (result.corrected || 0) + "件 / 要確認 " + result.invalid + "件");
  } finally {
    lock.releaseLock();
  }
}

function applyQualificationEasyPasteUpdateExisting() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const result = applyQualificationEasyPasteToManagementUpdateExisting_(false);
    toast_(
      "資格かんたん登録を更新反映しました。追加 " + result.added +
      "件 / 更新 " + result.updated +
      "件 / 変更なし " + result.unchanged +
      "件 / 取得区分補正 " + (result.corrected || 0) +
      "件 / 要確認 " + result.invalid + "件"
    );
  } finally {
    lock.releaseLock();
  }
}

function applyQualificationEasyPasteForceOverwrite() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "資格かんたん登録で強制上書き",
      "既存資格の取得区分・取得日・更新期限・コピー有無・状態・備考を空欄も含めて上書きします。実行しますか？",
      ui.ButtonSet.OK_CANCEL
    );
    if (response !== ui.Button.OK) {
      toast_("強制上書きをキャンセルしました");
      return;
    }

    const result = applyQualificationEasyPasteToManagementUpdateExisting_(true);
    toast_(
      "資格かんたん登録を強制上書きしました。追加 " + result.added +
      "件 / 上書き " + result.updated +
      "件 / 変更なし " + result.unchanged +
      "件 / 取得区分補正 " + (result.corrected || 0) +
      "件 / 要確認 " + result.invalid + "件"
    );
  } finally {
    lock.releaseLock();
  }
}

function clearQualificationEasyPasteSheet() {
  const sheet = ensureQualificationEasyPasteSheet_();
  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  toast_("資格かんたん登録シートをクリアしました（ヘッダーは残しています）");
}

function sortQualificationEasyPasteSheetMasterOrderBlankClean() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const sheet = ensureQualificationEasyPasteSheet_();
    const result = normalizeAndSortQualificationEasyPasteSheet_(sheet, true);
    toast_(
      "資格かんたん登録を整理しました。資格名並び替え " + result.summarySorted +
      "行 / 空欄下移動 " + result.blankRows + "行"
    );
  } finally {
    lock.releaseLock();
  }
}


function migrateQualificationEasyPasteHeaders_(sheet, desiredHeaders) {
  if (!sheet || sheet.getName() !== "資格かんたん登録") return;
  if (sheet.getLastRow() < 1) return;

  let headers = getHeaders_(sheet);
  if (!headers.length) return;

  // v9.7.46: 古い「保有区分」を「取得区分」へ改名。
  const oldPossessionIndex = headers.indexOf("保有区分");
  if (oldPossessionIndex >= 0 && headers.indexOf("取得区分") < 0) {
    sheet.getRange(1, oldPossessionIndex + 1).setValue("取得区分");
    headers = getHeaders_(sheet);
  }

  // 「保有状況」がなければ「取得区分」の右に追加。
  if (headers.indexOf("保有状況") < 0) {
    const acquisitionCol = headers.indexOf("取得区分") + 1;
    const insertAfter = acquisitionCol > 0 ? acquisitionCol : Math.min(headers.length, 5);
    sheet.insertColumnAfter(insertAfter);
    sheet.getRange(1, insertAfter + 1).setValue("保有状況");
    headers = getHeaders_(sheet);

    const statusCol = headers.indexOf("保有状況") + 1;
    const summaryCol = headers.indexOf("資格名まとめ") + 1;
    const ownerCol = headers.indexOf("所有者") + 1;
    if (statusCol > 0 && sheet.getLastRow() >= 2) {
      const lastRow = sheet.getLastRow();
      const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      const out = values.map(row => {
        const owner = ownerCol > 0 ? row[ownerCol - 1] : "";
        const summary = summaryCol > 0 ? row[summaryCol - 1] : "";
        return [(owner || summary) ? "保有" : ""];
      });
      sheet.getRange(2, statusCol, out.length, 1).setValues(out);
    }
  }

  // 不足ヘッダーを追加してから、指定順に並べ替える。
  headers = getHeaders_(sheet);
  (desiredHeaders || getSheetHeaders_()["資格かんたん登録"] || []).forEach(header => {
    headers = getHeaders_(sheet);
    if (headers.indexOf(header) >= 0) return;
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue(header);
  });

  const desired = desiredHeaders || getSheetHeaders_()["資格かんたん登録"] || [];
  if (desired.length) repairHeaderOrderPreserveData_(sheet, desired);
}

function ensureQualificationEasyPasteSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["資格かんたん登録"] || ["所有者", "区分", "追加する資格", "資格名まとめ", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "備考", "登録結果"];
  let sheet = ss.getSheetByName("資格かんたん登録");
  if (!sheet) sheet = ss.insertSheet("資格かんたん登録");
  migrateQualificationAcquisitionStatusColumns_(sheet);
  migrateQualificationEasyPasteHeaders_(sheet, headers);

  ensureColumns_(sheet, headers.length);
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const currentHeaders = getHeaders_(sheet);
    if (!currentHeaders.length || currentHeaders[0] !== "所有者") {
      sheet.clear();
      sheet.clearFormats();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      repairHeaderOrderPreserveData_(sheet, headers);
    }
  }

  formatQualificationEasyPasteSheet_(sheet);
  applyQualificationEasyPasteValidations_(sheet);
  return sheet;
}

function formatQualificationEasyPasteSheet_(sheet) {
  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  try {
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const lastCol = headers.length;

    // v9.7.38:
    // 入力補助シートは「入力のしやすさ」を優先。
    // 資格名まとめだけ折り返し、他列は中央揃えにして見た目のズレを抑える。
    sheet.setFrozenRows(1);
    showAllColumns_(sheet);

    const allRange = sheet.getRange(1, 1, lastRow, lastCol);
    allRange
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment("middle")
      .setWrap(false);

    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setBackground("#d9ead3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(false);

    const centerHeaders = ["所有者", "区分", "追加する資格", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "登録結果"];
    centerHeaders.forEach(header => {
      const col = headers.indexOf(header) + 1;
      if (col <= 0) return;
      sheet.getRange(2, col, Math.max(lastRow - 1, 1), 1)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(false);
    });

    const summaryCol = headers.indexOf("資格名まとめ") + 1;
    if (summaryCol > 0) {
      sheet.getRange(2, summaryCol, Math.max(lastRow - 1, 1), 1)
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle")
        .setWrap(true);
    }

    const noteCol = headers.indexOf("備考") + 1;
    if (noteCol > 0) {
      sheet.getRange(2, noteCol, Math.max(lastRow - 1, 1), 1)
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle")
        .setWrap(true);
    }

    const acquiredCol = headers.indexOf("取得日") + 1;
    const dueCol = headers.indexOf("更新期限") + 1;
    if (acquiredCol > 0) sheet.getRange(2, acquiredCol, Math.max(lastRow - 1, 1), 1).setNumberFormat("yyyy/mm/dd");
    if (dueCol > 0) sheet.getRange(2, dueCol, Math.max(lastRow - 1, 1), 1).setNumberFormat("yyyy/mm/dd");

    setColumnWidthIfExists_(sheet, headers, "所有者", 105);
    setColumnWidthIfExists_(sheet, headers, "区分", 115);
    setColumnWidthIfExists_(sheet, headers, "追加する資格", 285);
    setColumnWidthIfExists_(sheet, headers, "資格名まとめ", 470);
    setColumnWidthIfExists_(sheet, headers, "取得区分", 90);
    setColumnWidthIfExists_(sheet, headers, "保有状況", 90);
    setColumnWidthIfExists_(sheet, headers, "取得日", 100);
    setColumnWidthIfExists_(sheet, headers, "更新期限", 100);
    setColumnWidthIfExists_(sheet, headers, "コピー有無", 95);
    setColumnWidthIfExists_(sheet, headers, "状態", 90);
    setColumnWidthIfExists_(sheet, headers, "備考", 200);
    setColumnWidthIfExists_(sheet, headers, "登録結果", 210);

    sheet.setRowHeight(1, 42);
    if (lastRow >= 2) sheet.setRowHeights(2, Math.max(lastRow - 1, 1), 34);

    createFilterSafely_(sheet, lastCol);
  } catch (e) {
    console.log("formatQualificationEasyPasteSheet_ error: " + e.message);
  }
}

function applyQualificationEasyPasteValidations_(sheet) {
  // v9.7.39: 入力規則は最小修復関数へ集約。
  // 空白行や大量行へ広く設定しないことで、メニュー実行時の重さを抑える。
  repairQualificationEasyPasteColumnValidations_(sheet);
}

function getQualificationEasyPasteValidationRowCount_(sheet) {
  if (!sheet) return 12;
  // 実データ行 + 少しだけ予備。大量の空白行には入力規則を作らない。
  return Math.min(Math.max(sheet.getLastRow() - 1, 12), 25);
}

function repairQualificationEasyPasteColumnValidations_(sheet) {
  if (!sheet || sheet.getName() !== "資格かんたん登録") return;
  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  const rowCount = getQualificationEasyPasteValidationRowCount_(sheet);
  const lastCol = headers.length;
  if (rowCount <= 0 || lastCol <= 0) return;

  // 列ズレした入力規則を小さい範囲だけクリアしてから、必要列だけ再設定する。
  sheet.getRange(2, 1, rowCount, lastCol).clearDataValidations();

  const staffList = [CLEAR_LABEL, ...getStaffMembers_()];
  const categoryList = [CLEAR_LABEL, "資格", "技能講習", "特別教育", "安全教育", "免許", "その他"];
  const possessionList = [CLEAR_LABEL, "技士", "技士補", "保有", "修了", "免許"];
  const holdingStatusList = [CLEAR_LABEL, "保有", "未確認", "失効", "不要"];
  const copyList = [CLEAR_LABEL, "有", "無", "未確認"];
  const statusList = [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"];
  setDropdownByHeaderIfExists_(sheet, headers, "所有者", staffList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "区分", categoryList, rowCount);
  // v9.7.47: 「追加する資格」は区分ごとの行別候補だけを設定する。
  // 先に全候補を設定すると、途中で処理が止まった時に区分と候補がズレるため避ける。
  rebuildQualificationCandidateValidationsForRows_(sheet, 2, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "取得区分", possessionList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "保有状況", holdingStatusList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "コピー有無", copyList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "状態", statusList, rowCount);

  const summaryCol = headers.indexOf("資格名まとめ") + 1;
  const resultCol = headers.indexOf("登録結果") + 1;
  const acquiredCol = headers.indexOf("取得日") + 1;
  const dueCol = headers.indexOf("更新期限") + 1;

  if (summaryCol > 0) {
    sheet.getRange(2, summaryCol, rowCount, 1)
      .clearDataValidations()
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle")
      .setWrap(true);
  }
  if (resultCol > 0) sheet.getRange(2, resultCol, rowCount, 1).clearDataValidations();
  if (acquiredCol > 0) sheet.getRange(2, acquiredCol, rowCount, 1).setNumberFormat("yyyy/mm/dd");
  if (dueCol > 0) sheet.getRange(2, dueCol, rowCount, 1).setNumberFormat("yyyy/mm/dd");
}


/**
 * 資格かんたん登録：
 * 「追加する資格」プルダウンを選んだら、同じ行の「資格名まとめ」へ追記する。
 * 正式登録は行わず、資格管理への反映はメニュー実行に限定する。
 */
function handleQualificationEasyPasteAppend_(e) {
  if (!e || !e.range) return false;

  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "資格かんたん登録") return false;
  if (e.range.getRow() <= 1) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  if (editedHeader === "区分") {
    const targetRow = e.range.getRow();
    syncQualificationEasyPasteDropdownForRow_(sheet, targetRow, true);
    const resultCol = headers.indexOf("登録結果") + 1;
    const categoryCol = headers.indexOf("区分") + 1;
    const category = categoryCol > 0 ? normalizeQualificationCategoryValue_(sheet.getRange(targetRow, categoryCol).getValue()) : "";
    if (resultCol > 0) {
      sheet.getRange(targetRow, resultCol).setValue(
        category ? "区分「" + category + "」に合わせて資格プルダウンを更新しました" : "区分が空欄のため、全資格候補に戻しました"
      );
    }
    return true;
  }
  if (editedHeader !== "追加する資格") return false;

  const selectedRaw = String(e.value || e.range.getValue() || "").trim();
  if (!selectedRaw || selectedRaw === CLEAR_LABEL) {
    e.range.clearContent();
    return true;
  }

  const row = e.range.getRow();
  const summaryCol = headers.indexOf("資格名まとめ") + 1;
  const categoryCol = headers.indexOf("区分") + 1;
  const possessionCol = headers.indexOf("取得区分") + 1;
  const resultCol = headers.indexOf("登録結果") + 1;
  if (summaryCol <= 0) return false;

  const parsed = parseQualificationSelection_(selectedRaw);
  const selected = parsed.name;
  const selectedCategory = normalizeQualificationCategoryValue_(parsed.category || getQualificationMasterCategoryByName_(selected) || "");
  const selectedAcquisition = parsed.acquisition || "";

  if (!selected) {
    e.range.clearContent();
    if (resultCol > 0) sheet.getRange(row, resultCol).setValue("要確認：資格名を読み取れません");
    return true;
  }

  let messageParts = [];

  // 1行は「1人 + 1区分 + 複数資格」。区分違いを混ぜると正式データが崩れるため止める。
  if (categoryCol > 0 && selectedCategory) {
    const categoryCell = sheet.getRange(row, categoryCol);
    const currentCategory = normalizeQualificationCategoryValue_(categoryCell.getValue());
    if (!currentCategory) {
      categoryCell.setValue(selectedCategory);
    } else if (currentCategory !== selectedCategory) {
      e.range.clearContent();
      if (resultCol > 0) {
        sheet.getRange(row, resultCol).setValue(
          "区分違い：この行は「" + currentCategory + "」。選択は「" + selectedCategory + "」。同じ所有者で別行に入力してください"
        );
      }
      return true;
    }
  }

  // 取得区分も行全体へかかる。
  // 技士/技士補は施工管理技士系だけ許可。それ以外の資格は取得区分を分ける必要がある。
  if (possessionCol > 0) {
    const possessionCell = sheet.getRange(row, possessionCol);
    const currentPossession = String(possessionCell.getValue() || "").trim();
    const desiredAcquisition = selectedAcquisition || currentPossession;
    const normalized = normalizeQualificationPossessionForItem_(selectedCategory, selected, desiredAcquisition);

    if (!currentPossession || currentPossession === CLEAR_LABEL) {
      possessionCell.setValue(normalized.value);
      if (normalized.message) messageParts.push(normalized.message);
    } else if (selectedAcquisition && currentPossession !== selectedAcquisition && currentSummaryHasValue_(sheet, row, summaryCol)) {
      e.range.clearContent();
      if (resultCol > 0) {
        sheet.getRange(row, resultCol).setValue("取得区分違い：この行は「" + currentPossession + "」。選択は「" + selectedAcquisition + "」。別行に入力してください");
      }
      return true;
    } else if (normalized.blockAppend) {
      e.range.clearContent();
      if (resultCol > 0) {
        sheet.getRange(row, resultCol).setValue(
          normalized.message + "。同じ所有者で取得区分を分けて別行に入力してください"
        );
      }
      return true;
    } else if (normalized.value !== currentPossession) {
      possessionCell.setValue(normalized.value);
      if (normalized.message) messageParts.push(normalized.message);
    }
  }

  const holdingStatusCol = headers.indexOf("保有状況") + 1;
  if (holdingStatusCol > 0) {
    const holdingStatusCell = sheet.getRange(row, holdingStatusCol);
    const currentHoldingStatus = String(holdingStatusCell.getValue() || "").trim();
    if (!currentHoldingStatus || currentHoldingStatus === CLEAR_LABEL) holdingStatusCell.setValue("保有");
  }

  const summaryCell = sheet.getRange(row, summaryCol);
  // 旧版の入力規則が残っていても自動追記できるよう、資格名まとめセルは自由入力に戻す。
  summaryCell.clearDataValidations();
  const currentSummary = String(summaryCell.getValue() || "").trim();
  const existingNames = splitQualificationNames_(currentSummary);
  const existingKeys = {};
  existingNames.forEach(name => existingKeys[normalizeTextForKey_(parseQualificationSelection_(name).name)] = true);

  const selectedKey = normalizeTextForKey_(selected);
  if (existingKeys[selectedKey]) {
    messageParts.push("追加済み：" + selected);
  } else {
    const nextNames = existingNames.map(name => parseQualificationSelection_(name).name).filter(Boolean).concat([selected]);
    const sortedNames = sortQualificationNameListByMasterOrder_(nextNames, selectedCategory || "");
    summaryCell.setValue(sortedNames.join("\n"));
    summaryCell.setWrap(true);
    messageParts.push("追記：" + (selectedCategory ? selectedCategory + "｜" : "") + selected);
  }

  e.range.clearContent();
  if (resultCol > 0) sheet.getRange(row, resultCol).setValue(messageParts.join(" / "));
  return true;
}



function normalizeAndSortQualificationEasyPasteSheet_(sheet, sortRows) {
  const result = {summarySorted: 0, blankRows: 0};
  if (!sheet || sheet.getName() !== "資格かんたん登録") return result;

  migrateQualificationEasyPasteHeaders_(sheet, getSheetHeaders_()["資格かんたん登録"]);
  const headers = getHeaders_(sheet);
  if (!headers.length) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    repairQualificationEasyPasteColumnValidations_(sheet);
    return result;
  }

  const lastCol = headers.length;
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const summaryIndex = headers.indexOf("資格名まとめ");
  const categoryIndex = headers.indexOf("区分");
  const addIndex = headers.indexOf("追加する資格");
  const acquisitionIndex = headers.indexOf("取得区分");
  const holdingIndex = headers.indexOf("保有状況");
  const resultIndex = headers.indexOf("登録結果");

  const masterOrderMap = getQualificationMasterOrderMap_();
  const staffOrder = getStaffOrderMap_();
  const dataRows = [];
  const blankRows = [];

  values.forEach(row => {
    // 追加する資格は選択後に空欄へ戻す補助列なので、残っていたら整理時に消す。
    if (addIndex >= 0) row[addIndex] = "";

    const category = categoryIndex >= 0 ? String(row[categoryIndex] || "").trim() : "";
    if (summaryIndex >= 0) {
      const before = String(row[summaryIndex] || "").trim();
      const sortedNames = sortQualificationNameListByMasterOrder_(splitQualificationNames_(before), category, masterOrderMap);
      const after = sortedNames.join("\n");
      if (before !== after) result.summarySorted++;
      row[summaryIndex] = after;
    }

    if (acquisitionIndex >= 0 && String(row[acquisitionIndex] || "").trim() === CLEAR_LABEL) row[acquisitionIndex] = "";
    if (holdingIndex >= 0) {
      const summary = summaryIndex >= 0 ? String(row[summaryIndex] || "").trim() : "";
      const current = String(row[holdingIndex] || "").trim();
      if (summary && (!current || current === CLEAR_LABEL)) row[holdingIndex] = "保有";
      if (!summary && current === CLEAR_LABEL) row[holdingIndex] = "";
    }
    if (resultIndex >= 0 && isQualificationEasyPasteRowBlank_(row, headers)) row[resultIndex] = "";

    if (isQualificationEasyPasteRowBlank_(row, headers)) {
      blankRows.push(new Array(lastCol).fill(""));
    } else {
      dataRows.push(row);
    }
  });

  if (sortRows) {
    dataRows.sort((a, b) => compareQualificationEasyPasteRows_(a, b, headers, masterOrderMap, staffOrder));
  }

  result.blankRows = blankRows.length;
  const output = dataRows.concat(blankRows);
  if (output.length) sheet.getRange(2, 1, output.length, lastCol).setValues(output);

  formatQualificationEasyPasteSheet_(sheet);
  repairQualificationEasyPasteColumnValidations_(sheet);
  return result;
}

function isQualificationEasyPasteRowBlank_(row, headers) {
  const ignore = {"追加する資格": true, "登録結果": true};
  return headers.every((header, index) => {
    if (ignore[header]) return true;
    return String(row[index] || "").trim() === "";
  });
}

function compareQualificationEasyPasteRows_(a, b, headers, masterOrderMap, staffOrder) {
  const ownerIndex = headers.indexOf("所有者");
  const categoryIndex = headers.indexOf("区分");
  const summaryIndex = headers.indexOf("資格名まとめ");
  const acquisitionIndex = headers.indexOf("取得区分");
  const dueIndex = headers.indexOf("更新期限");

  const ownerA = ownerIndex >= 0 ? String(a[ownerIndex] || "").trim() : "";
  const ownerB = ownerIndex >= 0 ? String(b[ownerIndex] || "").trim() : "";
  const ownerOrderA = getStaffSortOrder_(ownerA, staffOrder);
  const ownerOrderB = getStaffSortOrder_(ownerB, staffOrder);
  if (ownerOrderA !== ownerOrderB) return ownerOrderA - ownerOrderB;
  const ownerNameCompare = ownerA.localeCompare(ownerB, "ja");
  if (ownerNameCompare !== 0) return ownerNameCompare;

  const categoryA = categoryIndex >= 0 ? String(a[categoryIndex] || "").trim() : "";
  const categoryB = categoryIndex >= 0 ? String(b[categoryIndex] || "").trim() : "";
  const categoryOrderA = getQualificationCategoryOrder_(categoryA);
  const categoryOrderB = getQualificationCategoryOrder_(categoryB);
  if (categoryOrderA !== categoryOrderB) return categoryOrderA - categoryOrderB;

  const minOrderA = getMinQualificationMasterOrderFromSummary_(summaryIndex >= 0 ? a[summaryIndex] : "", categoryA, masterOrderMap);
  const minOrderB = getMinQualificationMasterOrderFromSummary_(summaryIndex >= 0 ? b[summaryIndex] : "", categoryB, masterOrderMap);
  if (minOrderA !== minOrderB) return minOrderA - minOrderB;

  const acquisitionA = acquisitionIndex >= 0 ? String(a[acquisitionIndex] || "").trim() : "";
  const acquisitionB = acquisitionIndex >= 0 ? String(b[acquisitionIndex] || "").trim() : "";
  const acquisitionOrderA = getQualificationAcquisitionSortOrder_(acquisitionA);
  const acquisitionOrderB = getQualificationAcquisitionSortOrder_(acquisitionB);
  if (acquisitionOrderA !== acquisitionOrderB) return acquisitionOrderA - acquisitionOrderB;

  const dueA = dueIndex >= 0 ? getDateTimeSortValue_(a[dueIndex]) : 9999999999999;
  const dueB = dueIndex >= 0 ? getDateTimeSortValue_(b[dueIndex]) : 9999999999999;
  if (dueA !== dueB) return dueA - dueB;

  const summaryA = summaryIndex >= 0 ? String(a[summaryIndex] || "") : "";
  const summaryB = summaryIndex >= 0 ? String(b[summaryIndex] || "") : "";
  return summaryA.localeCompare(summaryB, "ja");
}

function getMinQualificationMasterOrderFromSummary_(summary, category, masterOrderMap) {
  const names = splitQualificationNames_(summary);
  if (!names.length) return 999999;
  return names.reduce((min, rawName) => {
    const parsed = parseQualificationSelection_(rawName);
    const name = parsed.name;
    const itemCategory = parsed.category || category || getQualificationMasterCategoryByName_(name) || "";
    const order = getQualificationMasterSortOrder_(itemCategory, name, masterOrderMap);
    return Math.min(min, order);
  }, 999999);
}

function sortQualificationNameListByMasterOrder_(names, categoryHint, masterOrderMap) {
  const map = masterOrderMap || getQualificationMasterOrderMap_();
  const seen = {};
  const rows = [];
  (names || []).forEach((rawName, index) => {
    const parsed = parseQualificationSelection_(rawName);
    const name = cleanQualificationName_(parsed.name || rawName);
    if (!name) return;
    const key = normalizeTextForKey_(name);
    if (seen[key]) return;
    seen[key] = true;
    const category = parsed.category || categoryHint || getQualificationMasterCategoryByName_(name) || "";
    rows.push({
      name: name,
      order: getQualificationMasterSortOrder_(category, name, map),
      categoryOrder: getQualificationCategoryOrder_(category),
      originalIndex: index
    });
  });

  rows.sort((a, b) => {
    if (a.categoryOrder !== b.categoryOrder) return a.categoryOrder - b.categoryOrder;
    if (a.order !== b.order) return a.order - b.order;
    const nameCompare = a.name.localeCompare(b.name, "ja");
    if (nameCompare !== 0) return nameCompare;
    return a.originalIndex - b.originalIndex;
  });
  return rows.map(row => row.name);
}

function getQualificationAcquisitionSortOrder_(value) {
  const text = String(value || "").trim();
  const order = {"技士": 1, "技士補": 2, "保有": 3, "修了": 4, "免許": 5};
  return order[text] || 99;
}

function getStaffOrderMap_() {
  const map = {};
  getStaffMembers_().forEach((name, index) => {
    const key = normalizeTextForKey_(name);
    if (key && map[key] === undefined) map[key] = index + 1;
  });
  return map;
}

function getStaffSortOrder_(name, staffOrder) {
  const key = normalizeTextForKey_(name);
  if (!key) return 999999;
  const map = staffOrder || getStaffOrderMap_();
  return map[key] || 900000;
}

function currentSummaryHasValue_(sheet, row, summaryCol) {
  if (!sheet || !row || !summaryCol) return false;
  return String(sheet.getRange(row, summaryCol).getValue() || "").trim() !== "";
}

function isEngineerPossessionValue_(value) {
  return ["技士", "技士補"].includes(String(value || "").trim());
}

function isEngineerPossessionSupportedQualification_(qualificationName) {
  const name = String(qualificationName || "").trim();
  if (!name) return false;
  // 紙の資格表の ●技士 / ○技士補 に相当する主対象。
  // 監理技術者、測量士、石綿調査者講習、配管技能士、経理事務士などは基本「保有」。
  return /施工管理技士/.test(name);
}

function normalizeQualificationPossessionForItem_(category, qualificationName, possessionValue) {
  const categoryText = String(category || "").trim();
  const raw = String(possessionValue || "").trim();
  const defaultValue = getDefaultQualificationAcquisitionByCategory_(categoryText);
  const value = (!raw || raw === CLEAR_LABEL) ? defaultValue : raw;

  if (!isQualificationAcquisitionValue_(value)) {
    return {value: defaultValue, message: "取得区分補正：不明な値は" + defaultValue + "扱い", blockAppend: false};
  }

  if (categoryText && categoryText !== "資格" && isEngineerPossessionValue_(value)) {
    return {value: getDefaultQualificationAcquisitionByCategory_(categoryText), message: "取得区分補正：" + categoryText + "は技士/技士補ではありません", blockAppend: false};
  }

  if (categoryText === "資格" && isEngineerPossessionValue_(value) && !isEngineerPossessionSupportedQualification_(qualificationName)) {
    return {
      value: value,
      message: "取得区分違い：「" + qualificationName + "」は技士/技士補ではなく保有で別行に入力してください",
      blockAppend: true
    };
  }

  return {value: value, message: "", blockAppend: false};
}

function getDefaultQualificationAcquisitionByCategory_(category) {
  const categoryText = String(category || "").trim();
  if (["技能講習", "特別教育", "安全教育"].includes(categoryText)) return "修了";
  if (categoryText === "免許") return "免許";
  return "保有";
}


/**
 * v9.7.42:
 * 資格かんたん登録の1行は「1人 + 1区分 + 1取得区分 + 複数資格」。
 * 技士/技士補の行に、測量士・石綿調査者講習などの対象外資格が混ざっている場合、
 * 途中まで更新せず、行全体を要確認にする。
 */
function validateQualificationEasyPasteRowRule_(category, names, possessionValue, masterMap) {
  const rowCategory = String(category || "").trim();
  const possession = String(possessionValue || "").trim();
  const issues = [];
  const items = [];

  (names || []).forEach(rawName => {
    const parsed = parseQualificationSelection_(rawName);
    const name = parsed.name;
    if (!name) return;
    const masterCategory = parsed.category || (masterMap || {})[normalizeTextForKey_(name)] || "その他";
    const itemCategory = rowCategory || masterCategory;
    items.push({name: name, masterCategory: masterCategory, itemCategory: itemCategory});
  });

  if (!items.length) return {ok: false, message: "要確認：資格名まとめが空欄"};

  // 安全優先：技士/技士補は誤更新が起きやすいため、1行1資格に限定する。
  // 例：測量士を技士補で更新しようとして、同じ行の施工管理技士まで技士補になる事故を防ぐ。
  if (isEngineerPossessionValue_(possession) && items.length > 1) {
    issues.push("取得区分安全制御：技士/技士補は1行1資格で入力してください。保有の資格とは別行に分けてください");
  }

  const categorySet = {};
  items.forEach(item => {
    if (item.masterCategory) categorySet[item.masterCategory] = true;
  });
  const categories = Object.keys(categorySet);

  if (!rowCategory && categories.length > 1) {
    issues.push("区分混在：" + categories.join(" / ") + "。区分ごとに行を分けてください");
  }

  if (rowCategory) {
    const mismatched = items
      .filter(item => item.masterCategory && item.masterCategory !== rowCategory)
      .map(item => item.name + "（" + item.masterCategory + "）");
    if (mismatched.length) {
      issues.push("区分違い：" + mismatched.join("、") + "。別行に分けてください");
    }
  }

  if (isEngineerPossessionValue_(possession)) {
    if (rowCategory && rowCategory !== "資格") {
      issues.push("取得区分違い：" + rowCategory + "は技士/技士補ではなく保有にしてください");
    }

    const unsupported = items
      .filter(item => item.itemCategory !== "資格" || !isEngineerPossessionSupportedQualification_(item.name))
      .map(item => item.name);
    if (unsupported.length) {
      issues.push("取得区分混在：技士/技士補にできない資格があります（" + unsupported.join("、") + "）。保有で別行に分けてください");
    }
  }

  return {ok: issues.length === 0, message: issues.join(" / ")};
}

function parseQualificationSelection_(value) {
  const raw = String(value || "").trim();
  if (!raw) return {category: "", name: "", acquisition: ""};

  const delimiter = raw.indexOf("｜") >= 0 ? "｜" : (raw.indexOf("|") >= 0 ? "|" : "");
  if (delimiter) {
    const parts = raw.split(delimiter).map(v => String(v || "").trim()).filter(Boolean);
    const category = parts.shift() || "";
    let acquisition = "";
    if (parts.length >= 2) {
      const last = parts[parts.length - 1];
      if (isQualificationAcquisitionValue_(last)) {
        acquisition = parts.pop();
      }
    }
    const name = cleanQualificationName_(parts.join(delimiter));
    return {category: category, name: name, acquisition: acquisition};
  }

  const name = cleanQualificationName_(raw);
  return {category: getQualificationMasterCategoryByName_(name) || "", name: name, acquisition: ""};
}

function isQualificationAcquisitionValue_(value) {
  return ["技士", "技士補", "保有", "修了", "免許"].includes(String(value || "").trim());
}

function getQualificationMasterDisplayList_() {
  return getQualificationMasterDisplayListByCategory_("");
}

function getQualificationMasterDisplayListByCategory_(categoryFilter) {
  const categoryText = normalizeQualificationCategoryValue_(categoryFilter);
  const out = [];
  getQualificationMasterItems_().forEach(item => {
    const category = normalizeQualificationCategoryValue_(item.category) || "その他";
    if (categoryText && category !== categoryText) return;
    if (category === "資格" && isEngineerPossessionSupportedQualification_(item.name)) {
      out.push(category + "｜" + item.name + "｜技士");
      out.push(category + "｜" + item.name + "｜技士補");
    } else {
      out.push(category + "｜" + item.name);
    }
  });
  return out;
}

function normalizeQualificationCategoryValue_(value) {
  const text = String(value || "")
    .replace(/[|｜].*$/g, "")
    .replace(/\u3000/g, " ")
    .trim();
  if (!text || text === CLEAR_LABEL) return "";
  const aliases = {
    "資格": "資格",
    "免許": "免許",
    "技能講習": "技能講習",
    "技能": "技能講習",
    "特別教育": "特別教育",
    "特教": "特別教育",
    "安全教育": "安全教育",
    "安全": "安全教育",
    "その他": "その他"
  };
  return aliases[text] || text;
}

function setQualificationCandidateValidationForRow_(sheet, row) {
  syncQualificationEasyPasteDropdownForRow_(sheet, row, false);
}

function syncQualificationEasyPasteDropdownForRow_(sheet, row, clearSelectedValue) {
  if (!sheet || sheet.getName() !== "資格かんたん登録" || row <= 1) return;
  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const addCol = headers.indexOf("追加する資格") + 1;
  if (addCol <= 0) return;

  const category = categoryCol > 0 ? normalizeQualificationCategoryValue_(sheet.getRange(row, categoryCol).getValue()) : "";
  const list = [CLEAR_LABEL].concat(getQualificationMasterDisplayListByCategory_(category));

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();

  const addCell = sheet.getRange(row, addCol);
  addCell.clearDataValidations();
  addCell.setDataValidation(rule);
  if (clearSelectedValue) addCell.clearContent();
  SpreadsheetApp.flush();
}

function rebuildQualificationCandidateValidationsForRows_(sheet, startRow, rowCount) {
  if (!sheet || sheet.getName() !== "資格かんたん登録") return;
  const headers = getHeaders_(sheet);
  const addCol = headers.indexOf("追加する資格") + 1;
  if (addCol <= 0 || rowCount <= 0) return;
  sheet.getRange(startRow, addCol, rowCount, 1).clearDataValidations();
  for (let r = startRow; r < startRow + rowCount; r++) {
    syncQualificationEasyPasteDropdownForRow_(sheet, r, false);
  }
}

function rebuildQualificationEasyPasteDependentDropdowns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("資格かんたん登録") || ensureQualificationEasyPasteSheet_();
  migrateQualificationEasyPasteHeaders_(sheet, getSheetHeaders_()["資格かんたん登録"]);
  const rowCount = getQualificationEasyPasteValidationRowCount_(sheet);
  repairQualificationEasyPasteColumnValidations_(sheet);
  rebuildQualificationCandidateValidationsForRows_(sheet, 2, rowCount);
  toast_("資格かんたん登録の区分別プルダウンを再構築しました（" + rowCount + "行）");
}

function getQualificationMasterCategoryByName_(name) {
  const key = normalizeTextForKey_(cleanQualificationName_(name));
  if (!key) return "";
  const map = getQualificationMasterCategoryMap_();
  return map[key] || "";
}


function getQualificationMasterNameList_() {
  const seen = {};
  const names = [];
  getQualificationMasterItems_().forEach(item => {
    const name = cleanQualificationName_(item.name);
    const key = normalizeTextForKey_(name);
    if (!name || seen[key]) return;
    seen[key] = true;
    names.push(name);
  });
  return names;
}

function cleanQualificationName_(value) {
  let name = String(value || "").trim();
  name = name.replace(/^[・･●○◎□■\-－ー\s]+/, "").trim();
  name = name.replace(/\s+/g, " ").trim();
  return name;
}

function applyQualificationEasyPasteToManagement_() {
  const importSheet = ensureQualificationEasyPasteSheet_();
  const targetSheet = ensureQualificationManagementSheet_();
  const importHeaders = getHeaders_(importSheet);
  const targetHeaders = getHeaders_(targetSheet);
  const resultCol = importHeaders.indexOf("登録結果") + 1;
  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) return {added: 0, skipped: 0, invalid: 0};

  const importValues = importSheet.getRange(2, 1, lastRow - 1, importHeaders.length).getValues();
  const masterMap = getQualificationMasterCategoryMap_();
  const existingKeys = getExistingQualificationKeys_(targetSheet, targetHeaders);
  const appendRows = [];
  const resultValues = [];
  let added = 0;
  let skipped = 0;
  let invalid = 0;
  let corrected = 0;

  importValues.forEach(row => {
    const obj = objectFromRow(importHeaders, row);
    const owner = String(obj["所有者"] || "").trim();
    const rawNames = obj["資格名まとめ"];
    let category = String(obj["区分"] || "").trim();
    let possession = String(obj["取得区分"] || "").trim();
    let status = String(obj["状態"] || "").trim();
    const acquired = obj["取得日"] || "";
    const due = obj["更新期限"] || "";
    const copy = String(obj["コピー有無"] || "").trim();
    const holdingStatus = String(obj["保有状況"] || "").trim();
    const note = obj["備考"] || "";

    const names = sortQualificationNameListByMasterOrder_(splitQualificationNames_(rawNames), category);

    if (!owner && !category && !String(rawNames || "").trim() && !possession && !acquired && !due && !copy && !status && !note) {
      resultValues.push([""]);
      return;
    }

    if (!owner || !names.length) {
      invalid++;
      resultValues.push(["要確認：所有者または資格名まとめが空欄"]);
      return;
    }

    if (!possession || possession === CLEAR_LABEL) possession = "";
    if (!status || status === CLEAR_LABEL) status = "有効";

    const rowRule = validateQualificationEasyPasteRowRule_(category, names, possession, masterMap);
    if (!rowRule.ok) {
      invalid++;
      resultValues.push([rowRule.message]);
      return;
    }

    let rowAdded = 0;
    let rowSkipped = 0;
    let rowInvalid = 0;
    let rowCorrected = 0;

    names.forEach(rawQualification => {
      const parsed = parseQualificationSelection_(rawQualification);
      const qualification = parsed.name;
      const masterCategory = parsed.category || masterMap[normalizeTextForKey_(qualification)] || "その他";
      const itemCategory = category || masterCategory;
      if (!qualification) {
        rowInvalid++;
        invalid++;
        return;
      }
      if (category && masterCategory && category !== masterCategory) {
        rowInvalid++;
        invalid++;
        return;
      }

      const possessionInfo = normalizeQualificationPossessionForItem_(itemCategory, qualification, possession);
      if (possessionInfo.blockAppend) {
        rowInvalid++;
        invalid++;
        return;
      }
      if (possessionInfo.value !== possession && possessionInfo.message) {
        rowCorrected++;
        corrected++;
      }
      const itemPossession = possessionInfo.value || "保有";

      const key = buildQualificationDuplicateKey_(owner, itemCategory, qualification);
      if (existingKeys[key]) {
        rowSkipped++;
        skipped++;
        return;
      }

      const rowObj = {
        "所有者": owner,
        "区分": itemCategory,
        "資格名": qualification,
        "取得区分": itemPossession,
        "保有状況": holdingStatus || "保有",
        "取得日": acquired,
        "更新期限": due,
        "コピー有無": copy || "未確認",
        "状態": status,
        "通知": getNoticeText(due, status),
        "備考": note,
        "資格ID": buildRecordId_("資格管理")
      };

      appendRows.push(targetHeaders.map(header => {
        if (header === READ_HEADER || getPersonalMembers_().includes(header)) return false;
        return rowObj[header] !== undefined ? rowObj[header] : "";
      }));
      existingKeys[key] = true;
      rowAdded++;
      added++;
    });

    const parts = [];
    if (rowAdded) parts.push("追加" + rowAdded + "件");
    if (rowSkipped) parts.push("重複" + rowSkipped + "件");
    if (rowCorrected) parts.push("取得区分補正" + rowCorrected + "件");
    if (rowInvalid) parts.push("要確認" + rowInvalid + "件");
    resultValues.push([parts.length ? parts.join(" / ") : "追加なし"]);
  });

  if (appendRows.length) {
    const startRow = Math.max(targetSheet.getLastRow() + 1, 2);
    targetSheet.getRange(startRow, 1, appendRows.length, targetHeaders.length).setValues(appendRows);
    try {
      setCheckboxesForDataRows(targetSheet);
      applyQualificationSingleSheetLightSettings_(targetSheet);
      applyQualificationManagementCompactDisplay_(targetSheet);
    } catch (e) {}
  }

  if (resultCol > 0 && resultValues.length) {
    importSheet.getRange(2, resultCol, resultValues.length, 1).setValues(resultValues);
  }

  return {added: added, skipped: skipped, invalid: invalid, corrected: corrected};
}


function applyQualificationEasyPasteToManagementUpdateExisting_(forceOverwrite) {
  const importSheet = ensureQualificationEasyPasteSheet_();
  const targetSheet = ensureQualificationManagementSheet_();
  const importHeaders = getHeaders_(importSheet);
  const targetHeaders = getHeaders_(targetSheet);
  const resultCol = importHeaders.indexOf("登録結果") + 1;
  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) return {added: 0, updated: 0, unchanged: 0, invalid: 0, corrected: 0};

  const importValues = importSheet.getRange(2, 1, lastRow - 1, importHeaders.length).getValues();
  const masterMap = getQualificationMasterCategoryMap_();
  const existingMap = getExistingQualificationRowMap_(targetSheet, targetHeaders);
  const appendRows = [];
  const resultValues = [];

  let added = 0;
  let updated = 0;
  let unchanged = 0;
  let invalid = 0;
  let corrected = 0;

  importValues.forEach(row => {
    const obj = objectFromRow(importHeaders, row);
    const owner = String(obj["所有者"] || "").trim();
    const rawNames = obj["資格名まとめ"];
    let category = String(obj["区分"] || "").trim();
    const rawPossession = String(obj["取得区分"] || "").trim();
    const rawStatus = String(obj["状態"] || "").trim();
    const holdingStatus = String(obj["保有状況"] || "").trim();
    const acquired = obj["取得日"] || "";
    const due = obj["更新期限"] || "";
    const copy = String(obj["コピー有無"] || "").trim();
    const note = obj["備考"] || "";
    const names = sortQualificationNameListByMasterOrder_(splitQualificationNames_(rawNames), category);

    if (!owner && !category && !String(rawNames || "").trim() && !rawPossession && !acquired && !due && !copy && !rawStatus && !note) {
      resultValues.push([""]);
      return;
    }

    if (!owner || !names.length) {
      invalid++;
      resultValues.push(["要確認：所有者または資格名まとめが空欄"]);
      return;
    }

    const possessionForRowCheck = (rawPossession && rawPossession !== CLEAR_LABEL) ? rawPossession : "";
    const rowRule = validateQualificationEasyPasteRowRule_(category, names, possessionForRowCheck, masterMap);
    if (!rowRule.ok) {
      invalid++;
      resultValues.push([rowRule.message]);
      return;
    }

    let rowAdded = 0;
    let rowUpdated = 0;
    let rowUnchanged = 0;
    let rowInvalid = 0;
    let rowCorrected = 0;

    names.forEach(rawQualification => {
      const parsed = parseQualificationSelection_(rawQualification);
      const qualification = parsed.name;
      const masterCategory = parsed.category || masterMap[normalizeTextForKey_(qualification)] || "その他";
      const itemCategory = category || masterCategory;

      if (!qualification) {
        rowInvalid++;
        invalid++;
        return;
      }
      if (category && masterCategory && category !== masterCategory) {
        rowInvalid++;
        invalid++;
        return;
      }

      const hasPossessionInput = !!rawPossession && rawPossession !== CLEAR_LABEL;
      const possessionForNew = hasPossessionInput ? rawPossession : "";
      const possessionInfo = normalizeQualificationPossessionForItem_(itemCategory, qualification, possessionForNew);
      if (possessionInfo.blockAppend) {
        rowInvalid++;
        invalid++;
        return;
      }
      if (possessionInfo.value !== possessionForNew && possessionInfo.message) {
        rowCorrected++;
        corrected++;
      }

      const statusForNew = (rawStatus && rawStatus !== CLEAR_LABEL) ? rawStatus : "有効";
      const key = buildQualificationDuplicateKey_(owner, itemCategory, qualification);
      const existing = existingMap[key];

      const rowObj = {
        "所有者": owner,
        "区分": itemCategory,
        "資格名": qualification,
        "取得区分": possessionInfo.value || "保有",
        "保有状況": holdingStatus || "保有",
        "取得日": acquired,
        "更新期限": due,
        "コピー有無": copy || "未確認",
        "状態": statusForNew,
        "通知": getNoticeText(due, statusForNew),
        "備考": note,
        "資格ID": buildRecordId_("資格管理")
      };

      if (!existing) {
        appendRows.push(targetHeaders.map(header => {
          if (header === READ_HEADER || getPersonalMembers_().includes(header)) return false;
          return rowObj[header] !== undefined ? rowObj[header] : "";
        }));
        existingMap[key] = {rowNumber: -1, rowValues: rowObj};
        rowAdded++;
        added++;
        return;
      }

      const updateResult = updateExistingQualificationRow_(
        targetSheet,
        targetHeaders,
        existing.rowNumber,
        rowObj,
        {
          forceOverwrite: !!forceOverwrite,
          hasPossession: hasPossessionInput,
          hasHoldingStatus: !!holdingStatus && holdingStatus !== CLEAR_LABEL,
          hasAcquired: acquired !== "",
          hasDue: due !== "",
          hasCopy: !!copy && copy !== CLEAR_LABEL,
          hasStatus: !!rawStatus && rawStatus !== CLEAR_LABEL,
          hasNote: note !== ""
        }
      );

      if (updateResult.updated) {
        rowUpdated++;
        updated++;
      } else {
        rowUnchanged++;
        unchanged++;
      }
    });

    const parts = [];
    if (rowAdded) parts.push("追加" + rowAdded + "件");
    if (rowUpdated) parts.push((forceOverwrite ? "上書き" : "更新") + rowUpdated + "件");
    if (rowUnchanged) parts.push("変更なし" + rowUnchanged + "件");
    if (rowCorrected) parts.push("取得区分補正" + rowCorrected + "件");
    if (rowInvalid) parts.push("要確認" + rowInvalid + "件");
    resultValues.push([parts.length ? parts.join(" / ") : "反映なし"]);
  });

  if (appendRows.length) {
    const startRow = Math.max(targetSheet.getLastRow() + 1, 2);
    targetSheet.getRange(startRow, 1, appendRows.length, targetHeaders.length).setValues(appendRows);
  }

  try {
    setCheckboxesForDataRows(targetSheet);
    applyQualificationSingleSheetLightSettings_(targetSheet);
    applyQualificationManagementCompactDisplay_(targetSheet);
  } catch (e) {}

  if (resultCol > 0 && resultValues.length) {
    importSheet.getRange(2, resultCol, resultValues.length, 1).setValues(resultValues);
  }

  return {added: added, updated: updated, unchanged: unchanged, invalid: invalid, corrected: corrected};
}


function repairQualificationPossessionSafety() {
  try {
    const result = repairQualificationPossessionSafety_();
    toast_("資格管理の取得区分を安全補正しました。補正 " + result.fixed + "件 / 確認 " + result.checked + "件");
  } catch (e) {
    SpreadsheetApp.getUi().alert("資格管理の取得区分安全補正でエラー：" + e.message);
  }
}

function repairQualificationPossessionSafety_() {
  const sheet = ensureQualificationManagementSheet_();
  const headers = getHeaders_(sheet);
  if (!sheet || sheet.getLastRow() < 2) return {fixed: 0, checked: 0};

  const categoryIndex = headers.indexOf("区分");
  const nameIndex = headers.indexOf("資格名");
  const possessionIndex = headers.indexOf("取得区分");
  if (categoryIndex < 0 || nameIndex < 0 || possessionIndex < 0) return {fixed: 0, checked: 0};

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length);
  const values = range.getValues();
  let fixed = 0;
  let checked = 0;

  values.forEach(row => {
    const category = String(row[categoryIndex] || "").trim();
    const qualification = String(row[nameIndex] || "").trim();
    const possession = String(row[possessionIndex] || "").trim();
    if (!qualification || !isEngineerPossessionValue_(possession)) return;
    checked++;

    // 技能講習・特別教育・安全教育・免許、または施工管理技士系以外は「保有」に戻す。
    if (category !== "資格" || !isEngineerPossessionSupportedQualification_(qualification)) {
      row[possessionIndex] = "保有";
      fixed++;
    }
  });

  if (fixed > 0) {
    range.setValues(values);
    try {
      applyQualificationManagementCompactDisplay_(sheet);
    } catch (e) {}
  }

  return {fixed: fixed, checked: checked};
}

function getExistingQualificationRowMap_(sheet, headers) {
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  const ownerCol = headers.indexOf("所有者") + 1;
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  if (ownerCol <= 0 || categoryCol <= 0 || nameCol <= 0) return map;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach((row, index) => {
    const owner = String(row[ownerCol - 1] || "").trim();
    const category = String(row[categoryCol - 1] || "").trim();
    const qualification = String(row[nameCol - 1] || "").trim();
    if (!owner || !category || !qualification) return;
    const key = buildQualificationDuplicateKey_(owner, category, qualification);
    if (!map[key]) {
      map[key] = {rowNumber: index + 2, rowValues: row};
    }
  });
  return map;
}

function updateExistingQualificationRow_(sheet, headers, rowNumber, sourceObj, options) {
  if (!sheet || rowNumber < 2) return {updated: false};
  const forceOverwrite = !!(options && options.forceOverwrite);
  const current = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const next = current.slice();
  let changed = false;

  const fieldRules = [
    {header: "取得区分", flag: "hasPossession"},
    {header: "保有状況", flag: "hasHoldingStatus"},
    {header: "取得日", flag: "hasAcquired"},
    {header: "更新期限", flag: "hasDue"},
    {header: "コピー有無", flag: "hasCopy"},
    {header: "状態", flag: "hasStatus"},
    {header: "備考", flag: "hasNote"}
  ];

  fieldRules.forEach(rule => {
    const colIndex = headers.indexOf(rule.header);
    if (colIndex < 0) return;
    const shouldUpdate = forceOverwrite || !!(options && options[rule.flag]);
    if (!shouldUpdate) return;
    const value = sourceObj[rule.header] !== undefined ? sourceObj[rule.header] : "";
    if (!sameCellValue_(next[colIndex], value)) {
      next[colIndex] = value;
      changed = true;
    }
  });

  // 期限・状態を変えた場合は通知も再計算する。強制上書き時も現在値から再計算する。
  const noticeIndex = headers.indexOf("通知");
  if (noticeIndex >= 0 && (forceOverwrite || (options && (options.hasDue || options.hasStatus)))) {
    const dueIndex = headers.indexOf("更新期限");
    const statusIndex = headers.indexOf("状態");
    const due = dueIndex >= 0 ? next[dueIndex] : "";
    const status = statusIndex >= 0 ? next[statusIndex] : "";
    const notice = getNoticeText(due, status);
    if (!sameCellValue_(next[noticeIndex], notice)) {
      next[noticeIndex] = notice;
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([next]);
  }
  return {updated: changed};
}

function sameCellValue_(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if (a instanceof Date || b instanceof Date) {
    const at = a instanceof Date ? a.getTime() : new Date(a).getTime();
    const bt = b instanceof Date ? b.getTime() : new Date(b).getTime();
    if (!isNaN(at) && !isNaN(bt)) return at === bt;
  }
  return String(a || "") === String(b || "");
}


function splitQualificationNames_(value) {
  const text = String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(/[，,、；;／\/]+/g, "\n");

  const seen = {};
  const result = [];
  text.split(/\n+/).forEach(part => {
    const parsed = parseQualificationSelection_(part);
    const name = parsed.name;
    if (!name) return;
    const key = normalizeTextForKey_(name);
    if (seen[key]) return;
    seen[key] = true;
    result.push(name);
  });
  return result;
}

/******************************
 * v9.7.32 資格一括登録シート方式・メニュー関数修正
 ******************************/

function createQualificationBulkImportSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    ensureQualificationBulkImportSheet_();
    toast_("資格一括登録シートを作成/修復しました");
  } finally {
    lock.releaseLock();
  }
}

function applyQualificationBulkImportToManagement() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const result = applyQualificationBulkImportToManagement_();
    toast_("資格一括登録を反映しました。追加 " + result.added + "件 / 重複スキップ " + result.skipped + "件 / 要確認 " + result.invalid + "件");
  } finally {
    lock.releaseLock();
  }
}

function clearQualificationBulkImportSheet() {
  const sheet = ensureQualificationBulkImportSheet_();
  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  toast_("資格一括登録シートをクリアしました（ヘッダーは残しています）");
}

function createEmployeeQualificationListSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const count = createEmployeeQualificationListSheet_();
    toast_("社員別資格一覧を作成/更新しました（" + count + "件）");
  } finally {
    lock.releaseLock();
  }
}

function createQualificationHolderListSheet() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const count = createQualificationHolderListSheet_();
    toast_("資格別保有者一覧を作成/更新しました（" + count + "件）");
  } finally {
    lock.releaseLock();
  }
}

function ensureQualificationBulkImportSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["資格一括登録"] || ["所有者", "区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "備考", "登録結果"];
  let sheet = ss.getSheetByName("資格一括登録");
  if (!sheet) sheet = ss.insertSheet("資格一括登録");
  migrateQualificationAcquisitionStatusColumns_(sheet);

  ensureColumns_(sheet, headers.length);
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const currentHeaders = getHeaders_(sheet);
    if (!currentHeaders.length || currentHeaders[0] !== "所有者") {
      sheet.clear();
      sheet.clearFormats();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      repairHeaderOrderPreserveData_(sheet, headers);
    }
  }

  formatQualificationBulkImportSheet_(sheet);
  applyQualificationBulkImportValidations_(sheet);
  return sheet;
}

function repairHeaderOrderPreserveData_(sheet, desiredHeaders) {
  const currentHeaders = getHeaders_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const currentValues = sheet.getRange(1, 1, lastRow, currentHeaders.length).getValues();
  const indexMap = {};
  currentHeaders.forEach((h, i) => { if (h) indexMap[h] = i; });

  const reordered = currentValues.map((row, rowIndex) => desiredHeaders.map(header => {
    if (rowIndex === 0) return header;
    const oldIndex = indexMap[header];
    return oldIndex === undefined ? "" : row[oldIndex];
  }));

  removeFilter_(sheet);
  sheet.clearContents();
  ensureColumns_(sheet, desiredHeaders.length);
  sheet.getRange(1, 1, reordered.length, desiredHeaders.length).setValues(reordered);
  if (sheet.getMaxColumns() > desiredHeaders.length) {
    try { sheet.deleteColumns(desiredHeaders.length + 1, sheet.getMaxColumns() - desiredHeaders.length); } catch (e) {}
  }
}

function formatQualificationBulkImportSheet_(sheet) {
  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  try {
    const lastRow = Math.max(sheet.getLastRow(), 2);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#d9ead3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    sheet.getRange(1, 1, lastRow, headers.length)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment("middle");
    sheet.setFrozenRows(1);
    setColumnWidthIfExists_(sheet, headers, "所有者", 100);
    setColumnWidthIfExists_(sheet, headers, "区分", 100);
    setColumnWidthIfExists_(sheet, headers, "資格名", 240);
    setColumnWidthIfExists_(sheet, headers, "取得区分", 90);
    setColumnWidthIfExists_(sheet, headers, "取得日", 95);
    setColumnWidthIfExists_(sheet, headers, "更新期限", 95);
    setColumnWidthIfExists_(sheet, headers, "コピー有無", 90);
    setColumnWidthIfExists_(sheet, headers, "状態", 90);
    setColumnWidthIfExists_(sheet, headers, "備考", 180);
    setColumnWidthIfExists_(sheet, headers, "登録結果", 150);
    createFilterSafely_(sheet, headers.length);
  } catch (e) {}
}

function applyQualificationBulkImportValidations_(sheet) {
  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  const rowCount = Math.min(Math.max(sheet.getMaxRows() - 1, 50), 150);
  const staffList = [CLEAR_LABEL, ...getStaffMembers_()];
  const categoryList = [CLEAR_LABEL, "資格", "技能講習", "特別教育", "安全教育", "免許", "その他"];
  const qualificationList = [CLEAR_LABEL, ...getQualificationMasterNames_()];
  const possessionList = [CLEAR_LABEL, "技士", "技士補", "保有", "修了", "免許"];
  const holdingStatusList = [CLEAR_LABEL, "保有", "未確認", "失効", "不要"];
  const copyList = [CLEAR_LABEL, "有", "無", "未確認"];
  const statusList = [CLEAR_LABEL, "有効", "更新予定", "更新済", "期限切れ", "失効"];

  setDropdownByHeaderIfExists_(sheet, headers, "所有者", staffList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "区分", categoryList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "資格名", qualificationList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "取得区分", possessionList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "保有状況", holdingStatusList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "コピー有無", copyList, rowCount);
  setDropdownByHeaderIfExists_(sheet, headers, "状態", statusList, rowCount);

  const acquiredCol = headers.indexOf("取得日") + 1;
  const dueCol = headers.indexOf("更新期限") + 1;
  if (acquiredCol > 0) sheet.getRange(2, acquiredCol, rowCount, 1).setNumberFormat("yyyy/mm/dd");
  if (dueCol > 0) sheet.getRange(2, dueCol, rowCount, 1).setNumberFormat("yyyy/mm/dd");
}


function setDropdownByHeaderIfExists_(sheet, headers, header, list, rowCount) {
  const col = headers.indexOf(header) + 1;
  if (col <= 0 || !list || !list.length) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, col, rowCount, 1).setDataValidation(rule);
}

function getQualificationMasterNames_() {
  return getQualificationMasterItems_().map(item => item.name);
}


function getQualificationPreferredCategoryMap_() {
  // p21:
  // 資格マスタに旧区分の重複行が残っていても、紙版標準マスタの区分を正として扱う。
  // 例: フルハーネス、石綿取扱、足場特別教育などが「資格」候補へ混ざる事故を防ぐ。
  const map = {};
  try {
    getQualificationPaperMasterRows_().forEach(item => {
      const name = cleanQualificationName_(item.name);
      const category = normalizeQualificationCategoryValue_(item.category);
      const key = normalizeTextForKey_(name);
      if (key && category) map[key] = category;
    });
  } catch (e) {
    console.log("getQualificationPreferredCategoryMap_ skip: " + e.message);
  }
  return map;
}

function getQualificationValidationListsByCategory_() {
  // p33:
  // 区分ごとの資格名リストは、入力規則を貼るたびに資格マスタを読み直すと重い。
  // 1回の実行内ではキャッシュし、資格管理のプルダウン復旧・ソート・行挿入復旧を軽くする。
  if (QUALIFICATION_VALIDATION_LISTS_CACHE_) return QUALIFICATION_VALIDATION_LISTS_CACHE_;

  const lists = {
    "全資格": [CLEAR_LABEL],
    "資格": [CLEAR_LABEL],
    "技能講習": [CLEAR_LABEL],
    "特別教育": [CLEAR_LABEL],
    "安全教育": [CLEAR_LABEL],
    "免許": [CLEAR_LABEL],
    "その他": [CLEAR_LABEL]
  };
  const seenByCategory = {};
  Object.keys(lists).forEach(key => seenByCategory[key] = {});

  sortQualificationMasterItemsForDropdown_(getQualificationMasterItems_()).forEach(item => {
    const itemCategory = normalizeQualificationCategoryValue_(item.category) || "その他";
    const name = cleanQualificationName_(item.name);
    const nameKey = normalizeTextForKey_(name);
    if (!name || !nameKey) return;

    const addToList_ = (categoryKey) => {
      if (!lists[categoryKey]) {
        lists[categoryKey] = [CLEAR_LABEL];
        seenByCategory[categoryKey] = {};
      }
      if (seenByCategory[categoryKey][nameKey]) return;
      seenByCategory[categoryKey][nameKey] = true;
      lists[categoryKey].push(name);
    };

    addToList_("全資格");
    addToList_(itemCategory);
  });

  Object.keys(lists).forEach(key => {
    if (!lists[key] || lists[key].length <= 1) lists[key] = [CLEAR_LABEL];
  });

  QUALIFICATION_VALIDATION_LISTS_CACHE_ = lists;
  return lists;
}

function getQualificationNamesForValidationCategory_(category) {
  const targetCategory = normalizeQualificationCategoryValue_(category);
  const lists = getQualificationValidationListsByCategory_();
  if (!targetCategory) return lists["全資格"] || [CLEAR_LABEL];
  return lists[targetCategory] || [CLEAR_LABEL];
}

function buildQualificationNameValidationRuleForCategory_(category, allowInvalid) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(getQualificationNamesForValidationCategory_(category), true)
    .setAllowInvalid(allowInvalid === true)
    .build();
}


function getQualificationManagementDropdownTargetRowCount_(sheet) {
  // p33:
  // 行挿入ではonEditが走らず、挿入行の入力規則が空になることがある。
  // そのため、資格管理だけは「最低500行」または「最終行＋200行」まで
  // 所有者・区分・資格名などのプルダウンを標準で維持する。
  if (!sheet) return 1;
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const targetLastRow = Math.min(
    maxRows,
    Math.max(lastRow + QUALIFICATION_MANAGEMENT_DROPDOWN_BUFFER_ROWS, QUALIFICATION_MANAGEMENT_DROPDOWN_MIN_ROWS + 1)
  );
  return Math.max(targetLastRow - 1, 1);
}

function getQualificationManagementNameValidationRowCount_(sheet) {
  return getQualificationManagementDropdownTargetRowCount_(sheet);
}

function repairQualificationManagementDropdownAroundEditedRow_(sheet, editedRow) {
  // p27:
  // onEditでは全行修復を避け、編集行とその下の予備行だけ資格名プルダウンを維持する。
  // 区分変更・資格名入力の直後でも、続けて次の行へ入力できるようにする。
  if (!sheet || sheet.getName() !== "資格管理") return {rows: 0};

  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  if (categoryCol <= 0 || nameCol <= 0) return {rows: 0};

  const maxRows = sheet.getMaxRows();
  const startRow = Math.max(2, Number(editedRow || 2));
  const endRow = Math.min(maxRows, Math.max(startRow + QUALIFICATION_MANAGEMENT_ONEDIT_BUFFER_ROWS, sheet.getLastRow() + 20));
  const rowCount = Math.max(endRow - startRow + 1, 1);

  const categories = sheet.getRange(startRow, categoryCol, rowCount, 1)
    .getValues()
    .map(row => normalizeQualificationCategoryValue_(row[0]));

  const ruleCache = {};
  let runStart = 0;
  let currentHeader = null;

  const applyRun_ = (startIndex, endIndex, helperHeader) => {
    if (startIndex > endIndex) return;
    if (!ruleCache[helperHeader]) {
      const categoryForRule = helperHeader === "全資格" ? "" : helperHeader;
      ruleCache[helperHeader] = buildQualificationNameValidationRuleForCategory_(categoryForRule, false);
    }
    sheet.getRange(startRow + startIndex, nameCol, endIndex - startIndex + 1, 1)
      .setDataValidation(ruleCache[helperHeader]);
  };

  categories.forEach((category, i) => {
    const helperHeader = getQualificationHelperHeaderForCategory_(category);
    if (i === 0) {
      currentHeader = helperHeader;
      runStart = 0;
      return;
    }
    if (helperHeader !== currentHeader) {
      applyRun_(runStart, i - 1, currentHeader);
      currentHeader = helperHeader;
      runStart = i;
    }
  });

  if (categories.length > 0) {
    applyRun_(runStart, categories.length - 1, currentHeader || "全資格");
  }

  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
  return {rows: rowCount};
}


function getQualificationManagementDropdownApplyRowCount_(sheet) {
  return getQualificationManagementDropdownTargetRowCount_(sheet);
}

function setQualificationManagementFixedDropdownsForRange_(sheet, startRow, rowCount) {
  if (!sheet || sheet.getName() !== "資格管理") return {rows: 0, cols: 0};

  const headers = getHeaders_(sheet);
  if (!headers.length) return {rows: 0, cols: 0};

  const firstRow = Math.max(2, Number(startRow || 2));
  const count = Math.max(1, Number(rowCount || 1));
  const maxRows = sheet.getMaxRows();
  if (firstRow > maxRows) return {rows: 0, cols: 0};
  const safeCount = Math.min(count, maxRows - firstRow + 1);

  const fixedHeaders = ["所有者", "区分", "取得区分", "保有状況", "コピー有無", "状態"];
  let appliedCols = 0;

  fixedHeaders.forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col <= 0) return;

    const list = getDropdownListForHeader_("資格管理", header);
    if (!list || !list.length) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(firstRow, col, safeCount, 1).setDataValidation(rule);
    appliedCols++;
  });

  return {rows: safeCount, cols: appliedCols};
}

function refreshQualificationManagementDropdownsLight_(sheet) {
  // p32:
  // 資格管理のプルダウン復旧を1か所に集約する。
  // A列「所有者」などの名前系プルダウンと、C列「資格名」の区分別プルダウンを両方戻す。
  if (!sheet || sheet.getName() !== "資格管理") return {fixedRows: 0, fixedCols: 0, nameRows: 0, cleared: 0};

  let fixed = {rows: 0, cols: 0};
  let nameResult = {rows: 0, cleared: 0};

  try {
    const fixedRowCount = getQualificationManagementDropdownApplyRowCount_(sheet);
    fixed = setQualificationManagementFixedDropdownsForRange_(sheet, 2, fixedRowCount);
  } catch (e) {
    console.log("refreshQualificationManagementDropdownsLight_ fixed dropdown skip: " + e.message);
  }

  try {
    nameResult = applyQualificationManagementNameValidationsStrictByRow_(sheet);
  } catch (e) {
    console.log("refreshQualificationManagementDropdownsLight_ name dropdown skip: " + e.message);
  }

  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}

  return {
    fixedRows: fixed.rows || 0,
    fixedCols: fixed.cols || 0,
    nameRows: nameResult.rows || 0,
    cleared: nameResult.cleared || 0
  };
}

function repairQualificationManagementDropdownsAroundEditedRow_(sheet, editedRow) {
  // p32:
  // onEdit後にC列「資格名」だけでなく、A列「所有者」などの名前プルダウンも
  // 編集行周辺へ貼り直す。これで入力を続けても名前候補が消えにくい。
  if (!sheet || sheet.getName() !== "資格管理") return {rows: 0};

  const maxRows = sheet.getMaxRows();
  const startRow = Math.max(2, Number(editedRow || 2));
  const endRow = Math.min(maxRows, Math.max(startRow + QUALIFICATION_MANAGEMENT_ONEDIT_BUFFER_ROWS, sheet.getLastRow() + 20));
  const rowCount = Math.max(endRow - startRow + 1, 1);

  try { setQualificationManagementFixedDropdownsForRange_(sheet, startRow, rowCount); } catch (e) {}
  return repairQualificationManagementDropdownAroundEditedRow_(sheet, editedRow);
}

function applyQualificationManagementNameValidationsStrictByRow_(sheet) {
  // p25:
  // C列「資格名」はB列「区分」に応じて厳密設定する。
  // p24の行単位setDataValidationは安全だが、行数が多いと重くなるため、
  // 同じ区分が連続する範囲ごとにまとめて入力規則を貼る。
  if (!sheet || sheet.getName() !== "資格管理") return {rows: 0, cleared: 0};

  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  if (categoryCol <= 0 || nameCol <= 0) return {rows: 0, cleared: 0};

  const rowCount = getQualificationManagementNameValidationRowCount_(sheet);
  const nameRange = sheet.getRange(2, nameCol, rowCount, 1);

  // p34:
  // ここで先に clearDataValidations() を実行すると、途中で時間超過・エラーになった場合に
  // C列の資格名プルダウンだけ消えた状態が残る。setDataValidation() は既存規則を上書きできるため、
  // 先消しせず、連続範囲ごとに上書き貼りする。

  const categories = sheet.getRange(2, categoryCol, rowCount, 1)
    .getValues()
    .map(r => normalizeQualificationCategoryValue_(r[0]));
  const names = nameRange.getValues();
  let cleared = 0;

  const cleanedNames = names.map((row, i) => {
    const category = categories[i];
    const name = cleanQualificationName_(row[0]);
    if (category && name && name !== CLEAR_LABEL && !isQualificationNameInCategory_(name, category)) {
      cleared++;
      return [""];
    }
    return [row[0]];
  });

  if (cleared > 0) {
    nameRange.setValues(cleanedNames);
  }

  const ruleCache = {};
  let runStart = 0;
  let currentHeader = null;

  const applyRun_ = (startIndex, endIndex, helperHeader) => {
    if (startIndex > endIndex) return;
    if (!ruleCache[helperHeader]) {
      const categoryForRule = helperHeader === "全資格" ? "" : helperHeader;
      ruleCache[helperHeader] = buildQualificationNameValidationRuleForCategory_(categoryForRule, false);
    }
    sheet.getRange(2 + startIndex, nameCol, endIndex - startIndex + 1, 1)
      .setDataValidation(ruleCache[helperHeader]);
  };

  categories.forEach((category, i) => {
    const helperHeader = getQualificationHelperHeaderForCategory_(category);
    if (i === 0) {
      currentHeader = helperHeader;
      runStart = 0;
      return;
    }

    if (helperHeader !== currentHeader) {
      applyRun_(runStart, i - 1, currentHeader);
      currentHeader = helperHeader;
      runStart = i;
    }
  });

  if (categories.length > 0) {
    applyRun_(runStart, categories.length - 1, currentHeader || "全資格");
  }

  return {rows: rowCount, cleared: cleared};
}

function getQualificationMasterItems_() {
  // p37:
  // 資格マスタシートの手動変更を正とする。
  // 紙版標準マスタは「不足候補の補完」と「表示順のフォールバック」にだけ使い、
  // 既存の資格名についてはシート側の区分を上書きしない。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("資格マスタ");
  if (!sheet || sheet.getLastRow() < 2) {
    try {
      rebuildQualificationMasterClean_();
      sheet = ss.getSheetByName("資格マスタ");
    } catch (e) {
      console.log("getQualificationMasterItems_ rebuild skip: " + e.message);
    }
  }

  const items = [];
  const seen = {};

  const addItem_ = (order, category, name, source) => {
    const cleanName = normalizeSpecialEducationQualificationName_(name);
    if (!cleanName) return;

    const nameKey = normalizeTextForKey_(cleanName);
    const sheetCategory = normalizeQualificationCategoryValue_(category);
    const fallbackCategory = source === "paper" ? (getQualificationCategoryOverrideByName_(cleanName) || sheetCategory) : sheetCategory;
    const cleanCategory = fallbackCategory || "その他";
    const numericOrder = Number(order) || 999999;
    const key = nameKey;

    if (seen[key] !== undefined) {
      const existing = items[seen[key]];

      if (source === "sheet") {
        // シートの手動変更を最優先。
        existing.category = cleanCategory;
        existing.order = numericOrder;
        existing.source = "sheet";
      } else if (existing.source !== "sheet") {
        // シートに無い候補だけ紙版標準を使う。
        existing.category = cleanCategory;
        if (numericOrder < existing.order) existing.order = numericOrder;
        existing.source = source || existing.source;
      } else {
        // 既存シート行の区分は変えない。表示順が空欄の場合だけ紙順をフォールバックにする。
        if ((!existing.order || existing.order === 999999) && numericOrder) existing.order = numericOrder;
      }
      return;
    }

    seen[key] = items.length;
    items.push({
      order: numericOrder,
      category: cleanCategory,
      name: cleanName,
      source: source || "sheet"
    });
  };

  if (sheet && sheet.getLastRow() >= 2) {
    const headers = getHeaders_(sheet);
    const orderCol = headers.indexOf("表示順");
    const categoryCol = headers.indexOf("区分");
    const nameCol = headers.indexOf("資格名");
    const visibleCol = headers.indexOf("一覧表示");

    if (nameCol >= 0) {
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
      values.forEach(row => {
        const visible = visibleCol >= 0 ? row[visibleCol] : true;
        if (visible === false || String(visible).toUpperCase() === "FALSE") return;
        addItem_(
          orderCol >= 0 ? row[orderCol] : 999999,
          categoryCol >= 0 ? row[categoryCol] : "その他",
          nameCol >= 0 ? row[nameCol] : "",
          "sheet"
        );
      });
    }
  }

  try {
    getQualificationPaperMasterRows_().forEach(item => {
      addItem_(item.order, item.category, item.name, "paper");
    });
  } catch (e) {
    console.log("getQualificationMasterItems_ paper rows skip: " + e.message);
  }

  items.sort((a, b) => {
    if (a.order !== b.order) return a.order - b.order;
    const categoryDiff = getQualificationCategoryOrder_(a.category) - getQualificationCategoryOrder_(b.category);
    if (categoryDiff !== 0) return categoryDiff;
    return a.name.localeCompare(b.name, "ja");
  });

  return items;
}

function applyQualificationBulkImportToManagement_() {
  const importSheet = ensureQualificationBulkImportSheet_();
  const targetSheet = ensureQualificationManagementSheet_();
  const importHeaders = getHeaders_(importSheet);
  const targetHeaders = getHeaders_(targetSheet);
  const resultCol = importHeaders.indexOf("登録結果") + 1;
  const lastRow = importSheet.getLastRow();
  if (lastRow < 2) return {added: 0, skipped: 0, invalid: 0};

  const importValues = importSheet.getRange(2, 1, lastRow - 1, importHeaders.length).getValues();
  const masterMap = getQualificationMasterCategoryMap_();
  const existingKeys = getExistingQualificationKeys_(targetSheet, targetHeaders);
  const appendRows = [];
  const resultValues = [];
  let added = 0;
  let skipped = 0;
  let invalid = 0;

  importValues.forEach(row => {
    const obj = objectFromRow(importHeaders, row);
    const owner = String(obj["所有者"] || "").trim();
    const qualification = String(obj["資格名"] || "").trim();
    let category = String(obj["区分"] || "").trim();
    let possession = String(obj["取得区分"] || "").trim();
    let status = String(obj["状態"] || "").trim();
    const acquired = obj["取得日"] || "";
    const due = obj["更新期限"] || "";
    const copy = String(obj["コピー有無"] || "").trim();
    const holdingStatus = String(obj["保有状況"] || "").trim();
    const note = obj["備考"] || "";

    if (!owner && !qualification && !category && !possession && !acquired && !due && !copy && !status && !note) {
      resultValues.push([""]);
      return;
    }

    if (!owner || !qualification) {
      invalid++;
      resultValues.push(["要確認：所有者または資格名が空欄"]);
      return;
    }

    if (!category) category = masterMap[normalizeTextForKey_(qualification)] || "その他";
    if (!possession) possession = "保有";
    if (!status) status = "有効";

    const key = buildQualificationDuplicateKey_(owner, category, qualification);
    if (existingKeys[key]) {
      skipped++;
      resultValues.push(["重複スキップ"]);
      return;
    }

    const rowObj = {
      "所有者": owner,
      "区分": category,
      "資格名": qualification,
      "取得区分": possession,
      "保有状況": holdingStatus || "保有",
      "取得日": acquired,
      "更新期限": due,
      "コピー有無": copy || "未確認",
      "状態": status,
      "通知": getNoticeText(due, status),
      "備考": note,
      "資格ID": buildRecordId_("資格管理")
    };

    appendRows.push(targetHeaders.map(header => {
      if (header === READ_HEADER || getPersonalMembers_().includes(header)) return false;
      return rowObj[header] !== undefined ? rowObj[header] : "";
    }));
    existingKeys[key] = true;
    added++;
    resultValues.push(["追加済"]);
  });

  if (appendRows.length) {
    const startRow = Math.max(targetSheet.getLastRow() + 1, 2);
    targetSheet.getRange(startRow, 1, appendRows.length, targetHeaders.length).setValues(appendRows);
    try {
      setCheckboxesForDataRows(targetSheet);
      applyQualificationSingleSheetLightSettings_(targetSheet);
    } catch (e) {}
  }

  if (resultCol > 0 && resultValues.length) {
    importSheet.getRange(2, resultCol, resultValues.length, 1).setValues(resultValues);
  }

  return {added: added, skipped: skipped, invalid: invalid};
}

function getQualificationMasterCategoryMap_() {
  // p37: 資格マスタの手動区分を正とするため、ここでは強制区分補正をかけない。
  const map = {};
  getQualificationMasterItems_().forEach(item => {
    const name = normalizeSpecialEducationQualificationName_(item.name);
    map[normalizeTextForKey_(name)] = item.category || "その他";
  });
  return map;
}

function getExistingQualificationKeys_(sheet, headers) {
  const keys = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return keys;

  const ownerCol = headers.indexOf("所有者");
  const categoryCol = headers.indexOf("区分");
  const qualificationCol = headers.indexOf("資格名");
  if (ownerCol < 0 || qualificationCol < 0) return keys;

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  values.forEach(row => {
    if (!isBusinessDataRow_(headers, row, "資格管理")) return;
    const owner = row[ownerCol];
    const category = categoryCol >= 0 ? row[categoryCol] : "";
    const qualification = row[qualificationCol];
    const key = buildQualificationDuplicateKey_(owner, category, qualification);
    if (normalizeTextForKey_(owner) && normalizeTextForKey_(qualification)) keys[key] = true;
  });
  return keys;
}

function buildQualificationDuplicateKey_(owner, category, qualification) {
  return [owner, category, qualification].map(normalizeTextForKey_).join("||");
}

function getQualificationManagementObjects_() {
  const sheet = ensureQualificationManagementSheet_();
  const headers = getHeaders_(sheet);
  if (sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const rows = [];
  values.forEach(row => {
    if (!isBusinessDataRow_(headers, row, "資格管理")) return;
    const obj = objectFromRow(headers, row);
    if (!normalizeTextForKey_(obj["所有者"]) || !normalizeTextForKey_(obj["資格名"])) return;
    rows.push(obj);
  });
  const masterOrderMap = getQualificationMasterOrderMap_();
  rows.sort((a, b) => {
    const ao = normalizeTextForKey_(a["所有者"]);
    const bo = normalizeTextForKey_(b["所有者"]);
    if (ao !== bo) return ao.localeCompare(bo, "ja");

    const acText = String(a["区分"] || "").trim();
    const bcText = String(b["区分"] || "").trim();
    const aqOrder = getQualificationMasterSortOrder_(acText, a["資格名"], masterOrderMap);
    const bqOrder = getQualificationMasterSortOrder_(bcText, b["資格名"], masterOrderMap);
    if (aqOrder !== bqOrder) return aqOrder - bqOrder;

    const ac = getQualificationCategoryOrder_(acText);
    const bc = getQualificationCategoryOrder_(bcText);
    if (ac !== bc) return ac - bc;

    return normalizeTextForKey_(a["資格名"]).localeCompare(normalizeTextForKey_(b["資格名"]), "ja");
  });
  return rows;
}

function createEmployeeQualificationListSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = getQualificationManagementObjects_();

  // v9.7.48f:
  // 社員名は見出し行に出すため、明細側の「所有者」列は出さない。
  // これにより、社員別資格一覧の左側に大きな空欄が出る問題を防ぐ。
  const outHeaders = ["区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"];
  let sheet = ss.getSheetByName("社員別資格一覧");
  if (!sheet) sheet = ss.insertSheet("社員別資格一覧");

  removeFilter_(sheet);
  clearAllRowGroupsAndShowRows_(sheet);
  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);

  const grouped = {};
  rows.forEach(obj => {
    const owner = String(obj["所有者"] || "").trim() || "（所有者未設定）";
    if (!grouped[owner]) grouped[owner] = [];
    grouped[owner].push(obj);
  });

  const masterOrderMap = getQualificationMasterOrderMap_();
  const staffOrder = getStaffOrderMap_();
  const output = [];
  const summaryRows = [];
  const groupRanges = [];

  Object.keys(grouped).sort((a, b) => {
    const ao = getStaffSortOrder_(a, staffOrder);
    const bo = getStaffSortOrder_(b, staffOrder);
    if (ao !== bo) return ao - bo;
    return a.localeCompare(b, "ja");
  }).forEach(owner => {
    const items = grouped[owner];
    items.sort((a, b) => {
      const ac = normalizeQualificationCategoryValue_(a["区分"]);
      const bc = normalizeQualificationCategoryValue_(b["区分"]);
      const acOrder = getQualificationCategoryOrder_(ac);
      const bcOrder = getQualificationCategoryOrder_(bc);
      if (acOrder !== bcOrder) return acOrder - bcOrder;
      const aqOrder = getQualificationMasterSortOrder_(ac, a["資格名"], masterOrderMap);
      const bqOrder = getQualificationMasterSortOrder_(bc, b["資格名"], masterOrderMap);
      if (aqOrder !== bqOrder) return aqOrder - bqOrder;
      const dueA = getDateTimeSortValue_(a["更新期限"]);
      const dueB = getDateTimeSortValue_(b["更新期限"]);
      if (dueA !== dueB) return dueA - dueB;
      return String(a["資格名"] || "").localeCompare(String(b["資格名"] || ""), "ja");
    });

    const alertCount = items.filter(obj => {
      const notice = String(obj["通知"] || "");
      const status = String(obj["状態"] || "");
      const holding = String(obj["保有状況"] || "");
      return (notice && notice !== "完了") || status === "期限切れ" || status === "更新予定" || status === "失効" || holding === "失効";
    }).length;
    const copyUnknownCount = items.filter(obj => ["", "未確認", "無"].includes(String(obj["コピー有無"] || "").trim())).length;

    const summaryRow = output.length + 2;
    summaryRows.push(summaryRow);
    output.push([
      owner + "（取得数" + items.length + "件 / 注意" + alertCount + "件 / コピー未確認" + copyUnknownCount + "件）",
      "", "", "", "", "", "", "", "", ""
    ]);

    const detailStart = output.length + 2;
    items.forEach(obj => {
      output.push([
        obj["区分"] || "",
        obj["資格名"] || "",
        obj["取得区分"] || "",
        obj["保有状況"] || "",
        obj["取得日"] || "",
        obj["更新期限"] || "",
        obj["コピー有無"] || "",
        obj["状態"] || "",
        obj["通知"] || "",
        obj["備考"] || ""
      ]);
    });
    const detailEnd = output.length + 1;
    if (detailEnd >= detailStart) groupRanges.push([detailStart, detailEnd]);
  });

  if (output.length) sheet.getRange(2, 1, output.length, outHeaders.length).setValues(output);
  formatQualificationListOutputSheet_(sheet, outHeaders.length);

  summaryRows.forEach(rowNo => {
    try {
      sheet.getRange(rowNo, 1, 1, outHeaders.length)
        .setBackground("#e2f0d9")
        .setFontWeight("bold")
        .setHorizontalAlignment("left");
    } catch (e) {}
  });

  // v9.7.48g:
  // 社員別資格一覧は、社員見出し行だけで区切る。
  // 行グループを付けると、既存グループや折りたたみ状態が残って行番号が飛び、
  // 「変な場所でグループがかかっている」ように見えるため、ここではグループ化しない。
  clearAllRowGroupsAndShowRows_(sheet);
  applyColorRules(sheet);
  return rows.length;
}

function createQualificationHolderListSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = getQualificationManagementObjects_();
  const outHeaders = ["資格名", "区分", "所有者", "取得区分", "保有状況", "取得日", "更新期限", "コピー有無", "状態", "通知", "備考"];
  let sheet = ss.getSheetByName("資格別保有者一覧");
  if (!sheet) sheet = ss.insertSheet("資格別保有者一覧");

  removeFilter_(sheet);
  clearAllRowGroupsAndShowRows_(sheet);
  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);

  const grouped = {};
  rows.forEach(obj => {
    const q = String(obj["資格名"] || "").trim();
    if (!grouped[q]) grouped[q] = [];
    grouped[q].push(obj);
  });

  const output = [];
  const groupRanges = [];
  const masterOrderMap = getQualificationMasterOrderMap_();
  Object.keys(grouped).sort((a, b) => {
    const firstA = grouped[a] && grouped[a][0] ? grouped[a][0] : {};
    const firstB = grouped[b] && grouped[b][0] ? grouped[b][0] : {};
    const ao = getQualificationMasterSortOrder_(firstA["区分"], a, masterOrderMap);
    const bo = getQualificationMasterSortOrder_(firstB["区分"], b, masterOrderMap);
    if (ao !== bo) return ao - bo;
    return a.localeCompare(b, "ja");
  }).forEach(q => {
    const items = grouped[q];
    output.push([q + "（" + items.length + "名）", "", "", "", "", "", "", "", "", "", ""]);
    const detailStart = output.length + 2;
    items.sort((a, b) => normalizeTextForKey_(a["所有者"]).localeCompare(normalizeTextForKey_(b["所有者"]), "ja"));
    items.forEach(obj => {
      output.push([
        "",
        obj["区分"] || "",
        obj["所有者"] || "",
        obj["取得区分"] || "",
        obj["保有状況"] || "",
        obj["取得日"] || "",
        obj["更新期限"] || "",
        obj["コピー有無"] || "",
        obj["状態"] || "",
        obj["通知"] || "",
        obj["備考"] || ""
      ]);
    });
    const detailEnd = output.length + 1;
    if (detailEnd >= detailStart) groupRanges.push([detailStart, detailEnd]);
  });

  if (output.length) sheet.getRange(2, 1, output.length, outHeaders.length).setValues(output);
  formatQualificationListOutputSheet_(sheet, outHeaders.length);
  groupRanges.forEach(r => groupRangeRows_(sheet, r[0], r[1]));
  return rows.length;
}

function formatQualificationListOutputSheet_(sheet, colCount) {
  try {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const lastCol = colCount;

    sheet.setFrozenRows(1);
    showAllColumns_(sheet);

    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setBackground("#d9ead3")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(false);

    sheet.getRange(1, 1, lastRow, lastCol)
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment("middle")
      .setWrap(false);

    if (lastRow >= 2) {
      sheet.getRange(2, 1, lastRow - 1, lastCol)
        .setVerticalAlignment("middle")
        .setWrap(false);

      const firstColValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      firstColValues.forEach((r, i) => {
        const rowNo = i + 2;
        const v = String(r[0] || "");
        if (v.indexOf("（取得数") >= 0 || v.indexOf("（所有者") >= 0) {
          sheet.getRange(rowNo, 1, 1, lastCol)
            .setFontWeight("bold")
            .setBackground("#e2f0d9")
            .setHorizontalAlignment("left")
            .setVerticalAlignment("middle");
          sheet.setRowHeight(rowNo, 34);
        } else {
          sheet.setRowHeight(rowNo, 30);
        }
      });
    }

    createFilterSafely_(sheet, lastCol);

    // v9.7.48f: 社員別資格一覧は所有者を明細列に出さないため、列幅を詰める。
    const widths = [110, 320, 90, 95, 95, 105, 90, 90, 90, 230];
    widths.forEach((w, i) => {
      if (i + 1 <= lastCol) sheet.setColumnWidth(i + 1, w);
    });

    // 詳細行の区分・日付・状態系は中央、資格名・備考は左寄せにする。
    if (lastRow >= 2) {
      [1, 3, 4, 5, 6, 7, 8, 9].forEach(col => {
        if (col <= lastCol) sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("center");
      });
      [2, 10].forEach(col => {
        if (col <= lastCol) sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment("left");
      });
    }
  } catch (e) {
    console.log("formatQualificationListOutputSheet_ error: " + e.message);
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
      if (name === "工事予定" && h === "契約金額") setConstructionAmountTextColumn_(sheet, i + 1);
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

function clearAllRowGroupsAndShowRows_(sheet) {
  if (!sheet) return;

  // v9.7.48g:
  // 通常の clearRowGroups_ だけでは、折りたたみ済みグループや多重グループが残ることがある。
  // 社員別資格一覧や資格管理を作り直す前に、できるだけ展開・表示・解除を強めに行う。
  try { sheet.expandAllRowGroups(); } catch (e) {}
  try { sheet.showRows(1, sheet.getMaxRows()); } catch (e) {}

  try {
    for (let i = 0; i < 20; i++) {
      sheet.getRange(1, 1, sheet.getMaxRows(), 1).shiftRowGroupDepth(-1);
    }
  } catch (e) {}

  try { sheet.showRows(1, sheet.getMaxRows()); } catch (e) {}
}

function resetEmployeeQualificationListGroups() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("社員別資格一覧");
  if (!sheet) {
    toast_("社員別資格一覧シートが見つかりません");
    return;
  }
  clearAllRowGroupsAndShowRows_(sheet);
  toast_("社員別資格一覧の行グループを完全解除しました");
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
      ["driver_licenses", "運転免許管理", "license_issue_date", "DATE", "", "運転免許管理", "免許証交付日", "免許証1枚の交付日"],
      ["driver_licenses", "運転免許管理", "renewal_due_date", "DATE", "", "運転免許管理", "免許証有効期限", "一覧・要確認に反映"],
      ["driver_license_details", "運転免許明細", "license_type", "TEXT", "NOT NULL", "運転免許明細", "免許種類", "免許ごとの取得日管理"],
      ["driver_license_details", "運転免許明細", "acquired_date", "DATE", "", "運転免許明細", "取得日", "普通・大型特殊などの個別取得日"],
      ["qualifications", "資格管理", "qualification_name", "TEXT", "NOT NULL", "資格管理", "資格名", "資格マスタ化候補"],
      ["repairs", "備品修理管理", "repair_vendor", "TEXT", "", "備品修理管理", "修理業者", "旧 場所 から変更"],
      ["daily_receipts", "日報レシート管理", "amount", "NUMERIC", "", "日報レシート管理", "金額", "帳簿PDF出力候補"]
    ],
    "移行対応表": [
      ["シート名", "列名", "SQLテーブル", "SQLカラム", "備考"],
      ["工事予定", "工事名", "construction_projects", "project_name", ""],
      ["工事予定", "依頼主", "construction_projects", "client_name", "フィードバックで追加"],
      ["工事予定", "契約金額", "construction_projects", "contract_amount", "数値推奨。表示形式は円"],
      ["工事予定", "税", "construction_projects", "tax_status", "税込/税抜/不明"],
      ["電話履歴", "日付", "phone_logs", "call_date", "日時から分離"],
      ["電話履歴", "時間", "phone_logs", "call_time", "hh:mm"],
      ["車検管理", "車番", "vehicles", "vehicle_number", "文字列"],
      ["車検管理", "車検予定日", "vehicles", "inspection_scheduled_date", "次回車検期限から変更"],
      ["運転免許管理", "所有者", "driver_licenses", "owner_name", "担当から変更"],
      ["運転免許管理", "免許証有効期限", "driver_licenses", "renewal_due_date", "免許証全体の有効期限"],
      ["運転免許明細", "取得日", "driver_license_details", "acquired_date", "免許種類ごとの取得日"],
      ["資格管理", "資格名", "qualifications", "qualification_name", ""],
      ["備品修理管理", "購入日", "repairs", "purchase_date", "登録日から変更"],
      ["備品修理管理", "修理依頼者", "repairs", "requester_name", "担当から変更"]
    ],
    "SQLサンプル集": [
      ["区分", "用途", "SQL", "説明"],
      ["期限確認", "期限近い車検", "SELECT * FROM vehicles WHERE inspection_due_date <= CURRENT_DATE + INTERVAL '7 days';", "7日以内の車検確認"],
      ["電話", "未対応電話", "SELECT * FROM phone_logs WHERE phone_status <> '完了';", "折返し・未対応の抽出"],
      ["免許", "更新期限が近い免許", "SELECT * FROM driver_licenses WHERE renewal_due_date <= CURRENT_DATE + INTERVAL '30 days';", "30日以内の免許証有効期限確認"],
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


























/******************************
 * v9.7.48 優先修正 1-10
 ******************************/

function repairQualificationMasterNoticeAndAlias() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    ensureQualificationMasterNoticeColumn_();
    ensureQualificationAliasMasterSheet_();
    QUALIFICATION_DROPDOWN_HELPER_READY_ = false;
    ensureQualificationDropdownHelperSheet_();
    rebuildQualificationEasyPasteDependentDropdowns();
    const categorySortResult = repairQualificationManagementCategoryNameSort_();
    const categoryDropdownRows = categorySortResult.rows;
    toast_("資格マスタ通知日前列・表記ゆれマスタ・資格プルダウンを修復しました");
  } finally {
    lock.releaseLock();
  }
}

function ensureQualificationMasterNoticeColumn_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["資格マスタ"];
  let sheet = ss.getSheetByName("資格マスタ");
  if (!sheet) sheet = ss.insertSheet("資格マスタ");

  if (sheet.getLastRow() < 1 || !getHeaders_(sheet).length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    repairHeaderOrderPreserveData_(sheet, headers);
  }

  const currentHeaders = getHeaders_(sheet);
  const noticeCol = currentHeaders.indexOf("通知日前") + 1;
  const categoryCol = currentHeaders.indexOf("区分") + 1;
  const lastRow = sheet.getLastRow();
  if (noticeCol > 0 && lastRow >= 2) {
    const values = sheet.getRange(2, noticeCol, lastRow - 1, 1).getValues();
    if (categoryCol > 0) {
      const categories = sheet.getRange(2, categoryCol, lastRow - 1, 1).getValues();
      const next = values.map((row, i) => {
        const current = row[0];
        const category = String(categories[i][0] || "").trim();
        if (current !== "" && current !== null) return [current];
        return [category === "免許" ? 90 : ""];
      });
      sheet.getRange(2, noticeCol, next.length, 1).setValues(next);
    }
    sheet.getRange(2, noticeCol, Math.max(lastRow - 1, 1), 1).setNumberFormat("0").clearDataValidations();
  }

  const visibleCol = currentHeaders.indexOf("一覧表示") + 1;
  if (visibleCol > 0 && lastRow >= 2) {
    try { sheet.getRange(2, visibleCol, lastRow - 1, 1).insertCheckboxes(); } catch (e) {}
  }

  formatSheetBase_(sheet, headers.length);
  createFilterSafely_(sheet, headers.length);
}

function buildQualificationMasterMetaKey_(category, name) {
  return normalizeTextForKey_(category) + "||" + normalizeTextForKey_(cleanQualificationName_(name));
}

function getExistingQualificationMasterMetaMap_(sheet) {
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分");
  const nameCol = headers.indexOf("資格名");
  const noticeCol = headers.indexOf("通知日前");
  const visibleCol = headers.indexOf("一覧表示");
  const idCol = headers.indexOf("資格マスタID");
  if (categoryCol < 0 || nameCol < 0) return map;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach(row => {
    const key = buildQualificationMasterMetaKey_(row[categoryCol], row[nameCol]);
    if (!key || key === "||") return;
    map[key] = {
      noticeDays: noticeCol >= 0 ? row[noticeCol] : undefined,
      visible: visibleCol >= 0 ? row[visibleCol] : undefined,
      id: idCol >= 0 ? row[idCol] : ""
    };
  });
  return map;
}

function ensureQualificationAliasMasterSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const headers = getSheetHeaders_()["資格表記ゆれマスタ"] || ["入力名", "正式資格名", "区分", "取得区分", "備考"];
  let sheet = ss.getSheetByName("資格表記ゆれマスタ");
  if (!sheet) sheet = ss.insertSheet("資格表記ゆれマスタ");

  if (sheet.getLastRow() < 1 || !getHeaders_(sheet).length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    repairHeaderOrderPreserveData_(sheet, headers);
  }

  if (sheet.getLastRow() < 2) {
    const rows = [
      ["管理技術者", "監理技術者", "資格", "保有", "自動統一"],
      ["玉掛", "玉掛け", "技能講習", "修了", "自動統一"],
      ["車両系建設建設機械（整地）", "車両系建設機械（整地）", "技能講習", "修了", "自動統一"],
      ["建築物石綿含有建材調査者講", "建築物石綿含有建材調査者講習", "資格", "保有", "自動統一"],
      ["工作物石綿含有建材調査者講", "工作物石綿含有建材調査者講習", "資格", "保有", "自動統一"],
      ["仮払い機", "刈払い機", "安全教育", "修了", "自動統一"],
      ["刈払機", "刈払い機", "安全教育", "修了", "自動統一"],
      ["刈払い機取扱作業者", "刈払い機", "安全教育", "修了", "自動統一"],
      ["高所作業", "高所作業車", "特別教育", "修了", "自動統一"],
      ["高所作業車運転", "高所作業車", "特別教育", "修了", "自動統一"],
      ["研削トイシ", "研削砥石", "特別教育", "修了", "自動統一"],
      ["フルハーネス", "フルハーネス型墜落制止用器具", "特別教育", "修了", "自動統一"]
    ];
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // p36: 既存の資格表記ゆれマスタが空でない場合でも、重要な表記ゆれは不足分だけ追加する。
  try {
    upsertQualificationAliasRows_(sheet, headers, [
      ["仮払い機", "刈払い機", "安全教育", "修了", "自動統一"],
      ["刈払機", "刈払い機", "安全教育", "修了", "自動統一"],
      ["刈払い機取扱作業者", "刈払い機", "安全教育", "修了", "自動統一"],
      ["刈払機取扱作業者", "刈払い機", "安全教育", "修了", "自動統一"],
      ["高所作業", "高所作業車", "特別教育", "修了", "自動統一"],
      ["高所作業車運転", "高所作業車", "特別教育", "修了", "自動統一"],
      ["高所作業車特別教育", "高所作業車", "特別教育", "修了", "自動統一"]
    ]);
  } catch (e) {
    console.log("資格表記ゆれマスタp36補完をスキップ: " + e.message);
  }

  formatSheetBase_(sheet, headers.length);
  try {
    sheet.setColumnWidth(1, 240);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 180);
    createFilterSafely_(sheet, headers.length);
  } catch (e) {}
  return sheet;
}


function upsertQualificationAliasRows_(sheet, headers, rows) {
  if (!sheet || !headers || !headers.length || !rows || !rows.length) return 0;

  const inputCol = headers.indexOf("入力名") + 1;
  const formalCol = headers.indexOf("正式資格名") + 1;
  if (inputCol <= 0 || formalCol <= 0) return 0;

  const lastRow = sheet.getLastRow();
  const existing = {};
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    values.forEach(row => {
      const key = normalizeTextForKey_(row[inputCol - 1]) + "||" + normalizeTextForKey_(row[formalCol - 1]);
      if (key !== "||") existing[key] = true;
    });
  }

  const append = [];
  rows.forEach(row => {
    const key = normalizeTextForKey_(row[0]) + "||" + normalizeTextForKey_(row[1]);
    if (!key || key === "||" || existing[key]) return;
    const obj = {"入力名": row[0], "正式資格名": row[1], "区分": row[2], "取得区分": row[3], "備考": row[4]};
    append.push(headers.map(header => obj[header] !== undefined ? obj[header] : ""));
    existing[key] = true;
  });

  if (append.length) {
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, append.length, headers.length).setValues(append);
  }
  return append.length;
}
function getQualificationAliasMap_() {
  ensureQualificationAliasMasterSheet_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格表記ゆれマスタ");
  const map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const headers = getHeaders_(sheet);
  const inputCol = headers.indexOf("入力名");
  const officialCol = headers.indexOf("正式資格名");
  const categoryCol = headers.indexOf("区分");
  const acquisitionCol = headers.indexOf("取得区分");
  if (inputCol < 0 || officialCol < 0) return map;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  values.forEach(row => {
    const input = cleanQualificationName_(row[inputCol]);
    const official = cleanQualificationName_(row[officialCol]);
    if (!input || !official) return;
    map[normalizeTextForKey_(input)] = {
      name: official,
      category: categoryCol >= 0 ? normalizeQualificationCategoryValue_(row[categoryCol]) : "",
      acquisition: acquisitionCol >= 0 ? String(row[acquisitionCol] || "").trim() : ""
    };
  });
  return map;
}

function canonicalizeQualificationNameWithAlias_(name) {
  const clean = cleanQualificationName_(name);
  if (!clean) return {name: "", changed: false, category: "", acquisition: ""};
  const alias = getQualificationAliasMap_()[normalizeTextForKey_(clean)];
  if (!alias) return {name: clean, changed: false, category: "", acquisition: ""};
  return {name: alias.name, changed: alias.name !== clean, category: alias.category || "", acquisition: alias.acquisition || ""};
}

function getQualificationMasterNameListByCategory_(categoryFilter) {
  const categoryText = normalizeQualificationCategoryValue_(categoryFilter);
  return getQualificationMasterItems_()
    .filter(item => !categoryText || normalizeQualificationCategoryValue_(item.category) === categoryText)
    .map(item => item.name);
}


function getQualificationDropdownHelperSheetName_() {
  return "資格プルダウン補助";
}

function ensureQualificationDropdownHelperSheet() {
  QUALIFICATION_DROPDOWN_HELPER_READY_ = false;
  ensureQualificationDropdownHelperSheet_();
  toast_("資格プルダウン補助を作成/更新しました。続けて資格管理のプルダウンを完全復旧してください。");
}

function ensureQualificationDropdownHelperSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getQualificationDropdownHelperSheetName_();
  const headers = ["区分", "全資格", "資格", "技能講習", "特別教育", "安全教育", "免許", "その他", "取得区分", "保有状況", "コピー有無", "状態"];
  let sheet = ss.getSheetByName(sheetName);
  if (QUALIFICATION_DROPDOWN_HELPER_READY_ && sheet) return sheet;
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const categories = [CLEAR_LABEL, "資格", "技能講習", "特別教育", "安全教育", "免許", "その他"];
  const items = sortQualificationMasterItemsForDropdown_(getQualificationMasterItems_());
  const allNames = [CLEAR_LABEL].concat(items.map(item => item.name));
  const byCategory = {};
  ["資格", "技能講習", "特別教育", "安全教育", "免許", "その他"].forEach(c => byCategory[c] = [CLEAR_LABEL]);

  items.forEach(item => {
    const category = normalizeQualificationCategoryValue_(item.category) || "その他";
    const key = byCategory[category] ? category : "その他";
    byCategory[key].push(item.name);
  });

  const columns = {
    "区分": categories,
    "全資格": allNames,
    "資格": byCategory["資格"],
    "技能講習": byCategory["技能講習"],
    "特別教育": byCategory["特別教育"],
    "安全教育": byCategory["安全教育"],
    "免許": byCategory["免許"],
    "その他": byCategory["その他"],
    "取得区分": [CLEAR_LABEL, "技士", "技士補", "保有", "修了", "免許"],
    "保有状況": [CLEAR_LABEL, "保有", "未確認", "失効", "不要"],
    "コピー有無": [CLEAR_LABEL, "有", "無", "未確認"],
    "状態": getStatusListForSheet_("資格管理")
  };

  const maxRows = Math.max(2, ...headers.map(h => (columns[h] || [""]).length));
  if (sheet.getMaxRows() < maxRows) sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < headers.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());

  sheet.clear();
  sheet.clearFormats();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  headers.forEach((header, index) => {
    const values = columns[header] || [CLEAR_LABEL];
    const padded = [];
    for (let i = 0; i < maxRows - 1; i++) padded.push([values[i] || ""]);
    sheet.getRange(2, index + 1, padded.length, 1).setValues(padded);
  });

  formatSheetBase_(sheet, headers.length);
  try {
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 160);
    createFilterSafely_(sheet, headers.length);
    sheet.hideSheet();
  } catch (e) {}
  QUALIFICATION_DROPDOWN_HELPER_READY_ = true;
  return sheet;
}

function getQualificationHelperRangeByHeader_(helperHeader) {
  const sheet = ensureQualificationDropdownHelperSheet_();
  const headers = getHeaders_(sheet);
  const col = headers.indexOf(helperHeader) + 1;
  if (col <= 0) return null;

  const maxRows = Math.max(sheet.getLastRow(), 2);
  const values = sheet.getRange(2, col, maxRows - 1, 1).getValues();
  let last = 0;
  values.forEach((row, i) => {
    if (String(row[0] || "").trim()) last = i + 1;
  });
  return sheet.getRange(2, col, Math.max(last, 1), 1);
}

function setDropdownFromHelperRange_(sheet, col, helperHeader, allowInvalid, rowCount) {
  const range = getQualificationHelperRangeByHeader_(helperHeader);
  if (!range) return;
  const count = rowCount || getApplyRowCount_(sheet);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  sheet.getRange(2, col, count, 1).setDataValidation(rule);
}

function isQualificationManagementDropdownHeader_(header) {
  return ["区分", "資格名", "取得区分", "保有状況", "コピー有無", "状態"].includes(String(header || ""));
}

function applyQualificationManagementColumnValidation_(sheet, header, col) {
  if (!sheet || sheet.getName() !== "資格管理") return;
  const rowCount = getApplyRowCount_(sheet);
  if (header === "区分") return setDropdownFromHelperRange_(sheet, col, "区分", false, rowCount);

  // p24:
  // 資格名は列全体へ「全資格」を貼らない。
  // ここで全資格候補を貼ると、後から「入力シートだけ設定反映」を実行した時に
  // B列区分別の候補が上書きされ、資格行に特別教育/安全教育が混ざる。
  // C列は applyQualificationManagementNameValidationsStrictByRow_() で行ごとに貼る。
  if (header === "資格名") return;

  if (header === "取得区分") return setDropdownFromHelperRange_(sheet, col, "取得区分", false, rowCount);
  if (header === "保有状況") return setDropdownFromHelperRange_(sheet, col, "保有状況", false, rowCount);
  if (header === "コピー有無") return setDropdownFromHelperRange_(sheet, col, "コピー有無", false, rowCount);
  if (header === "状態") return setDropdownFromHelperRange_(sheet, col, "状態", false, rowCount);
}

function getQualificationHelperHeaderForCategory_(category) {
  const c = normalizeQualificationCategoryValue_(category);
  if (["資格", "技能講習", "特別教育", "安全教育", "免許", "その他"].includes(c)) return c;
  return "全資格";
}

function setQualificationManagementNameValidationForRow_(sheet, row, clearSelectedValue) {
  if (!sheet || sheet.getName() !== "資格管理" || row <= 1) return;
  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  if (nameCol <= 0) return;

  const category = categoryCol > 0 ? normalizeQualificationCategoryValue_(sheet.getRange(row, categoryCol).getValue()) : "";
  const rule = buildQualificationNameValidationRuleForCategory_(category, false);

  const cell = sheet.getRange(row, nameCol);
  // p34: 先に入力規則を消さず、上書きで貼る。
  cell.setDataValidation(rule);
  if (clearSelectedValue) cell.clearContent();
}


function sortQualificationMasterItemsForDropdown_(items) {
  return (items || []).slice().sort((a, b) => {
    const ao = Number(a && a.order) || 999999;
    const bo = Number(b && b.order) || 999999;
    if (ao !== bo) return ao - bo;

    const ac = normalizeQualificationCategoryValue_(a && a.category) || "その他";
    const bc = normalizeQualificationCategoryValue_(b && b.category) || "その他";
    const co = getQualificationCategoryOrder_(ac) - getQualificationCategoryOrder_(bc);
    if (co !== 0) return co;

    const an = cleanQualificationName_(a && a.name);
    const bn = cleanQualificationName_(b && b.name);
    return String(an).localeCompare(String(bn), "ja");
  });
}

function syncQualificationManagementNameAfterCategoryEdit_(sheet, row) {
  if (!sheet || sheet.getName() !== "資格管理" || row <= 1) return;

  const headers = getHeaders_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  if (categoryCol <= 0 || nameCol <= 0) return;

  const category = normalizeQualificationCategoryValue_(sheet.getRange(row, categoryCol).getValue());
  const nameCell = sheet.getRange(row, nameCol);
  const currentName = cleanQualificationName_(nameCell.getValue());

  // 先に入力規則を作り直す。これでC列の候補が、B列の区分ごとに並ぶ。
  setQualificationManagementNameValidationForRow_(sheet, row, false);

  // 区分が空欄なら、資格名は残して全資格候補に戻す。
  if (!category || !currentName || currentName === CLEAR_LABEL) return;

  if (!isQualificationNameInCategory_(currentName, category)) {
    nameCell.clearContent();
  }
}

function isQualificationNameInCategory_(name, category) {
  const cleanName = cleanQualificationName_(name);
  const targetCategory = normalizeQualificationCategoryValue_(category);
  if (!cleanName || !targetCategory) return true;

  const parsed = parseQualificationSelection_(cleanName);
  const parsedName = normalizeSpecialEducationQualificationName_(parsed.name || cleanName);
  const masterCategory = normalizeQualificationCategoryValue_(getQualificationMasterCategoryByName_(parsedName) || parsed.category);

  // マスタ未登録の手入力値は消しすぎない。備考の要確認で拾う。
  if (!masterCategory) return true;
  return masterCategory === targetCategory;
}

function repairQualificationManagementCategoryNameSort() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const result = repairQualificationManagementCategoryNameSort_();
    toast_("資格管理の名前・資格名プルダウンを復旧しました（資格名:" + result.rows + "行 / 所有者など:" + (result.fixedRows || 0) + "行 / 不一致クリア:" + result.cleared + "件）。");
  } finally {
    lock.releaseLock();
  }
}

function repairQualificationManagementCategoryNameSort_() {
  // p32:
  // 手動メニュー「資格管理の資格名プルダウンだけ直す」では、C列だけでなく
  // A列「所有者」などの名前プルダウンも一緒に復旧する。
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
  const result = repairQualificationManagementCategoryNameSortFast_();
  const fixed = refreshQualificationManagementDropdownsLight_(sheet);
  return {
    rows: fixed.nameRows || result.rows || 0,
    fixedRows: fixed.fixedRows || 0,
    cleared: (result.cleared || 0) + (fixed.cleared || 0)
  };
}

function repairQualificationManagementDropdowns() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    try { ensureQualificationManagementOnChangeTriggerQuiet_(); } catch (e) {
      console.log("資格管理行挿入検知トリガー確認をスキップ: " + e.message);
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
    const result = refreshQualificationManagementDropdownsLight_(sheet);
    toast_(
      "資格管理の名前・資格名プルダウンを復旧しました"
      + "（所有者など:" + (result.fixedRows || 0)
      + "行 / 資格名:" + (result.nameRows || 0)
      + "行 / 不一致クリア:" + (result.cleared || 0) + "件）。"
    );
  } catch (err) {
    toast_("資格管理プルダウン復旧でエラー: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function handleQualificationManagementEdit_(e) {
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "資格管理") return false;
  if (e.range.getRow() <= 1) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  const row = e.range.getRow();
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  const dueCol = headers.indexOf("更新期限") + 1;
  const statusCol = headers.indexOf("状態") + 1;
  const noticeCol = headers.indexOf("通知") + 1;
  const noteCol = headers.indexOf("備考") + 1;

  if (editedHeader === "区分") {
    // 区分を変えた時は、その行の資格名プルダウン候補を区分別・資格マスタ表示順に作り直す。
    // 既存の資格名が新しい区分と食い違う場合だけ空欄へ戻す。
    // 取得区分・保有状況は手入力を尊重し、ここでは勝手に補完しない。
    syncQualificationManagementNameAfterCategoryEdit_(sheet, row);
    return true;
  }

  if (editedHeader !== "資格名") return false;
  if (nameCol <= 0) return false;

  const raw = String(sheet.getRange(row, nameCol).getValue() || "").trim();
  if (!raw || raw === CLEAR_LABEL) {
    e.range.clearContent();
    return true;
  }

  const canonical = canonicalizeQualificationNameWithAlias_(raw);
  let name = canonical.name;
  if (!name) return true;

  const specialName = normalizeSpecialEducationQualificationName_(name);
  const nameChanged = canonical.changed || specialName !== name;
  name = specialName;
  if (nameChanged) sheet.getRange(row, nameCol).setValue(name);

  const masterCategory = getQualificationMasterCategoryByName_(name) || canonical.category;
  if (categoryCol > 0 && masterCategory) {
    // 資格名を選んだ時は、資格マスタを正として区分だけ補正する。
    // v9.7.48e以降: 取得区分・保有状況は手入力欄なので、ここでは自動入力しない。
    sheet.getRange(row, categoryCol).setValue(masterCategory);
    setQualificationManagementNameValidationForRow_(sheet, row, false);
  } else if (!masterCategory && noteCol > 0) {
    const noteCell = sheet.getRange(row, noteCol);
    const note = String(noteCell.getValue() || "").trim();
    const msg = "要確認：マスタ未登録または表記ゆれ疑い「" + name + "」";
    if (!note.includes(msg)) noteCell.setValue(note ? note + " / " + msg : msg);
  }

  if (noticeCol > 0) {
    const due = dueCol > 0 ? sheet.getRange(row, dueCol).getValue() : "";
    const status = statusCol > 0 ? sheet.getRange(row, statusCol).getValue() : "";
    sheet.getRange(row, noticeCol).setValue(getNoticeText(due, status));
  }

  return true;
}

function getDefaultQualificationAcquisition_(category, name) {
  const c = normalizeQualificationCategoryValue_(category);
  if (c === "免許") return "免許";
  if (c === "技能講習" || c === "特別教育" || c === "安全教育") return "修了";
  if (isEngineerPossessionSupportedQualification_(name)) return "技士";
  return "保有";
}



function sortQualificationManagementByPaperOrderLight() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    clearSettingsCache_();
    repairQualificationMasterPaperOrderFast_();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
    if (!sheet) {
      toast_("資格管理シートが見つかりません。");
      return;
    }
    const count = sortQualificationManagementRowsOnlyByOwnerMaster_(sheet);
    try { cleanupQualificationManagementFalseDisplayFast_(); } catch (e) {}
    try { refreshQualificationManagementCheckboxesFast_(); } catch (e) {}
    try { refreshQualificationManagementDropdownsLight_(sheet); } catch (e) {}
    try { applyQualificationManagementMinimalDisplay_(sheet); } catch (e) {}
    try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
    toast_("資格管理を紙マスタ順に軽量並べ替えしました（" + count + "件）。");
  } catch (err) {
    toast_("資格管理の軽量並べ替えでエラー: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function sortQualificationManagementAfterEditLight_(e) {
  // p25で追加したonEdit用軽量ソート。
  // p27以降は、入力行が移動してプルダウンが消えたように見えるため、
  // 通常のonEditからは呼ばない。必要時の保守/検証用として残す。
  if (!e || !e.range) return 0;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "資格管理") return 0;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return 0;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  const sortHeaders = ["所有者", "区分", "資格名", "取得区分", "保有状況", "取得日", "更新期限", "状態"];
  if (!sortHeaders.includes(editedHeader)) return 0;

  try {
    return sortQualificationManagementRowsOnlyByOwnerMaster_(sheet);
  } catch (err) {
    console.log("sortQualificationManagementAfterEditLight_ skip: " + err.message);
    return 0;
  }
}

function sortQualificationManagementRowsOnlyByOwnerMaster_(targetSheet) {
  // p25:
  // sortQualificationManagementByOwner_ は表示調整・全入力規則・チェックボックスまで実行するため重い。
  // onEdit用には、値の並び替え＋C列資格名プルダウン再設定だけに絞る。
  const sheet = targetSheet || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
  if (!sheet) return 0;

  const headers = getHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 3 || !headers.length) return 0;

  const ownerIndex = headers.indexOf("所有者");
  const categoryIndex = headers.indexOf("区分");
  const qualificationIndex = headers.indexOf("資格名");
  const acquisitionIndex = headers.indexOf("取得区分");
  const holdingIndex = headers.indexOf("保有状況");
  const dueIndex = headers.indexOf("更新期限");
  const nameCol = qualificationIndex + 1;

  if (qualificationIndex < 0) return 0;

  const rowCount = lastRow - 1;
  const range = sheet.getRange(2, 1, rowCount, headers.length);
  const values = range.getValues();
  const dataRows = [];
  const emptyRows = [];
  const staffOrder = getStaffOrderMap_();
  const masterOrderMap = getQualificationMasterOrderMap_();

  values.forEach(row => {
    if (isBusinessDataRow_(headers, row, "資格管理")) {
      dataRows.push(row);
    } else {
      emptyRows.push(headers.map(() => ""));
    }
  });

  dataRows.sort((a, b) => {
    const ao = normalizeTextForKey_(ownerIndex >= 0 ? a[ownerIndex] : "");
    const bo = normalizeTextForKey_(ownerIndex >= 0 ? b[ownerIndex] : "");
    const staffA = getStaffSortOrder_(ao, staffOrder);
    const staffB = getStaffSortOrder_(bo, staffOrder);
    if (staffA !== staffB) return staffA - staffB;
    if (ao !== bo) return String(ao).localeCompare(String(bo), "ja");

    const ac = String(categoryIndex >= 0 ? a[categoryIndex] || "" : "").trim();
    const bc = String(categoryIndex >= 0 ? b[categoryIndex] || "" : "").trim();
    const aq = String(qualificationIndex >= 0 ? a[qualificationIndex] || "" : "").trim();
    const bq = String(qualificationIndex >= 0 ? b[qualificationIndex] || "" : "").trim();
    const aqOrder = getQualificationMasterSortOrder_(ac, aq, masterOrderMap);
    const bqOrder = getQualificationMasterSortOrder_(bc, bq, masterOrderMap);
    if (aqOrder !== bqOrder) return aqOrder - bqOrder;

    const acOrder = getQualificationCategoryOrder_(ac);
    const bcOrder = getQualificationCategoryOrder_(bc);
    if (acOrder !== bcOrder) return acOrder - bcOrder;

    const aqKey = normalizeTextForKey_(aq);
    const bqKey = normalizeTextForKey_(bq);
    if (aqKey !== bqKey) return String(aqKey).localeCompare(String(bqKey), "ja");

    const aa = normalizeTextForKey_(acquisitionIndex >= 0 ? a[acquisitionIndex] : "");
    const ba = normalizeTextForKey_(acquisitionIndex >= 0 ? b[acquisitionIndex] : "");
    if (aa !== ba) return String(aa).localeCompare(String(ba), "ja");

    const ah = normalizeTextForKey_(holdingIndex >= 0 ? a[holdingIndex] : "");
    const bh = normalizeTextForKey_(holdingIndex >= 0 ? b[holdingIndex] : "");
    if (ah !== bh) return String(ah).localeCompare(String(bh), "ja");

    const ad = dueIndex >= 0 ? getDateTimeSortValue_(a[dueIndex]) : 0;
    const bd = dueIndex >= 0 ? getDateTimeSortValue_(b[dueIndex]) : 0;
    return ad - bd;
  });

  const sorted = dataRows.concat(emptyRows);
  if (!sorted.length) return 0;

  // C列の区分別入力規則が残ったままだと、行の入れ替え時に入力規則違反で止まる場合がある。
  // p34では「入力規則を消す」のではなく、一時的に全資格候補の緩いプルダウンで上書きする。
  // 途中停止しても、C列がプルダウン無しの状態で残りにくい。
  if (nameCol > 0) {
    try {
      const temporaryRule = buildQualificationNameValidationRuleForCategory_("", true);
      sheet.getRange(2, nameCol, sorted.length, 1).setDataValidation(temporaryRule);
    } catch (e) {
      console.log("sortQualificationManagementRowsOnlyByOwnerMaster_ temporary name validation skip: " + e.message);
    }
  }

  range.setValues(sorted);

  if (nameCol > 0) {
    refreshQualificationManagementDropdownsLight_(sheet);
  }

  ensureIdsForSheet_(sheet);
  setCheckboxesForDataRows(sheet);
  hideSystemColumnsForSheet_(sheet);

  return dataRows.length;
}

function getSpecialEducationQualificationNameMap_() {
  // p36:
  // 現場の紙資格では、高所作業系・締固系は技能講習ではなく特別教育として扱う。
  // 「高所作業」など短い表記で入力されても正式名へ寄せる。
  return {
    "車両系建設機械（締固）": "車両系建設機械（締固）",
    "車両系建設機械（締固め）": "車両系建設機械（締固）",
    "車両系建設機械締固": "車両系建設機械（締固）",
    "車両系建設機械締固め": "車両系建設機械（締固）",
    "締固め用機械": "車両系建設機械（締固）",
    "締固用機械": "車両系建設機械（締固）",
    "高所作業車": "高所作業車",
    "高所作業": "高所作業車",
    "高所作業車運転": "高所作業車",
    "高所作業車運転特別教育": "高所作業車",
    "高所作業車特別教育": "高所作業車"
  };
}

function getSafetyEducationQualificationNameMap_() {
  // p36:
  // 刈払い機系は安全教育へ統一する。
  // 会社内の表記ゆれ（刈払機、仮払い機など）も安全教育の「刈払い機」に寄せる。
  return {
    "刈払い機": "刈払い機",
    "刈払機": "刈払い機",
    "仮払い機": "刈払い機",
    "刈払い機取扱作業者": "刈払い機",
    "刈払機取扱作業者": "刈払い機",
    "刈払い機取扱作業従事者": "刈払い機",
    "刈払機取扱作業従事者": "刈払い機"
  };
}

function normalizeSpecialEducationQualificationName_(name) {
  // 既存関数名は互換のため残す。p36以降は特別教育だけでなく安全教育の表記ゆれもここで統一する。
  const clean = cleanQualificationName_(name);
  if (!clean) return "";
  const specialMap = getSpecialEducationQualificationNameMap_();
  if (specialMap[clean]) return specialMap[clean];
  const safetyMap = getSafetyEducationQualificationNameMap_();
  if (safetyMap[clean]) return safetyMap[clean];
  return clean;
}

function getQualificationCategoryOverrideByName_(name) {
  const clean = cleanQualificationName_(name);
  if (!clean) return "";
  const specialMap = getSpecialEducationQualificationNameMap_();
  if (specialMap[clean]) return "特別教育";
  const safetyMap = getSafetyEducationQualificationNameMap_();
  if (safetyMap[clean]) return "安全教育";
  return "";
}

function fixQualificationSpecialEducationCategories_() {
  // 既存の資格マスタ・資格管理に残っている旧区分を軽く補正する。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {masterFixed: 0, managementFixed: 0};

  const fixSheet_ = (sheetName, categoryHeader, nameHeader) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;
    const headers = getHeaders_(sheet);
    const categoryCol = headers.indexOf(categoryHeader) + 1;
    const nameCol = headers.indexOf(nameHeader) + 1;
    if (categoryCol <= 0 || nameCol <= 0) return 0;

    const rowCount = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
    let fixed = 0;

    values.forEach(row => {
      const beforeName = cleanQualificationName_(row[nameCol - 1]);
      const fixedName = normalizeSpecialEducationQualificationName_(beforeName);
      const overrideCategory = getQualificationCategoryOverrideByName_(beforeName);
      if (!overrideCategory) return;

      let changed = false;
      if (!normalizeQualificationCategoryValue_(row[categoryCol - 1]) && overrideCategory) {
        row[categoryCol - 1] = overrideCategory;
        changed = true;
      }
      if (fixedName && beforeName !== fixedName) {
        row[nameCol - 1] = fixedName;
        changed = true;
      }
      if (changed) fixed++;
    });

    if (fixed > 0) {
      sheet.getRange(2, 1, rowCount, headers.length).setValues(values);
    }
    return fixed;
  };

  result.masterFixed = fixSheet_("資格マスタ", "区分", "資格名");
  result.managementFixed = fixSheet_("資格管理", "区分", "資格名");

  try {
    const master = ss.getSheetByName("資格マスタ");
    if (master) sortQualificationMasterSheetByOrderSafe_(master);
  } catch (e) {
    console.log("fixQualificationSpecialEducationCategories_ master sort skip: " + e.message);
  }

  return result;
}


function refreshQualificationManagementCheckboxes() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const result = refreshQualificationManagementCheckboxes_();
    toast_(
      "資格管理の既読・個人確認チェックを強制再設定しました"
      + "（入力済み行 " + result.dataRows
      + " 行 / チェック列 " + result.checkboxCols
      + " 列 / チェックボックス " + result.checkboxCount
      + " 個）。GPT診断用シートを再作成してください。"
    );
  } finally {
    lock.releaseLock();
  }
}


function refreshQualificationManagementCheckboxes_() {
  // p26: 既存の公開メニュー名は維持し、中身だけ軽量版へ差し替える。
  return refreshQualificationManagementCheckboxesFast_();
}


/**
 * 資格管理専用。
 * 入力済み行だけに、既読 + 個人確認列のチェックボックスを強制的に入れる。
 *
 * 目的:
 * - GPT診断の「資格管理 入力済み行チェック不足: 860」を消す。
 * - 空行にはチェックボックスを出さない。
 * - 既存の TRUE/FALSE は可能な限り保持する。
 */
function forceQualificationManagementCheckboxes_(sheet, headers, members) {
  if (!sheet || !headers || !headers.length) {
    return {
      dataRows: 0,
      checkboxCols: 0,
      checkboxCount: 0
    };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      dataRows: 0,
      checkboxCols: 0,
      checkboxCount: 0
    };
  }

  const fixedMembers = Array.isArray(members) ? members : getPersonalMembers_();

  const checkboxHeaders = [READ_HEADER].concat(fixedMembers)
    .filter((name, index, arr) => name && arr.indexOf(name) === index);

  const checkboxCols = checkboxHeaders
    .map(header => headers.indexOf(header) + 1)
    .filter(col => col > 0);

  if (!checkboxCols.length) {
    return {
      dataRows: 0,
      checkboxCols: 0,
      checkboxCount: 0
    };
  }

  const rowCount = lastRow - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();

  const dataFlags = values.map(rowValues => {
    return isQualificationManagementDataRow_(headers, rowValues, fixedMembers);
  });

  const dataRowCount = dataFlags.filter(Boolean).length;

  checkboxCols.forEach(col => {
    const targetRange = sheet.getRange(2, col, rowCount, 1);
    const oldValues = targetRange.getValues();

    // いったん対象範囲の入力規則と中身を消す。
    // 空行にチェックボックスが残ることを防ぐため。
    targetRange.clearDataValidations();
    targetRange.clearContent();

    getTrueRuns_(dataFlags).forEach(run => {
      const startRow = run.start + 2;
      const length = run.length;
      const range = sheet.getRange(startRow, col, length, 1);

      // p19:
      // setDataValidation(requireCheckbox) だけだと、環境や表示更新の順番によって
      // FALSE 文字表示のまま残ることがあるため、insertCheckboxes() で明示的に
      // チェックボックス表示へ戻してから値を復元する。
      range.insertCheckboxes();
      const checkboxRule = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .build();
      range.setDataValidation(checkboxRule);

      const restoredValues = oldValues
        .slice(run.start, run.start + length)
        .map(row => [isTruthy_(row[0])]);

      range.setValues(restoredValues);
      range
        .setVerticalText(false)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(false);
    });
  });

  normalizeQualificationManagementCheckboxDisplay_(sheet, headers, checkboxCols, rowCount);

  return {
    dataRows: dataRowCount,
    checkboxCols: checkboxCols.length,
    checkboxCount: dataRowCount * checkboxCols.length
  };
}


/**
 * 資格管理の既読・個人確認列の表示を整える。
 * ヘッダーだけ縦書き、2行目以降は通常表示に戻す。
 * FALSE が「A L S」のように表示される問題を防ぐ。
 */
function normalizeQualificationManagementCheckboxDisplay_(sheet, headers, checkboxCols, rowCount) {
  if (!sheet || !headers || !checkboxCols || !checkboxCols.length) return;

  checkboxCols.forEach(col => {
    try {
      sheet.setColumnWidth(col, 44);

      sheet.getRange(1, col)
        .setVerticalText(true)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(false);

      if (rowCount > 0) {
        sheet.getRange(2, col, rowCount, 1)
          .setVerticalText(false)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setWrap(false);
      }
    } catch (e) {
      console.log("資格管理チェック列表示調整をスキップ: " + e.message);
    }
  });
}


function repairQualificationManagementFalseDisplay() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const cleanup = cleanupQualificationManagementFalseDisplay_();
    const check = refreshQualificationManagementCheckboxes_();
    toast_(
      "資格管理のFALSE表示を修復しました。"
      + " 通常列のFALSE掃除:" + cleanup.cleaned + "件"
      + " / チェックボックス:" + check.checkboxCount + "個"
    );
  } catch (err) {
    toast_("資格管理のFALSE表示修復でエラー: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}


function cleanupQualificationManagementFalseDisplay_() {
  // p26: 既存の公開メニュー名は維持し、中身だけ1回読み取り中心の軽量版へ差し替える。
  return cleanupQualificationManagementFalseDisplayFast_();
}


// p23互換用:
// p22で一部処理が旧関数名 cleanupQualificationFalseDisplay_ を呼んでいたため、
// 人ごとグループ化などが ReferenceError で止まることがあった。
// 正式名称は cleanupQualificationManagementFalseDisplay_ だが、旧呼び出しが残っても止まらないようラッパーを残す。
function cleanupQualificationFalseDisplay_() {
  return cleanupQualificationManagementFalseDisplay_();
}

/******************************
 * v9.7.48p26 資格管理軽量パック用の高速処理
 ******************************/
function repairQualificationMasterPaperOrderFast_() {
  // p37:
  // 資格マスタは「シートで手動変更した区分」を正とする。
  // 以前は軽量パック実行時に紙版標準マスタの区分へ毎回戻していたため、
  // 手で「技能講習 → 特別教育」「安全教育」などへ直しても更新で戻っていた。
  // ここでは、既存行は表示順・区分・資格名を原則上書きしない。
  // 不足している候補だけ紙版標準から追加し、空欄の補助列だけ埋める。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expectedHeaders = getSheetHeaders_()["資格マスタ"] || ["表示順", "区分", "資格名", "期限管理", "有効年数", "通知日前", "一覧表示", "備考", "資格マスタID"];
  let sheet = ss.getSheetByName("資格マスタ");
  if (!sheet) sheet = ss.insertSheet("資格マスタ");

  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    setupSheet_(sheet, expectedHeaders);
  }

  expectedHeaders.forEach(header => {
    const headers = getHeaders_(sheet);
    if (headers.indexOf(header) >= 0) return;
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue(header);
  });

  const headers = getHeaders_(sheet);
  const colMap = {};
  headers.forEach((h, i) => { if (h) colMap[h] = i + 1; });

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const rowCount = Math.max(lastRow - 1, 0);
  const values = rowCount > 0 ? sheet.getRange(2, 1, rowCount, headers.length).getValues() : [];
  const rows = values.slice();
  const rowByName = {};

  const isBlank_ = (value) => value === "" || value === null || value === undefined;
  const getCell_ = (row, header, fallbackCol) => row[(colMap[header] || fallbackCol) - 1];
  const setCell_ = (row, header, value) => {
    const col = colMap[header];
    if (col > 0) row[col - 1] = value;
  };

  rows.forEach((row, i) => {
    const name = normalizeSpecialEducationQualificationName_(getCell_(row, "資格名", 3));
    const key = normalizeTextForKey_(name);
    if (key && rowByName[key] === undefined) rowByName[key] = i;
  });

  let added = 0;
  let filled = 0;

  getQualificationPaperMasterRows_().forEach(item => {
    const paperName = normalizeSpecialEducationQualificationName_(item.name);
    if (!paperName) return;
    const key = normalizeTextForKey_(paperName);
    let index = rowByName[key];

    if (index === undefined) {
      const newRow = headers.map(() => "");
      rows.push(newRow);
      index = rows.length - 1;
      rowByName[key] = index;
      added++;
    }

    const row = rows[index];
    const currentCategory = normalizeQualificationCategoryValue_(getCell_(row, "区分", 2));
    const defaultCategory = normalizeQualificationCategoryValue_(item.category) || "その他";
    const effectiveCategory = currentCategory || defaultCategory;

    // 既存行は手動変更を優先する。空欄だけ補完する。
    const fillIfBlank_ = (header, value) => {
      const col = colMap[header];
      if (!col) return;
      if (isBlank_(row[col - 1])) {
        row[col - 1] = value;
        filled++;
      }
    };

    fillIfBlank_("表示順", Number(item.order) || 999999);
    fillIfBlank_("区分", effectiveCategory);
    fillIfBlank_("資格名", paperName);
    fillIfBlank_("期限管理", item.expire || "期限なし");
    fillIfBlank_("有効年数", item.years !== undefined ? item.years : "");
    fillIfBlank_("通知日前", item.noticeDays !== undefined ? item.noticeDays : (effectiveCategory === "免許" ? 90 : ""));
    fillIfBlank_("一覧表示", true);
    fillIfBlank_("備考", item.note || "紙の資格一覧表より");
    fillIfBlank_("資格マスタID", buildRecordId_("資格マスタ"));
  });

  if (rows.length) {
    const requiredRows = rows.length + 1;
    if (sheet.getMaxRows() < requiredRows) sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  try { sortQualificationMasterSheetByOrderSafe_(sheet); } catch (e) {
    console.log("repairQualificationMasterPaperOrderFast_ sort skip: " + e.message);
  }

  try { hideSystemColumnsForSheet_(sheet); } catch (e) {
    console.log("repairQualificationMasterPaperOrderFast_ hide system columns skip: " + e.message);
  }

  return {added: added, updated: filled};
}

function ensureQualificationPaperMasterRowsForLightPack_() {
  // 毎回、紙版マスタ全件をセル単位で上書き・書式調整すると時間超過しやすい。
  // 軽量パックでは、最低限必要な候補が無い場合だけ従来の補完処理を呼ぶ。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("資格マスタ");
  if (!sheet || sheet.getLastRow() < 2) {
    return ensureQualificationPaperMasterRowsUpsertSafe_();
  }

  const requiredNames = ["除雪講習", "刈払い機", "高所作業車", "車両系建設機械（締固）"];
  const headers = getHeaders_(sheet);
  const nameCol = headers.indexOf("資格名") + 1;
  if (nameCol <= 0) return ensureQualificationPaperMasterRowsUpsertSafe_();

  const values = sheet.getRange(2, nameCol, sheet.getLastRow() - 1, 1).getValues();
  const existing = {};
  values.forEach(row => {
    const name = normalizeSpecialEducationQualificationName_(row[0]);
    if (name) existing[normalizeTextForKey_(name)] = true;
  });

  const missing = requiredNames.some(name => !existing[normalizeTextForKey_(normalizeSpecialEducationQualificationName_(name))]);
  if (missing) return ensureQualificationPaperMasterRowsUpsertSafe_();

  return {added: 0, updated: 0};
}

function repairQualificationManagementCategoryNameSortFast_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
  if (!sheet) return {rows: 0, cleared: 0};

  try { ensureQualificationMasterNoticeColumn_(); } catch (e) {}

  // 刈払い機・高所作業車・締固の区分補正は、C列入力規則を貼る前に軽く行う。
  const specialFix = fixQualificationSpecialEducationCategoriesFast_();
  const result = applyQualificationManagementNameValidationsStrictByRow_(sheet);
  result.cleared = (result.cleared || 0) + (specialFix.managementFixed || 0);

  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}
  return result;
}

function fixQualificationSpecialEducationCategoriesFast_() {
  // p26:
  // 従来版は対象シートの全列を読み書きしていた。
  // ここでは「区分」「資格名」2列だけを読み、変更がある列だけ書き戻す。
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {masterFixed: 0, managementFixed: 0};

  const fixSheet_ = (sheetName, categoryHeader, nameHeader) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return 0;

    const headers = getHeaders_(sheet);
    const categoryCol = headers.indexOf(categoryHeader) + 1;
    const nameCol = headers.indexOf(nameHeader) + 1;
    if (categoryCol <= 0 || nameCol <= 0) return 0;

    const rowCount = sheet.getLastRow() - 1;
    const categoryRange = sheet.getRange(2, categoryCol, rowCount, 1);
    const nameRange = sheet.getRange(2, nameCol, rowCount, 1);
    const categories = categoryRange.getValues();
    const names = nameRange.getValues();

    let categoryChanged = false;
    let nameChanged = false;
    let fixed = 0;

    for (let i = 0; i < rowCount; i++) {
      const beforeName = cleanQualificationName_(names[i][0]);
      const fixedName = normalizeSpecialEducationQualificationName_(beforeName);
      const overrideCategory = getQualificationCategoryOverrideByName_(beforeName);
      if (!overrideCategory) continue;

      let changed = false;
      if (!normalizeQualificationCategoryValue_(categories[i][0]) && overrideCategory) {
        categories[i][0] = overrideCategory;
        categoryChanged = true;
        changed = true;
      }
      if (fixedName && beforeName !== fixedName) {
        names[i][0] = fixedName;
        nameChanged = true;
        changed = true;
      }
      if (changed) fixed++;
    }

    if (categoryChanged) categoryRange.setValues(categories);
    if (nameChanged) nameRange.setValues(names);
    return fixed;
  };

  result.masterFixed = fixSheet_("資格マスタ", "区分", "資格名");
  result.managementFixed = fixSheet_("資格管理", "区分", "資格名");
  return result;
}

function cleanupQualificationManagementFalseDisplayFast_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) {
    return {checked: 0, cleaned: 0};
  }

  const headers = getHeaders_(sheet);
  const members = getPersonalMembers_();
  const checkboxHeaders = new Set([READ_HEADER].concat(members).filter(Boolean));
  const ignoreHeaders = new Set(["所有者", "区分", "資格名", "取得区分", "保有状況", "コピー有無", "状態", "通知", "資格ID"]);

  const targetIndexes = [];
  headers.forEach((header, index) => {
    if (!header) return;
    if (checkboxHeaders.has(header)) return;
    if (ignoreHeaders.has(header)) return;
    targetIndexes.push(index);
  });

  if (!targetIndexes.length) return {checked: 0, cleaned: 0};

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  let checked = 0;
  let cleaned = 0;
  let changed = false;

  values.forEach(row => {
    targetIndexes.forEach(index => {
      checked++;
      const value = row[index];
      const text = String(value || "").trim().toUpperCase();
      if (value === true || value === false || text === "TRUE" || text === "FALSE") {
        row[index] = "";
        cleaned++;
        changed = true;
      }
    });
  });

  if (changed) {
    sheet.getRange(2, 1, rowCount, headers.length).setValues(values);
  }

  return {checked: checked, cleaned: cleaned};
}

function refreshQualificationManagementCheckboxesFast_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("資格管理");
  if (!sheet) return {dataRows: 0, checkboxCols: 0, checkboxCount: 0};

  let headers = getHeaders_(sheet);
  if (!headers.length) return {dataRows: 0, checkboxCols: 0, checkboxCount: 0};

  let members = [];
  try {
    clearSettingsCache_();
    members = getPersonalMembers_();

    // 列の追加・削除は必要最小限。毎回全入力シートは触らない。
    deleteObsoletePersonalCheckColumns_(sheet, members);
    ensurePersonalCheckColumns_(sheet, members);
  } catch (e) {
    console.log("資格管理の個人確認列整理をスキップ: " + e.message);
  }

  headers = getHeaders_(sheet);
  const result = forceQualificationManagementCheckboxesFast_(sheet, headers, members);

  try { normalizeQualificationManagementCheckboxDisplay_(sheet, headers, result.checkboxColsArray || [], Math.max(sheet.getLastRow() - 1, 0)); } catch (e) {}
  try { hideSystemColumnsForSheet_(sheet); } catch (e) {}

  return {
    dataRows: result.dataRows || 0,
    checkboxCols: result.checkboxCols || 0,
    checkboxCount: result.checkboxCount || 0
  };
}

function forceQualificationManagementCheckboxesFast_(sheet, headers, members) {
  if (!sheet || !headers || !headers.length || sheet.getLastRow() < 2) {
    return {dataRows: 0, checkboxCols: 0, checkboxCount: 0, checkboxColsArray: []};
  }

  const fixedMembers = Array.isArray(members) ? members : getPersonalMembers_();
  const checkboxHeaders = [READ_HEADER].concat(fixedMembers)
    .filter((name, index, arr) => name && arr.indexOf(name) === index);

  const checkboxCols = checkboxHeaders
    .map(header => headers.indexOf(header) + 1)
    .filter(col => col > 0);

  if (!checkboxCols.length) {
    return {dataRows: 0, checkboxCols: 0, checkboxCount: 0, checkboxColsArray: []};
  }

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet.getRange(2, 1, rowCount, headers.length).getValues();
  const dataFlags = values.map(rowValues => isQualificationManagementDataRow_(headers, rowValues, fixedMembers));
  const dataRowCount = dataFlags.filter(Boolean).length;
  const runs = getTrueRuns_(dataFlags);
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  checkboxCols.forEach(col => {
    const targetRange = sheet.getRange(2, col, rowCount, 1);
    const oldValues = targetRange.getValues();

    // 空行にチェックボックスを残さないため、入力規則は全解除。
    // 値は一括で、入力済み行だけ TRUE/FALSE、空行は空欄へ戻す。
    targetRange.clearDataValidations();

    const nextValues = oldValues.map((row, i) => {
      if (!dataFlags[i]) return [""];
      return [isTruthy_(row[0])];
    });
    targetRange.setValues(nextValues);

    runs.forEach(run => {
      sheet.getRange(2 + run.start, col, run.length, 1).setDataValidation(checkboxRule);
    });
  });

  return {
    dataRows: dataRowCount,
    checkboxCols: checkboxCols.length,
    checkboxCount: dataRowCount * checkboxCols.length,
    checkboxColsArray: checkboxCols
  };
}

function applyQualificationManagementMinimalDisplay_(sheet) {
  // p26:
  // 軽量パック用。罫線・条件付き書式・全セル整形は重いので行わない。
  // 必要最低限の見出し、列幅、ID非表示だけ整える。
  if (!sheet || sheet.getName() !== "資格管理" || sheet.getLastColumn() < 1) return;

  const headers = getHeaders_(sheet);
  if (!headers.length) return;

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const personal = getPersonalMembers_();
  const widths = {
    "所有者": 105,
    "区分": 110,
    "資格名": 300,
    "取得区分": 90,
    "保有状況": 90,
    "取得日": 105,
    "更新期限": 105,
    "コピー有無": 95,
    "状態": 105,
    "通知": 105,
    "要確認メモ": 180,
    "備考": 230,
    "資格ID": 55
  };

  try { sheet.setFrozenRows(1); } catch (e) {}

  try {
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#d9ead3")
      .setWrap(false);
  } catch (e) {}

  headers.forEach((header, index) => {
    if (!header) return;
    const col = index + 1;
    try {
      if (header === READ_HEADER || personal.includes(header)) {
        sheet.setColumnWidth(col, 42);
        sheet.getRange(1, col)
          .setVerticalText(true)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle")
          .setWrap(false);
      } else {
        sheet.setColumnWidth(col, widths[header] || 90);
      }
    } catch (e) {}
  });

  try { sheet.setRowHeight(1, 46); } catch (e) {}
  if (lastRow <= 500) {
    try { sheet.setRowHeights(2, lastRow - 1, 34); } catch (e) {}
  }

  try { hideColumnByHeaderSafely_(sheet, "資格ID"); } catch (e) {}
}



/**
 * 資格管理の入力済み行判定。
 * 既読・個人確認・通知・備考・資格IDだけでは入力済み扱いにしない。
 *
 * 基本は「所有者」「区分」「資格名」のどれかがあれば入力済み行とする。
 * これにより、資格管理の既読チェック復旧を安定させる。
 */
function isQualificationManagementDataRow_(headers, rowValues, members) {
  const ownerCol = headers.indexOf("所有者");
  const categoryCol = headers.indexOf("区分");
  const qualificationCol = headers.indexOf("資格名");

  const owner = ownerCol >= 0 ? String(rowValues[ownerCol] || "").trim() : "";
  const category = categoryCol >= 0 ? String(rowValues[categoryCol] || "").trim() : "";
  const qualification = qualificationCol >= 0 ? String(rowValues[qualificationCol] || "").trim() : "";

  if (owner || category || qualification) return true;

  const ignoreHeaders = new Set([
    READ_HEADER,
    "通知",
    "備考",
    "資格ID"
  ]);

  (members || []).forEach(name => ignoreHeaders.add(name));

  return headers.some((header, index) => {
    if (!header || ignoreHeaders.has(header)) return false;

    const value = rowValues[index];

    if (value === true) return true;
    if (value === false) return false;
    if (value instanceof Date) return true;

    return value !== "" && value !== null && value !== undefined;
  });
}


function repairQualificationDiagnosticHeaders() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const summary = repairQualificationDiagnosticHeaders_();
    toast_("資格系の診断対象ヘッダーを軽量修復しました: " + summary);
  } finally {
    lock.releaseLock();
  }
}

function repairQualificationDiagnosticHeaders_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = [];

  try {
    ensureQualificationBulkImportSheet_();
    logs.push("資格一括登録");
  } catch (e) {
    logs.push("資格一括登録失敗:" + e.message);
  }

  try {
    const duplicate = ss.getSheetByName("資格重複チェック");
    const desired = getSheetHeaders_()["資格重複チェック"];
    if (duplicate) {
      migrateQualificationAcquisitionStatusColumns_(duplicate);
      repairHeaderOrderPreserveData_(duplicate, desired);
      formatSheetBase_(duplicate, desired.length);
      logs.push("資格重複チェック");
    } else {
      createQualificationDuplicateCheckSheet_();
      logs.push("資格重複チェック作成");
    }
  } catch (e) {
    logs.push("資格重複チェック失敗:" + e.message);
  }

  try {
    // 自動出力なので再生成してもよい。所有者列なし仕様を反映する。
    createEmployeeQualificationListSheet_();
    logs.push("社員別資格一覧");
  } catch (e) {
    logs.push("社員別資格一覧失敗:" + e.message);
  }

  try {
    createQualificationHolderListSheet_();
    logs.push("資格別保有者一覧");
  } catch (e) {
    logs.push("資格別保有者一覧失敗:" + e.message);
  }

  return logs.join(" / ");
}

function repairGptDiagnosticTargetSheets_() {
  // GPT診断シート作成時専用。エラーで診断自体が止まらないように握りつぶして要約だけ返す。
  const logs = [];
  try {
    logs.push(repairQualificationDiagnosticHeaders_());
  } catch (e) {
    logs.push("資格系修復スキップ:" + e.message);
  }

  try {
    repairVehicleHistoryAndVehicleNameDropdowns_();
    logs.push("車検管理/車検履歴");
  } catch (e) {
    logs.push("車検履歴修復スキップ:" + e.message);
  }

  try {
    repairConstructionContractAmountAndTax_();
    logs.push("工事予定契約金額/状態");
  } catch (e) {
    logs.push("工事予定修復スキップ:" + e.message);
  }

  try {
    // 一覧スケジュールに古い個人確認列（山田/高橋/鈴木など）が残る誤診断を避けるため、
    // 設定シートの既読確認者を正として一覧・要確認だけ軽量再生成する。
    createScheduleList();
    createAlertList();
    logs.push("一覧スケジュール/要確認一覧");
  } catch (e) {
    logs.push("一覧再生成スキップ:" + e.message);
  }

  return logs.filter(Boolean).join(" / ");
}

function rebuildQualificationManagementDependentDropdowns() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const count = rebuildQualificationManagementDependentDropdowns_();
    toast_("資格管理のプルダウンを完全復旧しました（" + count + "行分）");
  } finally {
    lock.releaseLock();
  }
}

function rebuildQualificationManagementDependentDropdowns_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("資格管理");
  if (!sheet) return 0;
  ensureQualificationMasterNoticeColumn_();
  QUALIFICATION_DROPDOWN_HELPER_READY_ = false;
  ensureQualificationDropdownHelperSheet_();

  const headers = getHeaders_(sheet);
  if (!headers.length) return 0;
  const rowCount = getApplyRowCount_(sheet);
  const categoryCol = headers.indexOf("区分") + 1;
  const nameCol = headers.indexOf("資格名") + 1;
  const acquisitionCol = headers.indexOf("取得区分") + 1;
  const holdingCol = headers.indexOf("保有状況") + 1;
  const copyCol = headers.indexOf("コピー有無") + 1;
  const statusCol = headers.indexOf("状態") + 1;

  if (categoryCol > 0) setDropdownFromHelperRange_(sheet, categoryCol, "区分", false, rowCount);
  if (acquisitionCol > 0) setDropdownFromHelperRange_(sheet, acquisitionCol, "取得区分", false, rowCount);
  if (holdingCol > 0) setDropdownFromHelperRange_(sheet, holdingCol, "保有状況", false, rowCount);
  if (copyCol > 0) setDropdownFromHelperRange_(sheet, copyCol, "コピー有無", false, rowCount);
  if (statusCol > 0) setDropdownFromHelperRange_(sheet, statusCol, "状態", false, rowCount);

  if (nameCol > 0) {
    // p24: C列「資格名」は、B列「区分」に応じて行ごとに厳密設定する。
    applyQualificationManagementNameValidationsStrictByRow_(sheet);
  }

  return rowCount;
}

function getRunsForQualificationHelperHeader_(categoryValues, helperHeader) {
  const runs = [];
  let start = -1;
  for (let i = 0; i < categoryValues.length; i++) {
    const currentHeader = getQualificationHelperHeaderForCategory_(categoryValues[i]);
    const match = currentHeader === helperHeader;
    if (match && start < 0) start = i;
    if ((!match || i === categoryValues.length - 1) && start >= 0) {
      const end = match && i === categoryValues.length - 1 ? i + 1 : i;
      runs.push({start: start, length: end - start});
      start = -1;
    }
  }
  return runs;
}


function syncSelectedVehicleHistoryRowFromVehicleManagement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getActiveSheet();
  if (!historySheet || historySheet.getName() !== "車検履歴") {
    toast_("車検履歴シートで修正したい履歴行を選択してから実行してください");
    return;
  }

  const range = historySheet.getActiveRange();
  if (!range || range.getRow() <= 1) {
    toast_("車検履歴のデータ行を選択してください");
    return;
  }

  const vehicleSheet = ss.getSheetByName("車検管理");
  if (!vehicleSheet) {
    toast_("車検管理シートが見つかりません");
    return;
  }

  const row = range.getRow();
  const result = syncVehicleHistoryRowFromVehicleManagement_(historySheet, row, vehicleSheet);
  if (result.synced) {
    applyVehicleInspectionHistoryValidations_(historySheet);
    toast_("選択行の車検履歴を車検管理から再同期しました（" + result.detail + "）");
  } else {
    toast_("再同期できませんでした（" + result.reason + "）");
  }
}

function syncVehicleHistoryRowFromVehicleManagement_(historySheet, historyRow, vehicleSheet) {
  if (!historySheet || !vehicleSheet || historyRow <= 1) return {synced: false, reason: "対象行が不正です"};

  const hHeaders = getHeaders_(historySheet);
  const vHeaders = getHeaders_(vehicleSheet);
  const hValues = historySheet.getRange(historyRow, 1, 1, hHeaders.length).getValues()[0];
  const hObj = objectFromRow(hHeaders, hValues);

  const hName = String(hObj["車両名"] || "").trim();
  const hNumber = String(hObj["車番"] || "").trim();
  if (!hName && !hNumber) return {synced: false, reason: "履歴側の車両名・車番が空欄です"};
  if (vehicleSheet.getLastRow() < 2) return {synced: false, reason: "車検管理にデータがありません"};

  const vehicleRows = vehicleSheet.getRange(2, 1, vehicleSheet.getLastRow() - 1, vHeaders.length).getValues();
  let best = null;
  let bestScore = -1;

  vehicleRows.forEach(row => {
    const obj = objectFromRow(vHeaders, row);
    const vName = String(obj["車両名"] || "").trim();
    const vNumber = String(obj["車番"] || "").trim();
    let score = 0;
    if (hNumber && vNumber && hNumber === vNumber) score += 10;
    if (hName && vName && hName === vName) score += 5;
    if (score > bestScore) {
      bestScore = score;
      best = obj;
    }
  });

  if (!best || bestScore <= 0) return {synced: false, reason: "車検管理側に一致する車両がありません"};

  const updateMap = {
    "車両名": best["車両名"] || "",
    "車番": best["車番"] || "",
    "担当": getVehicleInspectionStaffForHistory_(best)
  };

  Object.keys(updateMap).forEach(header => {
    const col = hHeaders.indexOf(header) + 1;
    if (col > 0) historySheet.getRange(historyRow, col).setValue(updateMap[header]);
  });

  // 日付は履歴そのものなので、過去記録保護のため同期しない。
  return {synced: true, detail: (best["車両名"] || "車両") + " / 日付は変更なし"};
}

function applyVehicleInspectionHistoryValidations_(sheet) {
  if (!sheet || sheet.getName() !== "車検履歴") return;

  const headers = getHeaders_(sheet);
  const rowCount = getApplyRowCount_(sheet);
  if (!headers.length || rowCount < 1) return;

  // 旧版の担当者プルダウンが車検予定日列に残る事故を防ぐため、履歴シートは一度まとめて入力規則を消す。
  try {
    sheet.getRange(2, 1, rowCount, headers.length).clearDataValidations();
  } catch (e) {}

  ["更新日", "旧車検期限", "新車検期限", "車検予定日"].forEach(header => {
    const col = headers.indexOf(header) + 1;
    if (col > 0) setDateColumn(sheet, col);
  });

  const numberCol = headers.indexOf("車番") + 1;
  if (numberCol > 0) setTextColumn(sheet, numberCol);

  const staffCol = headers.indexOf("担当") + 1;
  if (staffCol > 0) setDropdown(sheet, staffCol, [CLEAR_LABEL, ...getStaffMembers_()]);

  const idCol = headers.indexOf("履歴ID") + 1;
  if (idCol > 0) setTextColumn(sheet, idCol);

  try {
    sheet.setColumnWidth(Math.max(headers.indexOf("車検予定日") + 1, 1), 145);
    if (staffCol > 0) sheet.setColumnWidth(staffCol, 180);
  } catch (e) {}
}

function cleanupVehicleHistoryStaffEmails_(sheet) {
  if (!sheet || sheet.getName() !== "車検履歴" || sheet.getLastRow() < 2) return 0;
  const headers = getHeaders_(sheet);
  const staffCol = headers.indexOf("担当") + 1;
  if (staffCol <= 0) return 0;

  const range = sheet.getRange(2, staffCol, sheet.getLastRow() - 1, 1);
  const values = range.getValues();
  let count = 0;
  const cleaned = values.map(row => {
    const value = String(row[0] || "").trim();
    if (isEmailLike_(value)) {
      count++;
      return [""];
    }
    return [row[0]];
  });
  if (count) range.setValues(cleaned);
  return count;
}

function repairVehicleHistoryValidationOnly() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const sheet = ensureVehicleInspectionHistorySheet_();
    const cleaned = cleanupVehicleHistoryStaffEmails_(sheet);
    applyVehicleInspectionHistoryValidations_(sheet);
    toast_("車検履歴の日付列・担当列の入力規則を修復しました" + (cleaned ? "（メールアドレス" + cleaned + "件を空欄化）" : ""));
  } finally {
    lock.releaseLock();
  }
}

function repairVehicleHistoryAndVehicleNameDropdowns() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    repairVehicleHistoryAndVehicleNameDropdowns_();
    toast_("車検履歴列、車検管理の新車検期限・担当列、車両名プルダウン、履歴日付列の入力規則を修復しました");
  } finally {
    lock.releaseLock();
  }
}

function repairVehicleHistoryAndVehicleNameDropdowns_() {
  ensureDefaultSettingRows_("車両名", DEFAULT_COMPANY_VEHICLES, "車検管理用");
  clearSettingsCache_();

  // 車検管理には「担当」列を正式追加する。
  ensureSheetHeadersPreserveData_("車検管理");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("車検管理");
  if (sheet) {
    const headers = getHeaders_(sheet);
    const vehicleNameCol = headers.indexOf("車両名") + 1;
    if (vehicleNameCol > 0) setDropdown(sheet, vehicleNameCol, [CLEAR_LABEL, ...getVehicleNames_()]);
    applyDataValidationByHeaders_(sheet, headers);
  }

  // 車検履歴は旧入力規則が残りやすいため、ヘッダー・入力規則・メールアドレス混入をまとめて修復する。
  ensureSheetHeadersPreserveData_("車検履歴");
  const historySheet = ensureVehicleInspectionHistorySheet_();
  cleanupVehicleHistoryStaffEmails_(historySheet);
  applyVehicleInspectionHistoryValidations_(historySheet);
  sortVehicleInspectionHistory_();
}

function repairConstructionContractAmountAndTax() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    repairConstructionContractAmountAndTax_();
    toast_("工事予定の契約金額を円表示にし、短い『税』プルダウンと連絡先の文字列固定を修復しました。状態に請求済みも追加済みです");
  } finally {
    lock.releaseLock();
  }
}

function repairConstructionContractAmountAndTax_() {
  // v9.7.48p:
  // 契約金額は円表示、税込/税抜/不明は短い「税」列で選択する。
  // 旧「税区分」列があれば「税」列へ移行する。
  ensureSheetHeadersPreserveData_("工事予定");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("工事予定");
  if (!sheet) return;

  const headers = getHeaders_(sheet);
  const amountCol = headers.indexOf("契約金額") + 1;
  const taxCol = headers.indexOf("税") + 1;
  const statusCol = headers.indexOf("状態") + 1;
  const contactCol = headers.indexOf("連絡先") + 1;

  normalizeConstructionAmountAndTaxRows_(sheet);

  if (amountCol > 0) {
    setConstructionAmountCurrencyColumn_(sheet, amountCol);
    try { sheet.setColumnWidth(amountCol, 145); } catch (e) {}
  }

  if (taxCol > 0) {
    setDropdown(sheet, taxCol, [CLEAR_LABEL, "税込", "税抜", "不明"]);
    try { sheet.setColumnWidth(taxCol, 70); } catch (e) {}
  }

  if (statusCol > 0) {
    setDropdown(sheet, statusCol, getStatusListForSheet_("工事予定"));
  }

  if (contactCol > 0) {
    setTextColumnAllCurrentRows_(sheet, contactCol);
    try { sheet.setColumnWidth(contactCol, 130); } catch (e) {}
  }

  updateNoticeColumnForSheet_(sheet);
  applyColorRules(sheet);
  createFilterSafely_(sheet, headers.length);
}

function normalizeConstructionAmountAndTaxRows_(sheet) {
  if (!sheet || sheet.getName() !== "工事予定" || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const amountCol = headers.indexOf("契約金額") + 1;
  const taxCol = headers.indexOf("税") + 1;
  if (amountCol <= 0 || taxCol <= 0) return;

  const rowCount = sheet.getLastRow() - 1;
  const amountValues = sheet.getRange(2, amountCol, rowCount, 1).getValues();
  const taxValues = sheet.getRange(2, taxCol, rowCount, 1).getValues();

  let changed = false;
  for (let i = 0; i < rowCount; i++) {
    const amount = amountValues[i][0];
    const tax = String(taxValues[i][0] || "").trim();
    const detectedTax = extractTaxLabelFromAmount_(amount);
    const normalizedAmount = normalizeConstructionAmountValue_(amount);

    if (!tax && detectedTax) {
      taxValues[i][0] = detectedTax;
      changed = true;
    }
    if (String(amount || "").trim() !== normalizedAmount) {
      amountValues[i][0] = normalizedAmount;
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(2, amountCol, rowCount, 1).setValues(amountValues);
    sheet.getRange(2, taxCol, rowCount, 1).setValues(taxValues);
  }
}

function repairRepairVendorMasterAndDropdown() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    const added = rebuildRepairVendorMasterFromEquipmentSheet_();
    const sorted = sortRepairVendorSettingsRows_();
    repairRepairVendorDropdown_();
    toast_("修理業者マスタを修復しました（追加" + added + "件 / 整理" + sorted + "件）。修理業者は候補外も手入力できます。");
  } finally {
    lock.releaseLock();
  }
}

function handleConstructionContactEdit_(e) {
  // p39: 工事予定の連絡先は「0128」などの先頭0を保持するため、文字列として扱う。
  // 既に数値化されて 128 になった過去値は、元の桁数が分からないため自動復元しない。
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "工事予定") return false;
  if (e.range.getRow() <= 1) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  if (editedHeader !== "連絡先") return false;

  const cell = e.range;
  cell.setNumberFormat("@").clearDataValidations();

  // e.value がユーザー入力文字列として渡る環境では、0128 をそのまま保存し直す。
  // undefined や既に数値化された値は無理にゼロ補完しない。
  const raw = e.value !== undefined && e.value !== null ? String(e.value) : "";
  if (raw && /^0\d+$/.test(raw)) {
    cell.setValue(raw);
  }

  return true;
}

function handleRepairVendorEdit_(e) {
  if (!e || !e.range) return false;
  const sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "備品修理管理") return false;
  if (e.range.getRow() <= 1) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return false;

  const headers = getHeaders_(sheet);
  const editedHeader = headers[e.range.getColumn() - 1];
  if (editedHeader !== "修理業者") return false;

  const vendor = String(e.range.getValue() || "").trim();
  if (vendor === CLEAR_LABEL) {
    e.range.clearContent();
    repairRepairVendorDropdown_();
    return true;
  }

  if (addRepairVendorToSettingsIfNeeded_(vendor)) {
    sortRepairVendorSettingsRows_();
    repairRepairVendorDropdown_();
    toast_("修理業者を設定シートへ追加しました: " + vendor);
  }
  return true;
}

function isValidRepairVendorName_(vendor) {
  const value = String(vendor || "").trim();
  if (!value || value === CLEAR_LABEL) return false;
  if (value.length < 2) return false;
  if (/SAMPLE|サンプル/i.test(value)) return false;
  return true;
}

function addRepairVendorToSettingsIfNeeded_(vendor) {
  const value = String(vendor || "").trim();
  if (!isValidRepairVendorName_(value)) return false;

  const sheet = ensureSettingsSheet_();
  const lastRow = sheet.getLastRow();
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 2).getValues() : [];
  const key = normalizeTextForKey_(value);
  const exists = existing.some(row => {
    const type = String(row[0] || "").trim();
    return (type === "修理業者" || type === "修理業者一覧") && normalizeTextForKey_(row[1]) === key;
  });
  if (exists) return false;

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 3).setValues([["修理業者", value, "備品修理管理用"]]);
  clearSettingsCache_();
  return true;
}

function rebuildRepairVendorMasterFromEquipmentSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("備品修理管理");
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const headers = getHeaders_(sheet);
  const vendorCol = headers.indexOf("修理業者") + 1;
  if (vendorCol <= 0) return 0;

  const values = sheet.getRange(2, vendorCol, sheet.getLastRow() - 1, 1).getValues();
  let count = 0;
  values.forEach(row => {
    if (addRepairVendorToSettingsIfNeeded_(row[0])) count++;
  });
  clearSettingsCache_();
  return count;
}

function sortRepairVendorSettingsRows_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
  if (!sheet || sheet.getLastRow() < 2) return 0;

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const others = [];
  const vendors = [];
  const seen = {};

  rows.forEach(row => {
    const type = String(row[0] || "").trim();
    const value = String(row[1] || "").trim();
    const note = row[2];

    if (type === "修理業者" || type === "修理業者一覧") {
      if (!isValidRepairVendorName_(value)) return;
      const key = normalizeTextForKey_(value);
      if (seen[key]) return;
      seen[key] = true;
      vendors.push(["修理業者", value, note || "備品修理管理用"]);
    } else {
      others.push(row);
    }
  });

  vendors.sort((a, b) => String(a[1]).localeCompare(String(b[1]), "ja"));

  const output = others.concat(vendors);
  sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  if (output.length) {
    sheet.getRange(2, 1, output.length, 3).setValues(output);
  }
  clearSettingsCache_();
  return vendors.length;
}

function repairRepairVendorDropdown_() {
  clearSettingsCache_();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("備品修理管理");
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  const vendorCol = headers.indexOf("修理業者") + 1;

  // 修理業者は「手入力 → 設定シートへ自動追加」が前提なので、候補外入力を許可する。
  // ここを setAllowInvalid(false) にすると、マスタ追加前に入力自体が弾かれてしまう。
  if (vendorCol > 0) setDropdownAllowInvalid(sheet, vendorCol, [CLEAR_LABEL, ...getRepairVendors_()]);
}

function repairEventScheduleColumns() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }
  try {
    repairEventScheduleColumns_();
    toast_("行事予定を、備考手前・内容列非表示の形に整理しました");
  } finally {
    lock.releaseLock();
  }
}

function hideEventContentColumnIfExists_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("行事予定");
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  const contentCol = headers.indexOf("内容") + 1;
  if (contentCol > 0) {
    try { sheet.hideColumns(contentCol); } catch (e) {}
  }
}



/******************************
 * v9.7.48p12 資格管理：列ズレ＋区分別プルダウン安全復旧
 * - F列の「取得日」見出しが空欄になる事故を修復
 * - 既存データはヘッダー名で並べ替えて保護
 * - 紙版資格マスタを補完し、安全教育「除雪講習」を復旧
 * - 区分ごとの資格名プルダウンを再構築
 ******************************/
function repairQualificationHeaderDropdownSortSafe() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    toast_("他の処理中です。時間をおいて再実行してください。");
    return;
  }

  try {
    const headerResult = repairQualificationManagementHeadersSafe_();
    const masterResult = ensureQualificationPaperMasterRowsUpsertSafe_();
    QUALIFICATION_DROPDOWN_HELPER_READY_ = false;
    ensureQualificationDropdownHelperSheet_();
    const dropdownRows = rebuildQualificationManagementDependentDropdowns_();
    const sortedRows = sortQualificationManagementByOwner_();
    createQualificationDuplicateCheckSheet_();
    createEmployeeQualificationListSheet_();
    createQualificationHolderListSheet_();
    formatQualificationManagementColors_();

    toast_(
      "資格管理の列ズレ＋区分別プルダウンを復旧しました。" +
      "列修復:" + headerResult.fixed +
      " / マスタ追加:" + masterResult.added +
      " / マスタ更新:" + masterResult.updated +
      " / プルダウン:" + dropdownRows + "行" +
      " / ソート:" + sortedRows + "件"
    );
  } finally {
    lock.releaseLock();
  }
}

function repairQualificationManagementHeadersSafe_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("資格管理");
  if (!sheet) sheet = ss.insertSheet("資格管理");

  const desiredHeaders = getSheetHeaders_()["資格管理"];
  if (!desiredHeaders || !desiredHeaders.length) return {fixed: 0};

  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    setupSheet_(sheet, desiredHeaders);
    return {fixed: 1};
  }

  removeFilter_(sheet);
  showAllColumns_(sheet);
  ensureColumns_(sheet, desiredHeaders.length);

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const readColCount = Math.max(sheet.getLastColumn(), desiredHeaders.length);
  const values = sheet.getRange(1, 1, lastRow, readColCount).getValues();
  const currentHeadersRaw = values[0].map(v => String(v || "").trim());
  const currentHeaders = currentHeadersRaw.slice();

  // よくある事故: E=保有状況、F=空欄、G=更新期限 になった場合、Fは「取得日」として扱う。
  desiredHeaders.forEach((header, index) => {
    if (!currentHeaders[index]) currentHeaders[index] = header;
  });

  // 旧列名の吸収。
  currentHeaders.forEach((header, index) => {
    if (header === "保有区分") currentHeaders[index] = "取得区分";
    if (header === "有効期限") currentHeaders[index] = "更新期限";
    if (header === "証明書コピー") currentHeaders[index] = "コピー有無";
  });

  const indexMap = {};
  currentHeaders.forEach((header, index) => {
    if (!header) return;
    if (indexMap[header] === undefined) indexMap[header] = index;
  });

  const out = [];
  out.push(desiredHeaders);
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const newRow = desiredHeaders.map((header, desiredIndex) => {
      const sourceIndex = indexMap[header];
      if (sourceIndex !== undefined) return row[sourceIndex];
      // F列空欄事故の保険。取得日は本来の位置に値があれば拾う。
      if (header === "取得日" && row[desiredIndex] !== undefined) return row[desiredIndex];
      return "";
    });
    out.push(newRow);
  }

  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), desiredHeaders.length).clearContent().clearDataValidations();
  sheet.getRange(1, 1, out.length, desiredHeaders.length).setValues(out);
  if (sheet.getMaxColumns() > desiredHeaders.length) {
    sheet.deleteColumns(desiredHeaders.length + 1, sheet.getMaxColumns() - desiredHeaders.length);
  }

  applyQualificationSingleSheetLightSettings_(sheet);
  applyQualificationManagementCompactDisplay_(sheet);
  setCheckboxesForDataRows(sheet);
  hideSystemColumnsForSheet_(sheet);
  return {fixed: 1};
}

function ensureQualificationPaperMasterRowsUpsertSafe_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const expectedHeaders = getSheetHeaders_()["資格マスタ"] || ["表示順", "区分", "資格名", "期限管理", "有効年数", "通知日前", "一覧表示", "備考", "資格マスタID"];
  let sheet = ss.getSheetByName("資格マスタ");
  if (!sheet) sheet = ss.insertSheet("資格マスタ");

  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    setupSheet_(sheet, expectedHeaders);
  }

  expectedHeaders.forEach(header => {
    const headers = getHeaders_(sheet);
    if (headers.indexOf(header) >= 0) return;
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()).setValue(header);
  });

  const headers = getHeaders_(sheet);
  const colMap = {};
  headers.forEach((h, i) => { if (h) colMap[h] = i + 1; });

  const rowByCategoryName = {};
  const rowByName = {};
  if (sheet.getLastRow() >= 2) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
    values.forEach((row, i) => {
      const category = normalizeQualificationCategoryValue_(row[(colMap["区分"] || 2) - 1]) || "その他";
      const name = cleanQualificationName_(row[(colMap["資格名"] || 3) - 1]);
      if (!name) return;
      const rowNo = i + 2;
      rowByCategoryName[normalizeTextForKey_(category) + "||" + normalizeTextForKey_(name)] = rowNo;
      if (!rowByName[normalizeTextForKey_(name)]) rowByName[normalizeTextForKey_(name)] = rowNo;
    });
  }

  let added = 0;
  let updated = 0;
  getQualificationPaperMasterRows_().forEach(item => {
    const category = normalizeQualificationCategoryValue_(item.category) || "その他";
    const name = cleanQualificationName_(item.name);
    if (!name) return;
    const key = normalizeTextForKey_(category) + "||" + normalizeTextForKey_(name);
    const nameKey = normalizeTextForKey_(name);
    let rowNo = rowByCategoryName[key] || rowByName[nameKey];
    if (!rowNo) {
      rowNo = sheet.getLastRow() + 1;
      added++;
    } else {
      updated++;
    }

    const valuesByHeader = {
      "表示順": Number(item.order) || 999999,
      "区分": category,
      "資格名": name,
      "期限管理": item.expire || "期限なし",
      "有効年数": item.years || "",
      "通知日前": item.noticeDays !== undefined ? item.noticeDays : (category === "免許" ? 90 : ""),
      "一覧表示": true,
      "備考": item.note || "紙の資格一覧表より",
      "資格マスタID": buildRecordId_("資格マスタ")
    };

    Object.keys(valuesByHeader).forEach(header => {
      const col = colMap[header];
      if (!col) return;
      const cell = sheet.getRange(rowNo, col);
      const current = cell.getValue();
      // p37: 手動変更した表示順・区分・資格名・一覧表示は上書きしない。
      // 空欄だけ補完する。
      if (current === "" || current === null || current === undefined) {
        cell.setValue(valuesByHeader[header]);
      }
    });

    rowByCategoryName[key] = rowNo;
    if (!rowByName[nameKey]) rowByName[nameKey] = rowNo;
  });

  try {
    sortQualificationMasterSheetByOrderSafe_(sheet);
    formatQualificationMasterSheetSafe_(sheet);
  } catch (e) {
    console.log("ensureQualificationPaperMasterRowsUpsertSafe_ format skip: " + e.message);
  }

  return {added: added, updated: updated};
}

function sortQualificationMasterSheetByOrderSafe_(sheet) {
  if (!sheet || sheet.getLastRow() < 3) return;
  const headers = getHeaders_(sheet);
  const orderCol = headers.indexOf("表示順");
  const categoryCol = headers.indexOf("区分");
  const nameCol = headers.indexOf("資格名");
  if (nameCol < 0) return;

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const rows = values.filter(row => cleanQualificationName_(row[nameCol]));
  const blanks = values.length - rows.length;
  rows.sort((a, b) => {
    const ao = orderCol >= 0 ? Number(a[orderCol]) || 999999 : 999999;
    const bo = orderCol >= 0 ? Number(b[orderCol]) || 999999 : 999999;
    if (ao !== bo) return ao - bo;
    const ac = categoryCol >= 0 ? normalizeQualificationCategoryValue_(a[categoryCol]) : "その他";
    const bc = categoryCol >= 0 ? normalizeQualificationCategoryValue_(b[categoryCol]) : "その他";
    const co = getQualificationCategoryOrder_(ac) - getQualificationCategoryOrder_(bc);
    if (co !== 0) return co;
    return String(a[nameCol] || "").localeCompare(String(b[nameCol] || ""), "ja");
  });
  const out = rows.concat(Array.from({length: blanks}, () => headers.map(() => "")));
  if (out.length) sheet.getRange(2, 1, out.length, headers.length).setValues(out);
}

function formatQualificationMasterSheetSafe_(sheet) {
  if (!sheet) return;
  const headers = getHeaders_(sheet);
  if (!headers.length) return;
  try {
    formatSheetBase_(sheet, headers.length);
    sheet.setFrozenRows(1);
    const widths = [70, 95, 300, 90, 80, 85, 85, 220, 150];
    widths.forEach((w, i) => { if (i < headers.length) sheet.setColumnWidth(i + 1, w); });
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const visibleCol = headers.indexOf("一覧表示") + 1;
    if (visibleCol > 0 && lastRow >= 2) sheet.getRange(2, visibleCol, lastRow - 1, 1).insertCheckboxes();
    const noticeCol = headers.indexOf("通知日前") + 1;
    if (noticeCol > 0 && lastRow >= 2) sheet.getRange(2, noticeCol, lastRow - 1, 1).setNumberFormat("0");
    createFilterSafely_(sheet, headers.length);
    // p29: 整形やフィルタ再作成後に資格マスタID列を必ず非表示へ戻す。
    hideSystemColumnsForSheet_(sheet);
  } catch (e) {
    console.log("formatQualificationMasterSheetSafe_ error: " + e.message);
  }
}













