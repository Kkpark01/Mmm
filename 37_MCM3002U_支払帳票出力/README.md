# MCM3002U 支払い明細出力指示：VB帳票出力の移植

**2026-09-25受領srcと全8ファイル一致（テキストは改行差・BOM除外）。再反映不要です。以下は作成時の記録。新しい共通仕様への適合確認は別途必要です。**

## 2026-09-24 Jev品質レビュー後の更新

Jev暫定7.3/10（修正前）。Codexの根拠確認により、Excel行数上限を原紙複製前に判定する変更と、出力中止・0件・送信完了のログを追加しました。Controller・ExcelServiceの2ファイルを同フォルダ内で更新。金額計算・テンプレート・画面操作は変更なし。Java388ファイルのコンパイルと50項目の回帰テスト成功。再採点は未実施。詳細：`検証記録/review_20260924/3002/quality/Codex確認結果.md`。以下の49項目・ブラウザ6項目は初版の検証記録です。

2026-09-24受領src基準。元の `mcm-web/src` は変更していません。35・36と対象ファイルの重複はありません。

## 反映方法・操作

このフォルダの `src` を会社側プロジェクトの `src` に同じ階層でコピーし、ビルド・再起動してください。SQLの新規関数作成は不要です。

メニュー→支払い明細出力指示→支払い月・区分を指定→「出力」。ZIP内に「発注確認書.xls」「支払い明細.xls」の2ファイルを格納します。各Excelのシートは取引先別です。VBの2回の保存操作をWebの1回のZIP保存に置き換えています。既存の検索・一覧・閉じるは維持。

## 対象ファイル（8ファイル）

既存5ファイルの修正：
- `src/main/java/com/daifuku/mcm/controller/Mcm3002uController.java`
- `src/main/java/com/daifuku/mcm/service/Mcm3002uService.java`
- `src/main/java/com/daifuku/mcm/repository/Mcm3002uRepository.java`
- `src/main/java/com/daifuku/mcm/form/Mcm3002uForm.java`
- `src/main/resources/templates/mcm3002u/index.html`

追加3ファイル：
- `src/main/java/com/daifuku/mcm/service/Mcm3002uExcelService.java`（2帳票の生成を1クラスに集約）
- `src/main/resources/templates/excel/MCM3001P_発注確認書テンプレート.xls`
- `src/main/resources/templates/excel/MCM3005P_支払い明細テンプレート.xls`

テンプレートは今回の添付を内容変更なしで配置。拡張子も実体もxlsです。依存ライブラリの追加なし。

## 変更内容とVB根拠

- 出力未実装を解消。Mcm3002uScreen.OutputButton_Click、Mcm3001pExcel、Mcm3005pExcel、各Constantを参照。全取引先分を出力し、0円・製番なしもVBの現行コード同様に除外しません。
- 旧Oracle関数 `database/30_func/mcm_fn_siharai.sql` の金額配分をJavaへ移植。見積期間の月割り、初月端数、終了日なしの場合の毎年端数、初回支払前の累計、次回支払前月までの累計を再現。支払い明細はVB画面のCIntと同じHALF_EVEN丸めを適用。
- 旧関数LCUR_KANE/LCUR_TUKIと同じ関連テーブル・条件をパラメーター付きSQLで取得。同一契約期間の金額は出力処理内で再利用。DBは読み取りのみ。
- 宛先・各項目の位置、取引先別シート、合計・件数、実績のSUM/COUNTIF、印刷反復行、罫線を移植。原紙の列幅・書式を使用し、記入前の原紙を複製して取引先間の明細混入を防止。
- 金額や行データは画面から受け取らず、出力時にDBから取得。不正年月・区分、0件、DB/テンプレート障害で出力せず入力条件を保持。両帳票生成後にダウンロードを開始。
- 共通JSは変更せず、このフォームに `data-mcm-no-progress` を付与。ファイル保存後に画面遷移待ち表示が残ることを防止。

## 検証結果

記録：`検証記録/review_20260924/3002/`

- 最新src＋35＋36＋37のJava388ファイルをコンパイル成功（compile-result.json / compile.log）。
- OutputTest.java：49項目成功。実Service/Controller/Thymeleaf/POIを使用。計算12ケース、同一期間の取得再利用、2帳票・全取引先・金額・合計・混入防止・数式・反復行、200行時のスタイル数、HTTPダウンロード、不正入力・0件・DB障害・テンプレート障害・入力保持等。DB部分は代替データで、SQL Serverでの実行ではありません。
- browser-test.cjs：Chromeで6項目成功。警告を閉じて再操作、ZIP保存、条件保持、検索継続、JSエラーなし。実Thymeleafの出力HTML＋実common.jsを使用し、通信応答はテストサーバーで代替。画面画像はscreen.png。
- 元srcの対象5ファイルは受領版とハッシュ一致。提供テンプレート2点もコピー元と一致。
- サンドボックスのクラス参照制限があったため、Java検証は許可後に制限外で成功。出力ログのsimulated DB/template errorは異常系テストで意図的に発生させたものです。

## 残る確認

会社SQL Serverへの接続・実データのVB帳票との金額照合、認証/権限/CSRFフィルタを含む実サーバーでの通し操作、Excel本体での印刷プレビューは未実施です。VBアプリそのものを実行して同値比較した結果ではなく、提供VB/SQLの仕様に基づくテストです。

SQL Serverの `MCM.MCM_TK_SIHARAI_V` にKAISI_DT/SYURYO_DTがあり、見積・契約・支払の参照テーブル/列が移行済みであることが前提です。今回の修正ではDBスキーマは作成・変更しません。取引先別のxls上限を超える件数はエラーとして中止します。

会社では同じ月・区分でVBとWebを出力し、取引先別件数・金額合計・端数を照合してください。年払/月払、複数取引先、0件、無効な月を確認した後、両Excelの印刷範囲・改ページを確認してください。
