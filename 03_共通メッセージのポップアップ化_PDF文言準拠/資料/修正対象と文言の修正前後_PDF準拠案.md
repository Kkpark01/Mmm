# ポップアップ化の対象画面・メッセージ文言 前後比較

2026-09-09／比較資料（未実装）

**画面上部・画面内に表示される色付きの通知を、ポップアップで表示する変更です。対象は17画面ID・20テンプレートです。** あわせて、対応が確認できた通知文言を添付「MCM_メッセージ一覧_全136件.pdf」に統一します。

最新版のmcm-web/srcと、反映用01（#298）、02（#294）の上書き後に相当するソースを照合しました。今回は資料作成のみで、ソース・反映用フォルダを変更していません。

## 対象画面

「文言変更」は対応するPDF文言が明確なもの、「一致」は既に同じものです。「要整理」は文言の対応未確定または判定条件の差があるものです。いずれの画面も、残っている画面内通知のポップアップ化が対象です。

| 画面ID | 画面名 | テンプレート数 | 文言変更 | 一致 | 要整理 |
|---|---|---:|---:|---:|---:|
| MCM0021U | プラント付替え | 1 | 1 | 9 | 7 |
| MCM1001U | 見積依頼作成検索 | 1 | 2 | 0 | 0 |
| MCM1002U | 需要家見積作成／見積依頼機器選定（2段階） | 2 | 2 | 0 | 4 |
| MCM1003U | 取引先見積・契約一覧 | 1 | 6 | 0 | 8 |
| MCM1004U | 需要家契約一覧 | 1 | 3 | 0 | 3 |
| MCM1005U | 取引先契約内容 | 1 | 2 | 0 | 6 |
| MCM1006U | 取引先契約機器選定（2段階） | 2 | 3 | 1 | 2 |
| MCM1007U | 単価精査 | 1 | 0 | 0 | 2 |
| MCM1008U | プラント選択 | 1 | 0 | 0 | 2 |
| MCM1009U | パック契約・解約 | 1 | 3 | 0 | 1 |
| MCM1010U | 店舗直契約・解約 | 1 | 3 | 0 | 1 |
| MCM2001U | 保守見積作成検索 | 1 | 5 | 0 | 3 |
| MCM2002U | カスタマー見積機器選定 | 1 | 1 | 0 | 6 |
| MCM2003U | 店舗見積内容基本設定 | 1 | 2 | 1 | 5 |
| MCM2005U | 店舗見積ブランド詳細設定 | 1 | 1 | 0 | 2 |
| MCM2006U | ユーザ契約内容 | 1 | 1 | 0 | 2 |
| MCM2007U | ユーザ契約内容変更（2段階） | 2 | 3 | 0 | 5 |

## この表の読み方

- **変更**：変更後欄のPDF文言へそろえる案です。
- **一致**：文言はそのまま。画面内表示の経路はポップアップへ変更します。0021Uの既存エラーポップアップも文言照合のため併記しています。
- **条件差あり／対応要確認**：PDF・VBの対応文言を載せていますが、現Javaの判定条件に差があるため、単純置換の確定対象とは分けています。
- **未確定**：対応文言がない、または使う場面が一致するか未確定です。変更後欄は現行文言の仮置きです。**PDF準拠済みという意味ではありません。**

通知単位で108行（画面間の同じ文言は別行、同一画面内の重複定義は統合）。文言変更38行・一致11行・条件差／対応要確認5行・未確定54行です。これはPDF136件のうちの採用数や、修正ファイル数ではありません。

表中の {詳細} は実行時の例外内容、{行番号}・{タブ名}・{個体名} は変動部分です。PDFの {0}・{1} は実データに置換します。PDFの vbCrLf は文字として表示せず、改行として扱います。全角／半角、句読点、「下さい／ください」もPDFの表記を使用します。

## 画面別の修正前・修正後

### MCM0021U プラント付替え

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 一致 | 検索条件は、1項目以上選択して下さい。 | 検索条件は、1項目以上選択して下さい。 | MSG_0001・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:149>) |
| 一致 | 検索結果が1件も存在しません。 | 検索結果が1件も存在しません。 | MSG_0002・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:152>) |
| 一致 | 付替え対象が検索されていません。 | 付替え対象が検索されていません。 | MSG_0124・p.3 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:157>) |
| 一致 | 付替え元が１件もチェックされていません。 | 付替え元が１件もチェックされていません。 | MSG_0130・p.3 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:160>) |
| 未確定 | 付替え先が１件もチェックされていません。 | 付替え先が１件もチェックされていません。 | 対応未確定 | PDFに「付替え先が…チェックされていません。」の対応文言なし。付替え元のMSG_0130へ置き換えない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:163>) |
| 一致 | 機器情報が存在しない為、付け替え元に設定することは出来ません。 | 機器情報が存在しない為、付け替え元に設定することは出来ません。 | MSG_0133・p.3 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:166>) |
| 一致 | 付替え元が設定されていません。 | 付替え元が設定されていません。 | MSG_0047・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:169>) |
| 一致 | 付替え先が設定されていません。 | 付替え先が設定されていません。 | MSG_0048・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:172>) |
| 一致 | 付替え元と付替え先は異なるデータを指定して下さい。 | 付替え元と付替え先は異なるデータを指定して下さい。 | MSG_0127・p.3 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:175>) |
| 一致 | 付替え先ブランドが選択されていないデータがあります。 | 付替え先ブランドが選択されていないデータがあります。 | MSG_0086・p.2 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:178>) |
| 未確定 | 付替え対象を1件以上チェックしてください。 | 付替え対象を1件以上チェックしてください。 | 対応未確定 | MSG_0037が候補だが、VBの同じ付替え対象選択条件との対応は未確定。現行文言を仮置き。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:181>) |
| 変更 | ブランド[%s]の機器構成[%s]でセット数が一致しないため、付替えできません。 | ブランド”{0}”に紐付く個体の中で機器構成”{1}”を形成できない設定があるため、付替えすることができません。<br>機器構成”{1}”の内容を確認してください。 | MSG_0100・p.3 | VBのMcm0021uScreen.vb:541でMSG_0100を使用。{0}=ブランド名、{1}=機器構成名。改行・引用符もPDFに合わせる。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:184>) |
| 未確定 | 付替え対象が変更されています。再検索して付替え元を設定し直してください。 | 付替え対象が変更されています。再検索して付替え元を設定し直してください。 | 対応未確定 | Web側の再実行・対象変更検知。PDFに同じ対応文言なし。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:193>) |
| 未確定 | 付替えが完了しました。 | 付替えが完了しました。 | 対応未確定 | 単独の完了通知に対応する文言なし。MSG_0098はクリア確認を兼ねるため、そのまま置換しない。後述の確認との二重表示を実装時に整理する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm0021uConstants.java:196>) |
| 未確定 | 付替え元に設定しました。 | 付替え元に設定しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/修正コード/01_298_DTS連携エラー一覧_サポートIDリンク/src/main/java/com/daifuku/mcm/controller/Mcm0021uController.java:291>) |
| 未確定 | 付替え先に設定しました。 | 付替え先に設定しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/修正コード/01_298_DTS連携エラー一覧_サポートIDリンク/src/main/java/com/daifuku/mcm/controller/Mcm0021uController.java:359>) |
| 未確定 | 付替えは完了しましたが、検索結果を更新できませんでした。検索を再実行してください。 | 付替えは完了しましたが、検索結果を更新できませんでした。検索を再実行してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/修正コード/01_298_DTS連携エラー一覧_サポートIDリンク/src/main/java/com/daifuku/mcm/controller/Mcm0021uController.java:478>) |

### MCM1001U 見積依頼作成検索

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 変更 | 該当するデータがありません。 | 検索結果が1件も存在しません。 | MSG_0002・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1001uConstants.java:75>) ／ [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1001uController.java:95>) |
| 変更 | 検索条件を入力してください。 | 検索条件は、1項目以上選択して下さい。 | MSG_0001・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1001uConstants.java:78>) ／ [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1001uService.java:46>) |

### MCM1002U 需要家見積作成／見積依頼機器選定（2段階）

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 変更 | 登録しました | 登録を完了しました。 | MSG_0008・p.1 | VBのMcm1002u2Screen.vb:1319で登録完了にMSG_0008。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1002uController.java:121>) |
| 未確定 | 登録に失敗しました: {詳細} | 登録に失敗しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1002uController.java:123>) |
| 未確定 | 依頼書を発行しました | 依頼書を発行しました | 対応未確定 | PDFに「依頼書発行完了」の単独文言なし。VBの発行処理内で登録を実行した場合はMSG_0008だが、発行のみの場合まで同一と決めない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1002uController.java:154>) |
| 未確定 | 依頼書発行に失敗しました: {詳細} | 依頼書発行に失敗しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1002uController.java:156>) |
| 変更 | 行{行番号}: 保守契約時間帯（8H/24H）のいずれかを選択してください | 保守契約時間帯を選択してください。 | MSG_0024・p.1 | VBのMcm1002u2Screen.vb:820でMSG_0024。行番号の接頭辞は本文から外す。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1002uService.java:131>) |
| 未確定 | 行{行番号}: 保守契約時間帯は8Hまたは24Hのいずれか一方を選択してください | 行{行番号}: 保守契約時間帯は8Hまたは24Hのいずれか一方を選択してください | 対応未確定 | 8Hと24Hの両方選択のエラー。JavaコメントのVBメッセージIDはMSG_0145で、今回のPDF（0136まで）に含まれない。未選択用MSG_0024へまとめない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1002uService.java:134>) |

### MCM1003U 取引先見積・契約一覧

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | 見積を破棄しました。 | 見積を破棄しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:199>) |
| 未確定 | 見積破棄処理でエラーが発生しました。 | 見積破棄処理でエラーが発生しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:209>) |
| 未確定 | 契約を破棄しました。 | 契約を破棄しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:250>) |
| 未確定 | 契約破棄処理でエラーが発生しました。 | 契約破棄処理でエラーが発生しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:260>) |
| 未確定 | 削除しました。 | 削除しました。 | 対応未確定 | MSG_0073は添付ファイル削除の完了。行削除の完了には流用しない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:314>) |
| 未確定 | 削除処理でエラーが発生しました。 | 削除処理でエラーが発生しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:318>) |
| 未確定 | 出力対象のデータがありません。 | 出力対象のデータがありません。 | 対応未確定 | 出力対象なし。検索結果0件のMSG_0002とは場面が異なるため未確定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:357>) |
| 未確定 | プラント情報が取得できませんでした。 | プラント情報が取得できませんでした。 | 対応未確定 | 遷移先情報の取得失敗。PDFのMSG_0003は説明文形式であり、表示文言として採用するか未確定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1003uController.java:582>) |
| 変更 | 検索条件を入力してください。 | 検索条件は、1項目以上選択して下さい。 | MSG_0001・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1003uService.java:105>) |
| 変更 | 依頼日（開始）の日付形式が不正です。(yyyy/MM/dd) | 開始日付の書式を指定して下さい。(YYYY/MM/DD) | MSG_0058・p.2 | 開始日付の書式エラーとして対応。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1003uService.java:118>) |
| 変更 | 依頼日（終了）の日付形式が不正です。(yyyy/MM/dd) | 終了日付の書式を指定して下さい。(YYYY/MM/DD) | MSG_0059・p.2 | 終了日付の書式エラーとして対応。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1003uService.java:121>) |
| 変更 | 該当するデータがありません。 | 検索結果が1件も存在しません。 | MSG_0002・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1003uService.java:140>) |
| 変更 | 契約が存在するため、見積の破棄はできません。 | 店舗への見積依頼 又は 取引先との契約として使用されている為、破棄する事が出来ません。 | MSG_0025・p.1 | 既存Javaの変換元コメントもMSG_0025を指定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1003uService.java:194>) |
| 変更 | 支払処理済みの期間があるため、契約の破棄はできません。 | 取引先への支払が発生している為、破棄する事が出来ません。 | MSG_0026・p.1 | 既存Javaの変換元コメントもMSG_0026を指定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1003uService.java:274>) |

### MCM1004U 需要家契約一覧

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 変更 | 登録しました | 登録を完了しました。 | MSG_0008・p.1 | VBのMcm1004uScreen.vb:383でMSG_0008。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1004uController.java:73>) |
| 未確定 | 期間の開始日は必須です | 期間の開始日は必須です | 対応未確定 | 必須項目用MSG_0054が候補。項目名付きの現行表示に対するVB共通入力検証の文言対応は未確定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1004uService.java:132>) |
| 変更 | {タブ名}タブの終了日は必須です | {0}タブの終了日は必須項目です。 | MSG_0116・p.3 | VBのMcm1004uScreen.vb:531。{0}=対象タブ名。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1004uService.java:134>) |
| 条件差あり | 開始日は終了日以前に設定してください | 開始日は終了日よりも前の日付で入力して下さい。 | MSG_0027・p.1 | VB:685は終了日≦開始日をエラーとする。現Javaは同日なら通過するため条件差がある。文言だけの修正とは分ける。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1004uService.java:137>) |
| 未確定 | 開始日は必須です | 開始日は必須です | 対応未確定 | 必須項目用MSG_0054が候補。開始日単独の対応は未確定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1004uService.java:147>) |
| 変更 | タブ追加には終了日が必須です | タブを追加するには期間の終了日の入力が必要です。 | MSG_0040・p.1 | VBのMcm1004uScreen.vb:250でMSG_0040。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1004uService.java:148>) |

### MCM1005U 取引先契約内容

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | 機器情報を更新しました。 | 機器情報を更新しました。 | 対応未確定 | 機器選定から戻った際の通知。登録完了MSG_0008と同じタイミングとは決めない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:79>) |
| 未確定 | 機器情報の更新に失敗しました: {詳細} | 機器情報の更新に失敗しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:81>) |
| 未確定 | 変更摘要日を入力してください。 | 変更摘要日を入力してください。 | 対応未確定 | 変更摘要日の未入力。MSG_0045には「先に」の意味があり、単純な未入力用に置換しない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:136>) |
| 未確定 | 変更摘要日の形式が正しくありません（yyyy/MM/dd）。 | 変更摘要日の形式が正しくありません（yyyy/MM/dd）。 | 対応未確定 | 変更摘要日の書式エラー。開始日用MSG_0058へ置換しない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:146>) |
| 未確定 | 対象期間が見つかりません。先に登録してください。 | 対象期間が見つかりません。先に登録してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:162>) |
| 変更 | 登録しました。 | 登録を完了しました。 | MSG_0008・p.1 | VBのMcm1005uScreen.vb:1003でMSG_0008。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:326>) |
| 変更 | 申請しました。 | 申請を完了しました。 | MSG_0121・p.3 | VBのMcm1005uScreen.vb:1261でMSG_0121。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1005uController.java:351>) |
| 条件差あり | 終了日が開始日より前になっています。 | 開始日は終了日よりも前の日付で入力して下さい。 | MSG_0027・p.1 | VB:718は終了日≦開始日をエラーとする。現Javaは同日なら通過するため条件差がある。判定条件の変更はこの比較表では決めない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm1005uService.java:161>) |

### MCM1006U 取引先契約機器選定（2段階）

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | 期間（開始日・終了日）を入力してください。 | 期間（開始日・終了日）を入力してください。 | 対応未確定 | 現コードは未入力と日付形式不正の両方で同じ文言を返す。VB:114の未入力通知はFWM_0001（期間）で、今回のPDFに含まれない。形式不正との分岐・共通メッセージの対応を確認してから決める。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1006uConstants.java:37>) |
| 条件差あり | 期間の開始日は終了日より前の日付を入力してください。 | 開始日は終了日よりも前の日付で入力して下さい。 | MSG_0027・p.1 | VB:121は開始日≦終了日を許容するが、現Javaは同日をエラーとする。文言はMSG_0027に対応するが条件差は別に扱う。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1006uConstants.java:39>) |
| 一致 | 期間は1年以内で設定してください。 | 期間は1年以内で設定してください。 | MSG_0123・p.3 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1006uConstants.java:41>) |
| 変更 | 見積を1件以上選択してください。 | １件もチェックされていません。 | MSG_0037・p.1 | VBのMcm1006u1Screen.vb:154でMSG_0037。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1006uConstants.java:43>) |
| 変更 | 選択した見積に重複した個体管理情報が含まれています。 | 選択した見積内に重複した機器が存在します。 | MSG_0038・p.1 | VBのMcm1006u1Screen.vb:187でMSG_0038。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1006uConstants.java:45>) |
| 変更 | 選択された見積には、出精値引きが含まれている為、見積の期間と契約の期間が一致する必要があります。期間を確認してください。 | 選択された見積には、出精値引きが含まれている為、見積の期間と契約の期間が一致する必要があります。<br>期間を確認してください。 | MSG_0079・p.2 | 本文は一致しているが「期間を確認してください。」の前の改行が欠けている。VB:336、385。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1006uConstants.java:47>) |

### MCM1007U 単価精査

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | パラメータが不足しています。前の画面からやり直してください。 | パラメータが不足しています。前の画面からやり直してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1007uController.java:73>) |
| 未確定 | 該当するデータがありません。 | 該当するデータがありません。 | 対応未確定 | 画面起動時のデータ取得0件。検索時MSG_0002と遷移先なしMSG_0003のどちらかを意味だけで決めない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1007uController.java:91>) |

### MCM1008U プラント選択

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | プラント情報が取得できませんでした。前の画面からやり直してください。 | プラント情報が取得できませんでした。前の画面からやり直してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1008uController.java:74>) |
| 未確定 | 該当するプラントデータがありません。 | 該当するプラントデータがありません。 | 対応未確定 | プラントデータ0件なのでMSG_0125が候補。ただしVBのこの画面での使用箇所は確認できず、対応要確認。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm1008uController.java:84>) |

### MCM1009U パック契約・解約

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 変更 | 検索条件を入力してください。 | 検索条件は、1項目以上選択して下さい。 | MSG_0001・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1009uConstants.java:77>) |
| 変更 | 検索結果が見つかりません。 | 検索結果が1件も存在しません。 | MSG_0002・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1009uConstants.java:78>) |
| 変更 | 支払が発生しているため、破棄できません。 | 取引先への支払が発生している為、破棄する事が出来ません。 | MSG_0026・p.1 | 支払発生済み契約の破棄エラー。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1009uConstants.java:80>) |
| 未確定 | 契約を破棄しました。 | 契約を破棄しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1009uConstants.java:81>) |

### MCM1010U 店舗直契約・解約

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 変更 | 検索条件を入力してください。 | 検索条件は、1項目以上選択して下さい。 | MSG_0001・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1010uConstants.java:79>) |
| 変更 | 検索結果が見つかりません。 | 検索結果が1件も存在しません。 | MSG_0002・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1010uConstants.java:81>) |
| 変更 | 支払が発生しているため、破棄できません。 | 取引先への支払が発生している為、破棄する事が出来ません。 | MSG_0026・p.1 | 支払発生済み契約の破棄エラー。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1010uConstants.java:85>) |
| 未確定 | 契約を破棄しました。 | 契約を破棄しました。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/constants/Mcm1010uConstants.java:87>) |

### MCM2001U 保守見積作成検索

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | 納入先・プラントを選択し直してください。 | 納入先・プラントを選択し直してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2001uController.java:190>) |
| 未確定 | 選択した見積を確認できません。再検索してください。 | 選択した見積を確認できません。再検索してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2001uController.java:199>) |
| 未確定 | 画面の情報を取得できませんでした。再検索してから、もう一度操作してください。 | 画面の情報を取得できませんでした。再検索してから、もう一度操作してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2001uController.java:290>) |
| 変更 | 検索条件を入力してください。 | 検索条件は、1項目以上選択して下さい。 | MSG_0001・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2001uService.java:157>) |
| 変更 | 該当するデータが見つかりません。 | 検索結果が1件も存在しません。 | MSG_0002・p.1 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2001uService.java:170>) |
| 変更 | 納入先データが存在しません。先に検索を実行してください。 | 納入先情報が存在しません。 | MSG_0067・p.2 | VBのMcm2001uScreen.vb:225でMSG_0067。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2001uService.java:275>) |
| 変更 | 選択された見積に機器明細の重複があります。 | 個体管理情報が重複しています。 | MSG_0056・p.2 | VBの同じ重複確認箇所（Mcm2001uScreen.vb:281）はMSG_0056。似たMSG_0038には置き換えない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2001uService.java:308>) |
| 変更 | プラントを選択してください。 | プラント情報にデータが存在しません。 | MSG_0125・p.3 | VBのMcm2001uScreen.vb:320、324では選択行なし／プラントIDなしにMSG_0125。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2001uService.java:320>) |

### MCM2002U カスタマー見積機器選定

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | 画面の情報を取得できませんでした。再検索してから、もう一度操作してください。 | 画面の情報を取得できませんでした。再検索してから、もう一度操作してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2002uController.java:49>) |
| 未確定 | セッションが切れました。再度操作してください。 | セッションが切れました。再度操作してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2002uController.java:178>) |
| 未確定 | 契約開始日を入力してください。 | 契約開始日を入力してください。 | 対応未確定 | 必須項目用MSG_0054が候補。契約開始日の共通入力検証との対応は未確定。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2002uService.java:389>) |
| 変更 | 契約開始日の形式が不正です。（yyyy/MM/dd） | 開始日付の書式を指定して下さい。(YYYY/MM/DD) | MSG_0058・p.2 | 契約開始日付の書式エラーとして対応。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2002uService.java:394>) |
| 未確定 | 契約時間帯を入力してください。 | 契約時間帯を入力してください。 | 対応未確定 | 数値の契約時間帯の未入力。8H/24Hの選択不足用MSG_0024と同じ条件とは決めない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2002uService.java:400>) |
| 未確定 | 契約時間帯は0～24の範囲で入力してください。 | 契約時間帯は0～24の範囲で入力してください。 | 対応未確定 | 現行は0～24、PDFのMSG_0012は1～24。許容範囲が異なるため文言だけを変更しない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2002uService.java:405>) |
| 未確定 | 契約時間帯は数値で入力してください。 | 契約時間帯は数値で入力してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2002uService.java:408>) |

### MCM2003U 店舗見積内容基本設定

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | データの読み込みに失敗しました: {詳細} | データの読み込みに失敗しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:62>) |
| 変更 | 見積を保存しました。 | 登録を完了しました。 | MSG_0008・p.1 | VBのMcm2003uScreen.vb:1138で保存・登録完了にMSG_0008。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:97>) |
| 未確定 | 保存中にエラーが発生しました: {詳細} | 保存中にエラーが発生しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:99>) |
| 未確定 | 見積を発行しました。 | 見積を発行しました。 | 対応未確定 | PDFに「見積発行完了」の単独文言なし。VB:712以降では、変更ありの発行は登録処理を通りMSG_0008、変更なしは別経路。全件を登録完了と断定しない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:119>) |
| 未確定 | 見積発行中にエラーが発生しました: {詳細} | 見積発行中にエラーが発生しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:121>) |
| 変更 | 申請が完了しました。 | 申請を完了しました。 | MSG_0121・p.3 | VBのMcm2003uScreen.vb:830でMSG_0121。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:137>) |
| 未確定 | 申請中にエラーが発生しました: {詳細} | 申請中にエラーが発生しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2003uController.java:141>) |
| 一致 | 見積資料を添付してください。 | 見積資料を添付してください。 | MSG_0071・p.2 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2003uService.java:80>) |

### MCM2005U 店舗見積ブランド詳細設定

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | データの読み込みに失敗しました: {詳細} | データの読み込みに失敗しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2005uController.java:83>) |
| 変更 | 保存しました。 | 登録を完了しました。 | MSG_0008・p.1 | VBのMcm2005uScreen.vb:242でMSG_0008。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2005uController.java:123>) |
| 未確定 | 保存中にエラーが発生しました: {詳細} | 保存中にエラーが発生しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2005uController.java:126>) |

### MCM2006U ユーザ契約内容

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | データの読み込みに失敗しました: {詳細} | データの読み込みに失敗しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2006uController.java:78>) |
| 変更 | 保存しました。 | 登録を完了しました。 | MSG_0008・p.1 | VBのMcm2006uScreen.vb:633でMSG_0008。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2006uController.java:163>) |
| 未確定 | 保存中にエラーが発生しました: {詳細} | 保存中にエラーが発生しました: {詳細} | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2006uController.java:168>) |

### MCM2007U ユーザ契約内容変更（2段階）

| 判定 | 修正前（現行） | 修正後（案） | PDF ID・ページ | 根拠・注意点 |
|---|---|---|---|---|
| 未確定 | 画面の情報を取得できませんでした。再検索してから、もう一度操作してください。 | 画面の情報を取得できませんでした。再検索してから、もう一度操作してください。 | 対応未確定 | PDFに同じ場面の対応文言を確認できないため、文言は現状維持案。表示方式はポップアップへ変更する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/controller/Mcm2007uController.java:42>) |
| 未確定 | 終了日を入力してください。 | 終了日を入力してください。 | 対応未確定 | 終了日未入力。タブの終了日用MSG_0116と同じ場面ではない。共通必須MSG_0054が候補。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:85>) |
| 変更 | 開始日の形式が正しくありません（yyyy/MM/dd）。 | 開始日付の書式を指定して下さい。(YYYY/MM/DD) | MSG_0058・p.2 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:94>) |
| 変更 | 終了日の形式が正しくありません（yyyy/MM/dd）。 | 終了日付の書式を指定して下さい。(YYYY/MM/DD) | MSG_0059・p.2 | 現在の判定内容に対応するPDF文言。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:99>) |
| 変更 | 終了日は開始日以降の日付を入力してください。 | 開始日は終了日よりも前の日付で入力して下さい。 | MSG_0027・p.1 | VB:191でMSG_0027。同日許容の判定はVBとJavaで同じ。PDFは「前の日付」と表記しているが、文言に合わせるために判定を厳しくしない。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:105>) |
| 条件差あり | 契約期間は1年以内で設定してください。 | 期間は1年以内で設定してください。 | MSG_0123・p.3 | VB:196は開始日＋1年－1日まで、Javaは開始日＋1年までを許容する。文言はMSG_0123に対応するが、上限日の差は別に扱う。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:111>) |
| 対応要確認 | ブランドを1件以上選択してください。 | 店舗見積を選択してください。 | MSG_0078・p.2 | VBは見積・ブランド双方の選択状態を確認してMSG_0078。現Javaはブランド選択数を確認するため、判定条件との対応を実装前に確認する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:121>) |
| 未確定 | 個体「{個体名}」は既にこの契約に登録されています。 | 個体「{個体名}」は既にこの契約に登録されています。 | 対応未確定 | 既存契約への追加対象の重複。選択した複数見積内の重複（MSG_0038）とは区別する。 [現行コード](<C:/Users/teras/Desktop/遠隔操作用/MCM/mcm-web/src/main/java/com/daifuku/mcm/service/Mcm2007uService.java:164>) |

## ポップアップのタイトル・ボタン

対応するPDFメッセージが決まったものは、本文だけでなくタイトル・ボタン・アイコンもその定義に合わせます。例えば登録完了は「完了／OK／情報」、検索0件は「検索エラー／OK／警告」です。次の表は今回参照したIDの定義です。

| ID | タイトル | ボタン | アイコン | PDF |
|---|---|---|---|---|
| MSG_0001 | 検索エラー | OK | 警告 | p.1 |
| MSG_0002 | 検索エラー | OK | 警告 | p.1 |
| MSG_0008 | 完了 | OK | 情報 | p.1 |
| MSG_0024 | 入力エラー | OK | 警告 | p.1 |
| MSG_0025 | 排他エラー | OK | 警告 | p.1 |
| MSG_0026 | 排他エラー | OK | 警告 | p.1 |
| MSG_0027 | 入力エラー | OK | 警告 | p.1 |
| MSG_0037 | 入力エラー | OK | 警告 | p.1 |
| MSG_0038 | 入力エラー | OK | 警告 | p.1 |
| MSG_0040 | 入力エラー | OK | 警告 | p.1 |
| MSG_0047 | 入力エラー | OK | 警告 | p.1 |
| MSG_0048 | 入力エラー | OK | 警告 | p.1 |
| MSG_0056 | 入力エラー | OK | 警告 | p.2 |
| MSG_0058 | 入力エラー | OK | 警告 | p.2 |
| MSG_0059 | 入力エラー | OK | 警告 | p.2 |
| MSG_0067 | 入力エラー | OK | 警告 | p.2 |
| MSG_0071 | 入力エラー | OK | 警告 | p.2 |
| MSG_0078 | 入力エラー | OK | 警告 | p.2 |
| MSG_0079 | 入力エラー | OK | 警告 | p.2 |
| MSG_0086 | 入力エラー | OK | 警告 | p.2 |
| MSG_0100 | 入力エラー | OK | 警告 | p.3 |
| MSG_0116 | 入力エラー | OK | 警告 | p.3 |
| MSG_0121 | 完了 | OK | 情報 | p.3 |
| MSG_0123 | 入力エラー | OK | 警告 | p.3 |
| MSG_0124 | 入力エラー | OK | 警告 | p.3 |
| MSG_0125 | 入力エラー | OK | 警告 | p.3 |
| MSG_0127 | 入力エラー | OK | 警告 | p.3 |
| MSG_0130 | 入力エラー | OK | 警告 | p.3 |
| MSG_0133 | 入力エラー | OK | 警告 | p.3 |

## 付替え完了の確認は別扱い

MCM0021Uには既に次の確認ポップアップがあります。これらはPDFと一致しています。

| 場面 | 現行文言 | PDF準拠の文言 | 定義 |
|---|---|---|---|
| 付替え実行前 | 付替えしてもよろしいですか？ | 同左 | MSG_0099・p.3／確認／OK・キャンセル |
| 付替え完了後のクリア確認 | 付替えを完了しました。付替え元、付替え先をクリアしてよろしいですか？ | 同左 | MSG_0098・p.3／確認／OK・キャンセル |

完了の緑帯「付替えが完了しました。」を、機械的にMSG_0098へ置き換えると、クリア確認を二重に表示する可能性があります。移行時は既存のクリア確認を1回表示し、OK／キャンセルの選択に応じた処理を維持する必要があります。帯の単独完了文言にはPDF内の別の対応定義がないため、上表では未確定として区別しています。

## 対応文言を無理に当てはめない箇所

- 「保存中にエラーが発生しました: {詳細}」などのWeb側の例外通知、セッション切れ、再検索の案内は、PDFに同じ通知文言がありません。表示方式はポップアップに移せますが、本文のPDF準拠については未確定です。
- 「契約を破棄しました。」「削除しました。」などの単独完了通知に、処理前の確認文や添付ファイル専用の完了文を流用しません。
- MCM2002Uの0～24という現在の許容範囲に、PDFの1～24という文言だけを当てることはしません。
- PDFには「検索結果から遷移したが、対応するデータが存在しない場合のエラー」（MSG_0003）など、説明文のような項目もあります。現在の具体的な案内文との対応は未確定として扱っています。

## 調査範囲と既存修正への影響

この一覧は、前回特定した17画面の画面内通知と、その発生元の文言を対象にしたソース照合です。既にポップアップの他画面の全メッセージや、PDF136件すべての発生条件を検証した一覧ではありません。実DBで全操作を実行した結果でもありません。操作前の既存確認、一覧が空の場合の表中説明、入力欄のラベルは一般の結果通知と分けています。

以前の一覧でMCM1004Uの通知場面を「登録・申請」と記載しましたが、今回の現行コントローラ再確認では登録・タブ追加等の通知が対象です。本資料では申請完了通知があるとは扱っていません。

文言まで変更する場合は、画面のHTMLと共通表示部品に加えて、文言を作っているJavaの定数・コントローラ・サービスも対象になります。20はHTMLの数で、総修正ファイル数ではありません。

#298の01フォルダとはMCM0021UのHTML・コントローラ等を共有するため、その修正済み内容を引き継いで作成する必要があります。#294の02フォルダ（MCM0010U）はこの17画面には含まれず、今回の対象とのファイルの重なりはありません。#294のMSG_0110の本文もPDFと一致しています。

参照資料：C:/Users/teras/Downloads/MCM_メッセージ一覧_全136件.pdf（全3ページ）
