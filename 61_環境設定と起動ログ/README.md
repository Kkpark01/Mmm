# 61：環境設定の説明と起動ログ

個人PCでのlocal／dev、検証サーバーへの配備、結合・排他テストの手順は[使い方ガイド](使い方ガイド.md)を参照してください。

基準：2026-10-02 19:45受領src（543ファイル）。利用者の指示に従い、**既定localを維持**し、起動時の環境・DB接続設定・添付保存先の表示と、チーム向けの設定コメントを追加した。既存5ファイル、新規0。受領srcを直接変更していない。

## 対象ファイル

```text
src/main/java/com/daifuku/mcm/McmApplication.java
src/main/resources/application.yml
src/main/resources/application-local.yml
src/main/resources/application-dev.yml
src/main/resources/application-prod.yml
```

- McmApplication：ApplicationRunnerのBeanで、起動完了時に使用プロファイル・設定上のDBサーバー／インスタンス／ポート／DB名・添付保存先をINFOログへ表示。JDBCドライバーでURLを解析し、認証情報・URL全体・解析例外の内容はログへ出さない。DB接続・添付保存先への実アクセスは追加していない。
- YAML4ファイル：コメントだけ変更。全設定値は受領版と一致。通常の個人開発はlocal、共有添付確認・検証サーバーはdev、本番はprodを指定。結合・排他テストは検証サーバーの同じURL・同じ版・別利用者で確認する運用を明記。
- localのDBは検証DB、添付はPC内という既存構成を維持。起動ログにPC内保存とdevによる共有確認の案内を追加。

## 反映方法・競合

このフォルダのsrc内5ファイルを、プロジェクトのsrcへ同じ相対位置で上書きし、ビルド・再起動する。追加Java・JAR・テンプレートは不要。既存の外部設定やEclipseのVM引数・環境変数があれば、そちらの指定が優先される。

**56との重複はMcmApplication.javaとapplication.ymlの2ファイル。旧56を一式反映しない。61は最新srcのlocal／dev／prod方式を維持する版で、56の旧切替方式や保存処理とは統合していない。61の後に旧56を上書きすると今回のログ・設定方式が失われる。** 56の保存処理を追加する場合は別途差分統合が必要。58・60の配布フォルダは削除済みで、最新srcの別修正は保持。

配備先に外部config/application.yml等がある場合、起動時に実際に使われる設定がそこにある可能性がある。チーム用コメントを外部YAMLへ反映する際は既存の配備値を維持する。

## 確認方法

1. 個人PCで環境指定なしに起動。Eclipseコンソールで `[MCM起動設定]` を検索し、使用環境local、検証DB、PC内の添付保存先を確認。
2. 共有添付確認ではVM引数 `-Dspring.profiles.active=dev` または環境変数 `SPRING_PROFILES_ACTIVE=dev` を指定して再起動。使用環境dev・共有保存先のログを確認。
3. 検証サーバーではdev、本番ではprodを明示指定。外部設定の優先順位も含め、表示された接続先・保存先が配備条件と一致することを確認。

```text
[MCM起動設定] 使用環境=local
[MCM起動設定] DB接続先（設定）=サーバー=ILCM-SYSDEV / インスタンス=未指定 / ポート=1433 / DB=SYSDEV
[MCM起動設定] 添付保存先=C:/opt/mcm/upload
```

これは接続設定の表示であり、DBや共有への接続成功を示すログではない。INFOが無効な配備設定では表示されない。ログのために既存のログレベルは変更しない。

## 検証・制約

- 最新srcに61を重ねたmain Java404ファイルの隔離コンパイル成功。
- 実Spring Boot 3.3.5の設定読込・ApplicationRunner・実SQL Server JDBCドライバーのURL解析を使った18項目合格。YAML設定値不変、指定なしlocal、引数／環境変数／VM引数dev、prod環境変数、外部上書き、Hikari URL、URL内の認証情報非表示、解析不能時の表示、ログ値の改行除去を確認。
- 最新src543ファイルと56の全SHA-256不変、既存5／新規0、56重複2を確認。[検証記録](../../検証記録/review_20261003/environment61/README.md)。
- 実DB・共有フォルダ・配備先・Eclipse実機の起動は未確認。検証はDBを起動しない最小Spring構成で実施。添付処理・排他処理・既存フォルダ・DBの保存済みパスは変更していない。
