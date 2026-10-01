# 56：添付保存先と環境切替

2026-10-01受領のsrcを基準に作成。**受領src、配備先、DBは変更していません。**

## 変更内容

- 同じJARで`local`（開発PC）、`test`（検証）、`prod`（本番）を切り替え、DB設定と添付保存先を一緒に読む。
- 優先順位は明示プロファイル（起動引数／環境変数／YAML）→`MCM_ENV`→Javaが動くWindowsの`COMPUTERNAME`→`local`。外部`config/application.yml`、`spring.config.additional-location`も標準Spring設定として扱う。
- 検証ホスト名の初期値は以前共有された`ILCM-DEV`。**実際のJava稼働ホスト名は未確認**。URLの`ilcm-sysdev`、DBホスト名、閲覧PC名では判定しない。本番ホスト名は空欄。
- 検証保存先は画像の`F:/uploadfolder/MCM`を基準とし、1005の新規添付を`取引先/契約ID/期間ID/UUID/ファイル名`へ保存。店舗は`店舗/見積ID/UUID`、作業予定は`作業予定/登録ID/UUID`。
- 新規DBパスは相対パス（取引先／作業予定はファイル、店舗はフォルダ）。旧フルファイル／フォルダ両形式と旧英語フォルダを参照できる。3001の取引先添付にファイル名を二重連結する問題も修正。
- 旧絶対パスは、設定した旧ルートだけを現在の保存先へ読み替える。DBの値を自動変更したり、全ドライブから同名ファイルを探したりしない。
- 保存先未設定・未存在・アクセス不可・実ファイルなしは業務エラーとして既存共通ハンドラーへ渡す。内部パス／例外原因はログに残す。既存2003等の個別捕捉経路は今回変更していない。
- 権限、同名禁止、排他、登録失敗／ロールバック時の専用ファイル削除、コミット後削除を維持。読み替え先でもルート外・`..`・ジャンクション経由の逸脱を拒否する。

## 対象ファイル・反映方法

`src`以下の13ファイルを、プロジェクトの同じ相対位置へコピーする。**追加クラスとMETA-INF/spring.factoriesも必須**。YAMLだけの反映では旧パス・相対パス対応は完了しない。

| 区分 | ファイル |
|---|---|
| 環境設定 | `main/resources/application.yml` |
| 起動登録（追加） | `main/resources/META-INF/spring.factories` |
| 環境選択・本番検査（追加） | `config/McmEnvironmentSelector.java`、`McmEnvironmentGuard.java` |
| 共通保存先／業務例外（追加） | `common/AttachmentStorage.java`、`exception/AttachmentStorageException.java` |
| 添付登録・取得 | `service/Mcm1005uAttachmentService.java`、`Mcm2003uAttachmentService.java`、`Mcm3005uAttachmentService.java` |
| 承認画面の取得・捺印 | `service/Mcm3001uService.java`、`Mcm2008uService.java`、`Mcm2008uNatsuinService.java` |
| 共通アップロード | `common/FileStorageService.java` |

全Javaのパッケージ基準は`main/java/com/daifuku/mcm/`。帳票出力フォルダの設定は今回の対象外。

51～55との対象重複は0件で、反映順の指定はない。**53のSMTP未設定時のVB差異による反映保留は継続**。53の捺印サービスが56の相対パスを扱えることは確認した。古い43（2003取得／Office起動）は56と取得サービスが重複するため一括上書きしない。Office起動が必要な場合は最新版へ別途差分統合する。

アプリをビルドし直し、配備先の外部YAMLが新設定を上書きしていないか確認して再起動する。配備先`config`が標準検索対象になるかは**JARの位置ではなく起動時の作業ディレクトリ**に依存する。必要なら起動引数で外部設定位置を指定する。

## 配備時の設定

### 検証サーバー

1. Javaが動くサーバーで`hostname`を確認。名前が`ILCM-DEV`と異なる場合はYAMLの`mcm.runtime.test-server`か`MCM_TEST_SERVER`を変更する。
2. **そのサーバー上**で`F:/uploadfolder/MCM/取引先`を開けることを確認する。画像のFが別サーバー／RDP利用者だけの割当ドライブなら、そのまま使わずUNCへ変更する。
3. Javaを実行するWindowsユーザーに、保存ルートの読取・作成・変更・削除権限を設定する。RDP利用者が開けるだけでは確認完了にならない。
4. 初回は`--spring.profiles.active=test`を明示して確認。その後、ホスト名の自動判定で同じ設定が選ばれることを確認する。

```powershell
# 実際のJAR名に置き換える。保存ルート自体は管理者が事前に用意する。
$env:MCM_ENV = 'test'
$env:MCM_UPLOAD_PATH = 'F:/uploadfolder/MCM'
java -jar '実際のJAR名.jar'
```

検証DBは従来の`ILCM-SYSDEV:1433 / SYSDEV`を維持する。URL変更は`MCM_TEST_DB_URL`、認証変更は標準の`SPRING_DATASOURCE_USERNAME`／`SPRING_DATASOURCE_PASSWORD`でも指定できる。秘密情報は共有README等へ書かない。

### 開発PC

検証DBを使う場合、全PCで**検証サーバーと同じ実ファイル**を参照する必要がある。共有名が未定なので`local`の保存先は空欄とし、添付以外の開発は続けられるようにした。保存先が空のまま添付操作すると未設定の案内が出る。

```powershell
# 以下のUNCは例。実際に決まったサーバー名・共有名へ置き換える。
$env:MCM_ENV = 'local'
$env:MCM_UPLOAD_PATH = '//保存サーバー/共有名/MCM'
# 検証DBにFドライブの絶対パスが残る場合に、同じ相対位置へ読み替える。
$env:MCM_LEGACY_ROOTS = 'F:/uploadfolder/MCM'
java -jar '実際のJAR名.jar'
```

この設定はEclipseの実行構成の「環境」欄でも指定できる。**検証DBのまま個人PCの別フォルダへ保存すると他PCから開けない**。個人保存先を使う場合はDBも個人用へ切り替える。新しい相対パスはWeb版での運用を前提とし、同じ添付テーブルを旧VB版が直接利用する併用運用は別途確認が必要。

### 本番

本番名・保存先・DBは未定のため値を入れていない。`prod`を選んだとき、DB URL・ユーザー・パスワード・保存先のいずれかが空なら**DB接続／Bean初期化前に停止**する。起動した時点でフォルダの存在・権限まで保証する仕組みではない。

設定する項目は`MCM_PROD_SERVER`（ホスト自動判定用）、`MCM_PROD_DB_URL`、`MCM_PROD_DB_USERNAME`、`MCM_PROD_DB_PASSWORD`、`MCM_UPLOAD_PATH`。初回配備では`--spring.profiles.active=prod`を明示し、検証DBへ接続しないことを確認する。未登録ホストは`local`になるため、ホスト登録前に自動判定だけで本番運用を始めない。

## 既存ファイルの移行

ファイルを保存先へコピーし、旧ルートから下の構成を維持する。例：旧`D:/旧保存先/MCM/取引先/顧客フォルダ/資料.xls`を新`F:/uploadfolder/MCM/取引先/顧客フォルダ/資料.xls`へコピーした場合、`MCM_LEGACY_ROOTS=D:/旧保存先/MCM`を設定する。

旧ルートは**DBの実パスを見てから**決める。画像だけでは確定しない。`取引先`を含む位置の対応を揃える。複数はカンマ区切りで指定できる。`C:/opt/mcm/upload`も、その配下の実ファイルを同じ構成で移した場合だけ登録する。

読み替えによってDBを更新せずに既存リンクを検証できる。DBを相対パスへ統一するSQLの実行は今回行っていない。未コピーの資料、構成が異なるコピー、アクセスできない共有先は設定だけでは直らない。

## 検証結果・残る確認

### 追加の影響確認（2026-10-01）

保存先切替と削除の順序を追加13項目で確認し、すべて合格。[影響確認記録](../../検証記録/review_20261001/attachment-environments/影響確認.md)。コードの追加変更はない。

**切り戻し時には注意が必要**。56で新規添付を登録すると、DBに相対パスが入る。旧JARの取得処理は相対パスを保存ルートに結合しないため、JARだけを旧版へ戻すと新規分を開けなくなる。切り戻す場合は、56のパス互換処理を維持するか、対象の相対パスを旧版が読める絶対パスへ戻す手順を事前に用意する。旧VB版と同じ添付テーブルを併用する場合も、相対パスの互換性を確認する。

詳細は[検証記録](../../検証記録/review_20261001/attachment-environments/README.md)。統合main Java405ファイルのコンパイル成功。ローカル実ファイル／実Spring設定読込／HTTP共通エラー経路／登録・ロールバック／承認参照・捺印／Windowsジャンクションを検証。受領src532ファイルは全SHA-256一致。

実DB、サーバーF、UNC接続、実行ユーザーのACL、実利用者の追加→他PC参照、本番、Officeアプリで開く操作は未確認。**取得処理の修正だけでOfficeを直接開く動作には変わらない**。この版のリンクは既存のダウンロード動作を維持する。
