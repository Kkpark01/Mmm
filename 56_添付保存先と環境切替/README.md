# 56：添付保存先と環境切替

基準：2026-10-01 14:34受領src（536ファイル）。**受領src・配備先・DBは変更していません。**

## 今回の整理：新規ファイルをなくす

保存ルート、相対パス・旧絶対パスの読み替え、取得時の検査を既存の`FileStorageService.java`へ統合した。6サービスの参照先も変更し、例外は既存の`McmBusinessException`を使う。利用者向けの文言・原因ログ、権限、ロールバック、削除タイミングは維持する。

`AttachmentStorage.java`と`AttachmentStorageException.java`を削除し、**反映対象は既存9ファイル、新規0ファイル**。設定ファイルと起動チェックは前回のコメントアウト方式から変更していない。[今回の検証記録](../../検証記録/review_20261001/attachment-storage-consolidation/README.md)。

## 今回の整理：コメントアウトで切替

利用者の指示により、サーバー名による自動判定とlocal/test/prodのプロファイル選択を廃止した。検証用を有効、本番用をコメントアウトで用意する。

- `application.yml`の検証用DBと`F:/uploadfolder/MCM`を現在の設定として使う。
- 本番は末尾のコメントアウトされたYAML文書にまとめた。`# ---`から最後まで先頭の`# `を外すと、後続文書が検証用DB・保存先・環境区分を上書きする。
- 本番DB・認証・保存先は未定。空欄のまま有効にすると、通常Bean（DB接続を含む）の生成前に起動を止める。保存先が相対パスの場合も止める。
- 起動チェックは既存の`McmApplication.java`へ集約した。静的`BeanFactoryPostProcessor`なので、既存のどちらの起動クラスからもコンポーネント検索で読み込まれる。
- `McmEnvironmentSelector.java`、`McmEnvironmentGuard.java`、`META-INF/spring.factories`は廃止。今回の共通保存処理の統合と合わせ、初版56の新規5ファイルをすべて不要にした。
- 以前の`MCM_ENV`、ホスト名、環境選択用`spring.profiles.active`で切り替える手順は使わない。サーバーごとの外部YAMLに、そのサーバーで使う設定を置く。

## 対象ファイル・反映方法

`src`以下の既存9ファイルを、プロジェクトの同じ相対位置へコピーする。

| 区分 | ファイル |
|---|---|
| 環境設定 | `main/resources/application.yml` |
| 起動検査（既存） | `main/java/com/daifuku/mcm/McmApplication.java` |
| 添付登録・取得 | `service/Mcm1005uAttachmentService.java`、`Mcm2003uAttachmentService.java`、`Mcm3005uAttachmentService.java` |
| 承認画面の取得・捺印 | `service/Mcm3001uService.java`、`Mcm2008uService.java`、`Mcm2008uNatsuinService.java` |
| 共通アップロード・パス解決・業務エラー | `common/FileStorageService.java`（例外は既存`McmBusinessException`を使用） |

Javaのパッケージ基準は`main/java/com/daifuku/mcm/`。YAMLだけの反映では旧パス・相対パスの対応は完了しない。

**旧56をプロジェクトへ反映済みの場合**は、以下の旧追加ファイルも削除してからクリーンビルドする。コピーだけでは旧クラスが残る。

- `main/java/com/daifuku/mcm/config/McmEnvironmentSelector.java`
- `main/java/com/daifuku/mcm/config/McmEnvironmentGuard.java`
- `main/resources/META-INF/spring.factories`（他の登録を加えている場合はファイルごと削除せず、上記2クラスの登録だけを除く）
- `main/java/com/daifuku/mcm/common/AttachmentStorage.java`
- `main/java/com/daifuku/mcm/exception/AttachmentStorageException.java`

Eclipseは「プロジェクト → クリーン」、JAR作成は`mvn clean package`等で旧生成物を除く。その後、配備先の外部YAMLも更新して再起動する。受領srcには旧追加5ファイルは存在しない。

## 配備先と開発PCの設定

### 検証サーバー

Javaが動くサーバー上に`F:/uploadfolder/MCM`があり、Java実行ユーザーで読取・作成・変更・削除できることを確認する。画像のFが別サーバーやRDP利用者だけの割当ドライブの場合は、実際の共有UNCへ変更する。検証DBは従来の`ILCM-SYSDEV:1433 / SYSDEV`を維持。

推奨は、今回のYAMLを配備先の`config/application.yml`へ配置し、検証用を有効にしたまま使うこと。例：`D:/Webアプリ/mcm-web/config/application.yml`。実際の起動設定・外部設定の配置場所は未確認。

```powershell
# 設定例。実際のJAR名・配置へ置き換える。
Set-Location -LiteralPath 'D:/Webアプリ/mcm-web'
java -jar 'jar/実際のJAR名.jar' --spring.config.additional-location=file:./config/
```

外部YAMLの標準検索はJARの位置ではなく起動時の作業ディレクトリに依存するため、上記では設定位置も明示している。同じJARでも、各サーバーの外部YAMLで設定を変えられる。

### 開発PC

検証DBを使うPCは、検証サーバーと**同じ実ファイル**を指す共有UNCを設定する。自PCのFドライブや個人フォルダへ保存すると、他PCから開けない。共有名は未定。

```powershell
# 以下は例。実際に決まったサーバー名・共有名へ置き換える。
$env:MCM_UPLOAD_PATH = '//保存サーバー/共有名/MCM'
$env:MCM_LEGACY_ROOTS = 'F:/uploadfolder/MCM'
java -jar '実際のJAR名.jar'
```

Eclipseの実行構成の「環境」欄でも指定できる。添付以外の開発だけを行う場合、開発PCの外部YAMLで`mcm.file.upload-path: ''`にすれば、添付操作時に未設定の案内を出す。個人用保存先を使う場合はDBも個人用へ変更する。

### 本番サーバー

本番側の外部`application.yml`で、末尾の`# ---`から最後までの先頭`# `を外す。`---`も含め、ブロック全体を有効にする。DBと保存先を別々に切り替えない。

必要な値は`MCM_PROD_DB_URL`、`MCM_PROD_DB_USERNAME`、`MCM_PROD_DB_PASSWORD`、`MCM_PROD_UPLOAD_PATH`。旧パスがある場合は`MCM_PROD_LEGACY_ROOTS`も設定する。環境変数を使わず、YAMLの該当値を直接設定することも可能。秘密情報は共有READMEへ記載しない。

本番名・DB・保存先は未定。起動チェックは未設定を検出するが、接続先が本当に本番か、フォルダの存在・権限まで保証するものではない。本番切替は手動であり、設定変更後に再起動する。

## 添付保存・既存資料の移行

- 新規は日本語業務フォルダを使い、取引先は`取引先/契約ID/期間ID/UUID/ファイル名`、店舗は`店舗/見積ID/UUID`、作業予定は`作業予定/登録ID/UUID`へ保存する。
- 新規DBパスは相対パス。旧フルファイル／フォルダ形式、旧英語フォルダにも対応。取引先添付のファイル名二重連結も修正済み。
- 旧絶対パスは`legacy-roots`へ明示した旧ルートだけを現在の保存先へ読み替える。実ファイルを新ルートへコピーし、旧ルートから下の構成を維持する。DBを自動更新したり、同名ファイルを全ドライブから探したりしない。
- 例：旧`D:/旧保存先/MCM/取引先/顧客/資料.xls`を`F:/uploadfolder/MCM/取引先/顧客/資料.xls`へコピーした場合、旧ルートは`D:/旧保存先/MCM`。DBの実パスを見て設定し、複数はカンマ区切りで指定する。
- 未設定・未存在・読取不可・実ファイルなしは業務エラーとして共通ハンドラーへ渡す。権限、同名禁止、排他、登録失敗／ロールバック時の削除、コミット後の削除、ルート外／リンク経由の逸脱防止は維持した。

**切り戻し／旧VBとの併用には注意**。新規相対パスはWeb版向けであり、旧JARだけに戻すと新規分を開けない。パス互換処理を残すか、旧版が読める絶対パスへ戻す手順が必要。同じ添付テーブルを旧VBが使う場合も相対パスの互換性を確認する。

## 競合・検証

最新srcの`Mcm2008uService.approve()`のint返却・件数・破棄／解約のtrim判定は、[前回統合](../../検証記録/review_20261001/latest-source-conflicts/README.md)を維持。55との対象重複は0。旧55は新srcの性能対策と異なるので一括上書きしない。53のSMTP未設定時の既知差異は受領コードに残る。旧43のOffice起動対応は取得サービスが重複するため別途差分統合が必要。

今回の[共通保存処理の統合検証](../../検証記録/review_20261001/attachment-storage-consolidation/README.md)を参照。前回の[コメント切替検証28項目](../../検証記録/review_20261001/attachment-comment-config/README.md)、初版の[84項目](../../検証記録/review_20261001/attachment-environments/README.md)・[追加影響13項目](../../検証記録/review_20261001/attachment-environments/影響確認.md)は履歴として保持する。

今回の統合後はmain Java401ファイルのコンパイル成功、ローカル111項目合格。受領src536ファイルのハッシュ一致、新規ファイル0、55との対象重複0も確認した。

実DB、サーバーF、UNC、実行ユーザーのACL、他PCからの参照、本番、Office起動は未確認。リンクは既存のダウンロード動作を維持する。
