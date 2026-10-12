# 2026-10-10 配布内容の同期

78（542・543）と順不同で反映できるよう、HTMLを78と同一内容へ更新し、対応するMcm1005uController.javaも同梱しました。現在の76は2ファイルです。515修正を含み、削除ボタン・画面文言とController通知も同期しています。サーバー側の543状態検査は78のServiceにあるため、542・543には78一式を反映してください。
今回追加した削除関連変更については利用者指示に従いテスト未実施です。以下のEdge21判定等は515作成時点の結果で、今回更新後の結果ではありません。旧76のコピーを後から上書きしないでください。

# 76：チケット515 添付追加時の不要な離脱確認

【不具合原因】
ファイルinputのonchangeでuploadAttachment→sendForm→requestSubmitを実行し、dirty=falseにした後、同じchangeイベントがmainFormにバブリングしてdirty=trueに戻ります。画面固有のbeforeunloadがこの値を見て不要な離脱確認を要求します。submitで共通待機表示はすでに開始しているため、離脱をキャンセルすると表示が残ります。ブラウザーの通常の離脱制限ではなく、画面のイベント順序に起因します。

【再現手順】
MCM1005Uの編集可能な登録済み期間で「ファイル追加」→ファイル選択。送信開始後に不要な離脱確認が出ます。キャンセルすると待機表示が残ります。
修正前の実ハンドラーと共通JSを使ったEdge試験で、change後のbeforeunloadが取消状態になることを再現しました。修正後の実送信試験ではダイアログ0回を確認しています（HTTP保存応答は模擬）。

【調査結果】
- 画面：main/resources/templates/mcm1005u/index.html。fileは既にhiddenで、「ファイル追加」からchooseAttachmentで選択、onchangeのuploadAttachmentでmainFormを送信する構成。
- Action：POSTでoperation=uploadを指定し、attachmentPeriodIdとattachmentFile_期間IDを送信。
- Controller：Mcm1005uControllerの登録処理で添付追加を分岐し、Mcm1005uAttachmentService.addを実行。成功時はsubmittedValuesを引き継いで同画面へリダイレクトし、添付一覧を再取得する構成。
- Service：作成権限、契約編集可否、期間所属・編集可否、空ファイル・名前・重複を検査し、ファイル保存とMCM_TK_TENPU登録を実行。失敗時のファイル後処理あり。
- 成功通知：「ファイルの追加を完了しました。」。VBのCPMessage.xmlのMSG_0072Gと一致。
- VB：Mcm1005uTabControlのFileUpButton（CPFileDialogButton）で選択・追加し、MSG_0072を表示。不要なブラウザー離脱確認はない。
- 共通JS：common.jsのsubmitで待機表示と意図した遷移を記録するが、画面独自beforeunloadは別途dirtyを見る。
- 他画面：3005Uのfileにはdata-mcm-ignore-dirtyがある。2003Uもhidden fileから選択・追加する構成。画面固有フローが異なるため一括変更しない。

【修正方針】
ファイル選択inputにdata-mcm-ignore-dirtyを付与。画面固有のinput/changeハンドラーでもfileを未保存判定から除外し、送信後にdirtyが戻ることを防ぎます。一般の契約項目のinput/changeは従来どおりdirtyを立てます。sendForm、アップロード、リダイレクト、通知の流れは維持します。

【影響範囲】
MCM1005UのHTML1ファイルのみです。Controller・Service・共通JSは変更しません。権限・承認状態・未登録期間のdisabled条件、添付削除・ダウンロードのURL／ハンドラーも変更なし。75とは変更ファイルの重複なし。
旧70／72に同名HTMLが含まれるため、76反映後に旧修正版を上書きしないでください。76は最新受領版を元にしています。

【確認項目】
実施済み：Edge21判定、実際のファイル選択→multipart送信→模擬成功画面、MSG_0072G文言・一覧、画面編集済みでも添付追加では離脱確認なし、通常項目の未保存警告維持、file inputで既存dirtyを変更しない、未選択では送信・待機なし、file無効時は追加なし。HTML内JavaScript全文の構文合格。受領src546ファイルのハッシュ不変。Java変更がないため再ビルド試験は実施していません。

反映後に実施：
1. 編集可能な期間でファイル追加→保存→一覧表示・完了通知。離脱確認が出ず、待機表示が終了すること。
2. ファイル選択ダイアログでキャンセルし、送信されず画面が操作できること。同じファイルの再選択ができること。
3. 契約項目を編集した後のファイル追加でも離脱確認が出ず、その編集値が戻り画面に保持されること。
4. 通常項目編集後にタブを閉じる／別URLへ移る場合は未保存警告が残ること。
5. 同名ファイル・保存失敗で適切なエラーが出て、画面が操作できること。
6. 参照権限・承認中／承認済みなど編集不可期間・未登録期間では、従来どおり追加できないこと。
7. 既存添付の削除・ダウンロードが従来どおりできること。

【追加調査が必要な場合】
実DBとサーバー保存先は今回未使用です。添付保存自体が失敗する場合はMcm1005uAttachmentService.add、application-環境別.yamlの保存先、稼働アカウントの書込み権限とサーバーログを確認してください。権限・承認状態ごとの実機ケースも配備環境で確認してください。

## 反映方法
src/main/resources/templates/mcm1005u/index.htmlをプロジェクトの同名ファイルへ上書きし、再ビルド・配備してください。ZIPは作成していません。最新mcm-web/srcへは未反映です。
