# VB機能・条件分岐一覧の再確認

入口：[VB機能・条件分岐一覧](VB機能・条件分岐一覧.md)。既存122項目を原文と再照合し、15項目を追加して137項目へ補足。既存機能IDと原文根拠IDを維持しています。

137業務項目と以下の共通10項目の単体・結合への分担は、[結合テスト仕様書のVB機能照合](MCM契・解約機能_結合テスト仕様書.md#sheet-16)へ記録しています。内部の似た説明は責任範囲を整理し、原文引用の文脈や別画面・別タブの確認対象は保持しました。対応先・残条件と、実行済みの判定は別です。

今回は、前回の全件索引にあった機器反映・計算・集約の分岐を業務要約へ対応づけ、共通処理・入力キー・DataSet制約を追加照合しました。元VBの動作とコメントが異なる箇所を区別し、Web期待仕様の採否は別判定です。

## 主な訂正・補足

| 対象 | 前回の不足・解釈 | 再確認した内容 |
| --- | --- | --- |
| 1006Uほか | 通知後の中止が未確定 | 画面のDisplayMessageはE末尾なら例外で中止。MSG_0023E・0037E・0038Eも該当。CPMessageUtility直接呼出しとは異なる |
| 1005U | 月割り・選定差分が初期復元等の大きな項目に含まれていた | 構成・明細・個体・単価の追加／削除／維持更新、点検重複、数量・金額・月割り・余りを個別項目化 |
| 1005U | 月チェックの不一致が抽象的 | 指定回数<ON月数だけエラー。同数・ON月数不足を同じ禁止条件としない |
| 2006U | ブランド初期化・依頼店舗・保守集約の細条件が不足 | 未選定ブランドのマスタ採用、店舗取得0／1／複数、方法不一致時の変数違い、個体ON→ON分岐なしを明記 |
| 3005U | 予定日空欄を登録可能と解釈 | NewRowへのDBNull代入は現在行の必須解除ではない。通常保存はRequired=Trueと自動検査を通る |
| 3005U | 新規起動の違いが不足 | 受渡しありID0はE通知で中止。受渡しなしのID0はSearchで新規作成 |
| 1011U | 画面内の押印とFalseの分岐中心 | .xlsx以外は捺印せずTrue。SaveAs失敗Falseと伝播例外を分け、DBとExcelの更新範囲も区別 |
| 1011U通知先 | 親階層と金額の比較列が抽象的 | 金額WHEREは上位候補MOAではなくログイン本人MOB。上限>金額の厳密比較と候補0件を共通C03に明記 |
| 共通 | 追加確認の項目としてのみ記載 | 参照権限、E通知、未保存確認、入力検査、DBコミット／ロールバック、検索条件省略を原文で補足 |

## 共通の機能・条件

画面ごとの有効・無効・参照経路は原文一覧のView等と併せて確認します。ここだけで全画面の最終表示を確定しないでください。

<a id="C06"></a>

### C06　エラー通知と確認取消

条件：画面DisplayMessageの末尾E／それ以外、CPMessageUtility直接呼出し、登録確認Cancel。

画面のE通知は表示後にCPInputCheckExceptionをthrow。直接ユーティリティ呼出しは通常DialogResultを返す。DisplaySaveConfirm取消はCPProcessCanceledException。

<details>
<summary>原文根拠</summary>

[CPCoreUserControl.vb:936](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPCoreUserControl.vb>)～985

```vb
    ''' <summary>
    ''' メッセージを表示する。
    ''' </summary>
    ''' <param name="messageId">メッセージID</param>
    ''' <param name="replaceArg">置換え文字列</param>
    ''' <returns>ダイアログ表示後に押下したボタンの種類</returns>
    ''' <remarks></remarks>
    Protected Function DisplayMessage(ByVal messageId As String, ByVal replaceArg As String) As DialogResult
        Return Me.DisplayMessage(messageId, New String() {replaceArg})
    End Function

    ''' <summary>
    ''' メッセージを表示する。
    ''' </summary>
    ''' <param name="messageId">メッセージID</param>
    ''' <param name="replaceArgs">置換え文字列 ２つ以上</param>
    ''' <returns>ダイアログ表示後に押下したボタンの種類</returns>
    ''' <remarks></remarks>
    Protected Function DisplayMessage(ByVal messageId As String, ByVal replaceArgs As String()) As DialogResult

        Dim result As DialogResult = CPMessageUtility.DisplayMessage(messageId, replaceArgs)

        If messageId.Substring(messageId.Length - 1) = CPMessageConstant.MESSAGE_SUFFIX_ERROR Then
            Throw New CPInputCheckException
        End If

        Return result
    End Function

    ''' <summary>
    ''' メッセージを表示する。
    ''' </summary>
    ''' <param name="messageId">メッセージID</param>
    ''' <returns>ダイアログ表示後に押下したボタンの種類</returns>
    ''' <remarks></remarks>
    Protected Function DisplayMessage(ByVal messageId As String) As DialogResult
        Return Me.DisplayMessage(messageId, New String() {})
    End Function

    ''' <summary>
    ''' 登録時の確認メッセージを表示する。キャンセル時は処理を抜ける。
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub DisplaySaveConfirm()

        If CPMessageUtility.DisplayMessage(CPMessageConstant.MSG_0007) = DialogResult.Cancel Then
            Throw New CPProcessCanceledException
        End If

    End Sub
```

[CPMessageUtility.vb:22](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPMessageUtility.vb>)～62

```vb
    Public Shared Function DisplayMessage(ByRef messageId As String, _
                                                 ByRef replaceString As String()) As DialogResult

        '' メッセージを取得する
        Dim message As String = Nothing
        Dim messageData As CPMessageData = CPSettingInfo.MessageResource.Item(messageId)

        '' CPMessage.xmlになければmessageIdを表示
        If messageData Is Nothing Then
            messageData = New CPMessageData
            messageData.Message = messageId
            messageData.Buttons = MessageBoxButtons.OK
            messageData.Icon = MessageBoxIcon.Warning
        End If

        '' メッセージがなければ明示的にエラーをthrowする
        If CPCommonUtility.IsNull(messageData.Message) Then
            Throw New CPSpecifiedException("メッセージがありません。")
        Else
            '' 置換え文字をメッセージに埋め込む
            message = ConvertReplaceString(messageData.Message, replaceString)
            message = message.Replace("vbCrLf", vbCrLf)
        End If

        '' メッセージダイアログを表示する
        Dim result As DialogResult = MessageBox.Show( _
            message, messageData.Caption, messageData.Buttons, messageData.Icon)

        Return result
    End Function

    ''' <summary>
    ''' メッセージを表示する。
    ''' </summary>
    ''' <param name="messageId">メッセージID</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DisplayMessage(ByRef messageId As String) As DialogResult

        Return DisplayMessage(messageId, Nothing)
    End Function
```

[CPMessageConstant.vb:10](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONSTANT/CPMessageConstant.vb>)～16

```vb
''' <remarks></remarks>
Public Class CPMessageConstant

    Public Const MESSAGE_SUFFIX_ERROR = "E"

    Public Const MESSAGE_SUFFIX_GUIDE = "G"

```

[CPMessageConstant.vb:140](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONSTANT/CPMessageConstant.vb>)～149

```vb
    Public Const MSG_0021 As String = "MSG_0021E"

    ''' <summary>{0}既に登録されている為、削除する事が出来ません。</summary>
    Public Const MSG_0022 As String = "MSG_0022E"

    ''' <summary>見積依頼する機器を選定してください。</summary>
    Public Const MSG_0023 As String = "MSG_0023E"

    ''' <summary>保守契約時間帯を選択してください。</summary>
    Public Const MSG_0024 As String = "MSG_0024E"
```

[CPMessageConstant.vb:185](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONSTANT/CPMessageConstant.vb>)～194

```vb
    Public Const MSG_0036 As String = "MSG_0036E"

    ''' <summary>１件もチェックされていません。</summary>
    Public Const MSG_0037 As String = "MSG_0037E"

    ''' <summary>選択した見積内に重複した機器が存在します。</summary>
    Public Const MSG_0038 As String = "MSG_0038E"

    ''' <summary>上位者IDに対応する上位者のログインIDが存在しません。</summary>
    Public Const MSG_0039 As String = "MSG_0039E"
```

</details>

<a id="C07"></a>

### C07　起動と参照権限

条件：権限制御ON／OFF、画面権限未登録／1／2、AuthorityIsThrough、画面IDの枝番。

権限取得は画面ID先頭8文字。権限情報なしは起動を拒否。Loadイベント後、参照1では通過設定なしのCPコントロールを制限。グリッドは読取専用・新規行禁止、テキストは読取専用、コンボ・チェック・ボタン等は非活性。検索など通過設定は除外。値0が文字列として存在する場合の共通起動判定を「必ず拒否」とは置き換えない。

<details>
<summary>原文根拠</summary>

[CPBaseForm.vb:414](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPBaseForm.vb>)～438

```vb
    ''' <summary>
    ''' 権限チェックを行い、権限がなければ画面は表示しない。
    ''' </summary>
    ''' <param name="targetScreen"></param>
    ''' <remarks></remarks>
    Private Function ControllAuthorityDisplayScreen(ByVal targetScreen As String) As Boolean

        '' 開発用
        If Not CPSettingInfo.IsAuthorityControl Then
            Return True
        End If

        '' 権限区分を取得する
        Dim authorityDivision As String = CPSettingInfo.GetUserInfo.GetAuthorityDivision(targetScreen)

        '' 権限区分がない場合は、権限がないと判断されるためエラーとする
        If CPCommonUtility.IsNull(authorityDivision) Then
            CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0012)
            Return False
            'Throw New CPAuthorityException()
        End If

        Return True

    End Function
```

[CPCoreUserControl.vb:154](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPCoreUserControl.vb>)～297

```vb
    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)

        Try
            '' Loadのイベントを起動する
            MyBase.OnLoad(e)

            '' Daoがある場合はコミット処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Commit()
            End If

            '' 権限制御
            Me.AuthorityControllComponent()

        Catch ex As Exception
            '' Daoがある場合はロールバック処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Rollback()
            End If

            '' Exceptionごとの処理
            CPComponentUtility.HandleException(ex)

        Finally
            '' Daoがある場合はクローズ処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Close()
                dao = Nothing
                CPDaoContainer.SetDao(Nothing)
            End If

            '' システム日付をクリア
            CPSettingInfo.ClearSystemDate()

        End Try

    End Sub

    ''' <summary>
    ''' 参照権限のみの場合は、コンポーネントを非活性にする。
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub AuthorityControllComponent()

        If Not CPSettingInfo.IsAuthorityControl Then
            Return
        End If

        '' 権限区分を取得し、権限制御を行う
        Dim authorityDivision As String = Me.GetAuthorityDivision()

        If CPConstant.AUTHORITY_DIVISION_INSPECTION.Equals(authorityDivision) Then
            '' 参照のみ権限の場合のみ権限制御を行う
        Else
            Return
        End If

        '' ユーザコントロールに宣言されたフィールドをリフレクションによりすべて取得する
        Dim members As FieldInfo() = Me.GetType.GetFields(BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly)

        '' ループさせ、フィールドの型に合わせた処理を行う
        For Each fieldObject As FieldInfo In members

            '' フィールド名を取得
            Dim name As String = fieldObject.Name

            '' 先頭1桁目が"_"になっているため除去
            If name.Length > 0 Then
                name = name.Substring(1, name.Length - 1)
            End If

            '' リフレクションで実際のフィールドを取得
            Dim obj As Object = CPObjectUtility.GetProperty(Me, name)

            '' フィールドの型によって処理を振り分ける
            If TypeOf obj Is CPDataGridView Then
                '' DataGridViewの場合
                Dim grid As CPDataGridView = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not grid.AuthorityIsThrough Then
                    grid.ReadOnly = True
                    grid.AllowUserToAddRows = False
                End If

            ElseIf TypeOf obj Is CPTextBox Then
                '' TextBoxの場合
                Dim text As CPTextBox = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not text.AuthorityIsThrough Then
                    text.ReadOnly = True
                End If

            ElseIf TypeOf obj Is CPCombobox Then
                '' ComboBoxの場合
                Dim comboBox As CPCombobox = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not comboBox.AuthorityIsThrough Then
                    comboBox.Enabled = False
                End If

            ElseIf TypeOf obj Is CPCheckBox Then
                '' CheckBoxの場合
                Dim checkBox As CPCheckBox = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not checkBox.AuthorityIsThrough Then
                    checkBox.Enabled = False
                End If
            ElseIf TypeOf obj Is CPButton Then
                '' Buttonの場合
                Dim button As CPButton = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not button.AuthorityIsThrough Then
                    button.Enabled = False
                End If
            ElseIf TypeOf obj Is CPButtonDataGridViewRemoveRows Then
                '' 行削除用Buttonの場合
                Dim button As CPButtonDataGridViewRemoveRows = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not button.AuthorityIsThrough Then
                    button.Enabled = False
                End If

            ElseIf TypeOf obj Is CPFileDialogButton Then
                '' 行削除用Buttonの場合
                Dim button As CPFileDialogButton = obj

                '' 権限制御無視フラグがOffの場合のみ処理を行う
                If Not button.AuthorityIsThrough Then
                    button.Enabled = False
                End If

            End If

        Next
    End Sub
```

[CPUserInfo.vb:281](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/INFO/CPUserInfo.vb>)～311

```vb
    Public Function GetAuthorityDivision(ByVal functionId As String) As String

        If functionId.Length > CPConstant.FUNCTION_ID_LENGTH Then
            functionId = functionId.Substring(0, CPConstant.FUNCTION_ID_LENGTH)
        End If

        Dim thisId As String = functionId.ToUpper

        '' 権限対象外のIDの場合は処理を抜ける
        For Each id As String In CPSettingInfo.AuthorityIsThroughId
            If id.ToUpper.Equals(thisId) Then
                Return CPConstant.AUTHORITY_DIVISION_UPDATE
            End If
        Next

        Dim authorityList As List(Of CPAuthorityData) = Me.FunctionAuthority
        Dim authorityDivision As String = Nothing

        For Each authority In authorityList

            Dim id As String = authority.FunctionId.ToUpper
            Dim value As String = authority.FunctionAuthority

            If id.Equals(thisId) Then
                authorityDivision = value
                Exit For
            End If
        Next

        Return authorityDivision
    End Function
```

</details>

<a id="C08"></a>

### C08　入力検査の条件と復元

条件：必須、未入力、ValidateKey有無、監査項目、型・文字バイト数・数値桁・大小・日付、ReadOnly／Immediate／非活性／非表示。

キーなしは検査を通過、定義なしは例外。必須の空欄はエラー、任意の空欄は後続検査を省略。文字列長はShift-JISバイト、数値長は文字数。下限未満・上限超過を拒否し等値は許容。全画面自動検査では対象外属性を省略し、エラーでは変更前値・フォーカスを復元して中止。CPTextBoxのLeaveはRequired=Falseで検査するため、登録時の必須と区別する。

<details>
<summary>原文根拠</summary>

[CPValidateUtilty.vb:24](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPValidateUtilty.vb>)～175

```vb
    Public Shared Function IsErrorCell(ByVal isRequiredFlag As Boolean, ByVal validateKey As String, _
                         ByVal value As Object, ByVal tagColumnName As String) As Boolean

        '' 入力チェックキーがない場合
        If CPCommonUtility.IsNull(validateKey) Then
            '************** <例外エラー> 2009/09/02 T.Iwasawa Update Start
            Return False
            'Throw New CPSpecifiedException("データバインドがありません。Immidate=trueに設定するか、ValidateKeyを設定してください")
            '************** <例外エラー> 2009/09/02 T.Iwasawa Update End
        End If

        '' 作成日・作成者などは回避
        If validateKey.Equals(CPConstant.TABLE_COLUMN_CREATED_DT) _
            Or validateKey.Equals(CPConstant.TABLE_COLUMN_CREATED_BY) _
            Or validateKey.Equals(CPConstant.TABLE_COLUMN_LASTUPDATE_DT) _
            Or validateKey.Equals(CPConstant.TABLE_COLUMN_LASTUPDATE_BY) Then
            Return False
        End If

        '' 入力チェック情報を取得
        Dim validate As CPValidateData = CPSettingInfo.ValidateProperty.Item(validateKey)

        '' 入力チェック情報がない場合はエラー
        If validate Is Nothing Then
            Throw New CPSpecifiedException("CPValidate.xmlに設定がありません:" & validateKey)
        End If

        '' 項目名称格納用
        Dim headerText As String = Nothing

        '' 引数の項目名称がない場合は、入力チェック情報の項目名称をセット
        If CPCommonUtility.IsNull(tagColumnName) Then
            headerText = validate.Name
        Else
            headerText = tagColumnName
        End If

        '' 必須チェック
        If isRequiredFlag Then
            If CPCommonUtility.IsNull(value) Then
                CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0001, New String() {headerText})
                Return True
            End If
        Else
            '' 必須チェックがない場合は、後続の処理を行わない
            If CPCommonUtility.IsNull(value) Then
                Return False
            End If
        End If

        '' 値の型を判断
        If validate.Type.Equals(CPConstant.VALIDATE_TYPE_STRING) Then
            '' Stringの場合
            If CPCommonUtility.IsNotNull(validate.Length) Then
                '' 桁数チェック
                If Not CPCommonUtility.IsCheckMaxByte(validate.Length, value.ToString) Then
                    CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0002, _
                                                    New String() {headerText, validate.Length.ToString})
                    Return True
                End If
            End If

            '' Maskから文字種のチェックを行う
            If CPCommonUtility.IsNotNull(validate.Mask) Then
                If validate.Mask.Equals(CPConstant.VALIDATE_MASK_TYPE_HANKAKU) Then
                    '' 半角チェック
                    If Not CPCommonUtility.IsCheckHankaku(value.ToString) Then
                        CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0008, New String() {headerText})
                        Return True
                    End If
                End If
            End If

        ElseIf validate.Type.Equals(CPConstant.VALIDATE_TYPE_DECIMAL) Then
            '' 数値の場合
            If Not IsNumeric(value) Then
                '' 数値かどうかのチェック
                CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0003, New String() {headerText})
                Return True
            End If

            If CPCommonUtility.IsNotNull(validate.Length) Then
                '' 桁数チェック
                If value.ToString.Length > validate.Length Then
                    CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0002, _
                                                    New String() {headerText, validate.Length.ToString})
                    Return True
                End If
            End If

            If CPCommonUtility.IsNotNull(validate.Mask) Then
                '' 時間チェック
                If validate.Mask.Equals(CPConstant.VALIDATE_MASK_TYPE_JIKAN) Then
                    If Not CPCommonUtility.IsCheckHour(CType(value, Decimal)) Then
                        CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0014, New String() {headerText})
                        Return True
                    End If

                    If CPCommonUtility.IsNull(validate.Max) Then
                        '' デフォルト値をセット 23:59まで
                        validate.Max = 2359
                    End If
                End If

                '' 時間分チェック
                If validate.Mask.Equals(CPConstant.VALIDATE_MASK_TYPE_JIKAN_HUN) Then
                    If Not CPCommonUtility.IsCheckHourMinute(CType(value, Decimal)) Then
                        CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0015, New String() {headerText})
                        Return True
                    End If

                    If CPCommonUtility.IsNull(validate.Max) Then
                        '' デフォルト値をセット 23:59まで
                        validate.Max = 2359
                    End If
                End If
            End If

            If CPCommonUtility.IsNotNull(validate.Min) Then
                '' 下限チェック
                If CType(value, Decimal).CompareTo(CType(validate.Min, Decimal)) < 0 Then
                    CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0004, _
                                                    New String() {headerText, validate.Min.ToString})
                    Return True
                End If
            End If

            If CPCommonUtility.IsNotNull(validate.Max) Then
                '' 上限チェック
                If CType(value, Decimal).CompareTo(CType(validate.Max, Decimal)) > 0 Then
                    CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0005, _
                                                    New String() {headerText, validate.Max.ToString})
                    Return True
                End If
            End If

        ElseIf validate.Type.Equals(CPConstant.VALIDATE_TYPE_DATETIME) Then
            '' 日付の場合
            If TypeOf value Is Date Or TypeOf value Is DateTime Then
                '' 値の型がDate/DateTimeの場合はチェックは不要
            Else
                '' 日付妥当性チェック
                If Not CPCommonUtility.IsCheckDate(value.ToString, validate.Mask) Then
                    CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0006, _
                                                    New String() {headerText, validate.Mask.ToString.ToUpper})
                    Return True
                End If
            End If
        End If

        Return False
    End Function
```

[CPValidateUtilty.vb:182](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPValidateUtilty.vb>)～318

```vb
    Public Shared Sub ValidateAuto(ByRef thisUserControl As CPCoreUserControl)

        '' 画面上のすべてのコンポーネントを取得　※クラス上のすべてのメンバ変数を取得
        Dim members As FieldInfo() = thisUserControl.GetType.GetFields( _
            BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly)

        '' 取得したメンバ変数をループ
        For Each fieldObject As FieldInfo In members

            Dim name As String = fieldObject.Name

            '' 取得した変数名に"_"がデフォルトで付いているため、"_"を除去
            If name.Length > 0 Then
                name = name.Substring(1, name.Length - 1)
            End If

            '' 変数名から実際のオブジェクトを取得
            Dim obj As Object = CPObjectUtility.GetProperty(thisUserControl, name)

            '' オブジェクトの型を判断
            If TypeOf obj Is CPDataGridView Then
                '' DataGridViewの場合

                Dim grid As CPDataGridView = obj

                '' 表示のみ、非表示の場合は処理を抜ける
                If grid.ReadOnly Or Not grid.Enabled Or Not grid.Visible Then
                    Continue For
                End If

                '' 入力チェックを実行
                If grid.Validate() Then
                    '' エラーの場合はExceptionをThrowし処理を抜ける
                    Throw New CPInputCheckException()
                End If

            ElseIf TypeOf obj Is CPTextBox Then
                '' TextBoxの場合

                Dim text As CPTextBox = obj

                '' 表示のみ、入力チェック対象外、非表示の場合は処理を抜ける
                If text.ReadOnly Or text.Immediate Or Not text.Enabled Or Not text.Visible Then
                    Continue For
                End If

                '' 列名
                Dim columnName As String = Nothing
                '' 行データ
                Dim rowData As DataRowView = Nothing
                '' 検索時の値
                Dim orgValue As Object = Nothing

                '' データバインド項目かどうかを判断
                If text.DataBindings.Count > 0 Then
                    '' データバインド項目の場合、列名などを取得
                    columnName = text.DataBindings.Item(0).BindingMemberInfo.BindingField
                    rowData = text.DataBindings.Item(0).BindingManagerBase.Current
                    orgValue = GetOrginalValue(rowData.Row, columnName)
                End If

                '' 入力チェック項目キーを取得
                Dim validateKey As String = text.ValidateKey
                '' タグ名を取得
                Dim headerText As String = text.Tag

                '' 入力チェック項目キーない場合は、列名を入力チェック項目キーにセット
                If CPCommonUtility.IsNull(validateKey) Then
                    validateKey = columnName
                End If

                '' 入力チェックを実行
                If CPValidateUtilty.IsErrorCell(text.Required, validateKey, text.Text, headerText) Then
                    '' 変更前の値がない場合はブランクをセット
                    If CPCommonUtility.IsNull(orgValue) Then
                        orgValue = ""
                    End If
                    '' 変更前の値を表示する値としてセット
                    text.Text = orgValue
                    '' フォーカスを移動
                    text.Focus()
                    '' 処理を抜けるためExceptionをThrow
                    Throw New CPInputCheckException(columnName)
                End If
            ElseIf TypeOf obj Is CPCombobox Then
                '' ComboBoxの場合

                Dim comboBox As CPCombobox = obj

                '' 入力チェック対象外、非表示の場合は処理を抜ける
                If comboBox.Immediate Or Not comboBox.Enabled Or Not comboBox.Visible Then
                    Continue For
                End If

                '' 列名
                Dim columnName As String = Nothing
                '' 行データ
                Dim rowData As DataRowView = Nothing
                '' 検索時の値
                Dim orgValue As Object = Nothing

                '' データバインド項目かどうかを判断
                If comboBox.DataBindings.Count > 0 Then
                    '' データバインド項目の場合、列名などを取得
                    columnName = comboBox.DataBindings.Item(0).BindingMemberInfo.BindingField
                    rowData = comboBox.DataBindings.Item(0).BindingManagerBase.Current
                    orgValue = GetOrginalValue(rowData.Row, columnName)
                End If

                '' 入力チェック項目キーを取得
                Dim validateKey As String = comboBox.ValidateKey
                '' タグ名を取得
                Dim headerText As String = comboBox.Tag

                '' 入力チェック項目キーない場合は、列名を入力チェック項目キーにセット
                If CPCommonUtility.IsNull(validateKey) Then
                    validateKey = columnName
                End If

                '' 入力チェックを実行
                If CPValidateUtilty.IsErrorCell(comboBox.Required, validateKey, comboBox.SelectedValue, headerText) Then
                    '' 変更前の値がない場合はブランクをセット
                    If CPCommonUtility.IsNull(orgValue) Then
                        orgValue = ""
                    End If
                    '' 変更前の値を表示する値としてセット
                    comboBox.SelectedValue = orgValue
                    '' フォーカスを移動
                    comboBox.Focus()
                    '' 処理を抜けるためExceptionをThrow
                    Throw New CPInputCheckException(columnName)
                End If
            End If

        Next
    End Sub

```

[CPDataGridView.vb:440](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPDataGridView.vb>)～525

```vb
    Private Function Validate(ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Boolean

        '' 必須フラグ
        Dim requiredFlag As Boolean
        '' セル
        Dim cell As Object = Me.Columns(columnIndex)
        '' 入力チェックキー
        Dim validateKey As String = Nothing
        '' ヘッダーの名称
        Dim headerText As String = Nothing

        '' セルの型によって処理を振り分ける
        If TypeOf cell Is CPDataGridTextBoxColumn Then
            '' テキストボックスの場合
            Dim column As CPDataGridTextBoxColumn = cell

            '' 入力チェック無視フラグがONの場合、表示のみの場合は処理を抜ける
            If column.Immediate Or column.ReadOnly Then
                Return False
            End If

            '' 必須フラグ取得
            requiredFlag = column.Required
            '' 入力チェックキー取得
            validateKey = column.ValidateKey
            '' ヘッダーの名称を取得
            headerText = column.HeaderText

        ElseIf TypeOf cell Is CPDataGridComboBoxColumn Then
            '' コンボボックスの場合
            Dim column As CPDataGridComboBoxColumn = cell

            '' 入力チェック無視フラグがONの場合、表示のみの場合は処理を抜ける
            If column.Immediate Or column.ReadOnly Then
                Return False
            End If

            '' 必須フラグ取得
            requiredFlag = column.Required
            '' 入力チェックキー取得
            validateKey = column.ValidateKey
            '' ヘッダーの名称を取得
            headerText = column.HeaderText
        Else
            '' チェックボックス、リンクの場合
            Return False
        End If

        '' データバインドされた項目名を取得
        Dim dbColumnName As String = Me.Columns(columnIndex).DataPropertyName

        '' 通常はデータバインドされた項目名をキーにして入力チェックを行うが、
        '' 入力チェックキーを手動で設定した場合はそちらを優先する
        If CPCommonUtility.IsNull(validateKey) Then
            If CPCommonUtility.IsNull(dbColumnName) Then
            Else
                validateKey = dbColumnName
            End If
        End If

        '' 入力した値を取得
        Dim value As Object = Me(columnIndex, rowIndex).Value
        '' DataRowを取得
        Dim dataRow As DataRow = GetDataRow(rowIndex)
        '' 変更前の値を取得
        'Dim orgValue As Object = GetOrginalValue(dataRow, rowIndex, columnIndex)

        '' 入力チェックを実行する
        If CPValidateUtilty.IsErrorCell(requiredFlag, validateKey, value, headerText) Then
            '' 一度エラーになった場合は、入力チェックを起動させないためにフラグを落とす
            Me.IsValidateFlag = False

            '' エラーのセルにフォーカスを移動

            Try
                Me(columnIndex, rowIndex).Selected = True
            Catch ex As InvalidOperationException

            End Try

            Return True
        End If

        Return False

    End Function
```

[CPTextBox.vb:225](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPTextBox.vb>)～277

```vb
    Private Sub CPTextBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Leave

        IsSelectText = False

        '' 自動入力チェックの対象項目の場合
        If CPCommonUtility.CompareTo(Me.BackupValue, Me.Text) Then

            If Not Immediate Then
                Dim columnName As String = Nothing
                If Me.DataBindings.Count > 0 Then
                    columnName = Me.DataBindings.Item(0).BindingMemberInfo.BindingField
                End If

                Dim validateKey As String = Me.ValidateKey

                Dim headerText As String = Nothing

                If Me.RelationLabel IsNot Nothing Then
                    headerText = Me.RelationLabel.Text
                End If

                If Me.Tag IsNot Nothing Then
                    headerText = Me.Tag
                End If

                If CPCommonUtility.IsNull(validateKey) Then
                    validateKey = columnName
                End If

                If CPValidateUtilty.IsErrorCell(False, validateKey, Me.Text, headerText) Then
                    Me.Text = Me.BackupValue
                    Return
                End If
            End If

            '' 変更ステータスをセットする
            If Not ChangedCheckIsThrough Then
                Dim thisForm As CPBaseForm = Me.FindForm
                If thisForm IsNot Nothing Then
                    thisForm.SetChangeStatus()
                End If
            End If
        End If

    End Sub

    Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
        Try
            MyBase.OnLeave(e)
        Catch ex As CPInputCheckException

        End Try
    End Sub
```

[CPCommonUtility.vb:81](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPCommonUtility.vb>)～99

```vb
    Public Shared Function IsCheckMaxByte(ByVal maxByte As Integer, _
                                          ByVal text As String) As Boolean

        Dim byteCount As Integer = 0        '' textのバイト数格納
        Dim judgment As Boolean = True      '' 判定格納変数

        '' テキストのバイト数を取得し、byteCountに格納する
        byteCount = System.Text.Encoding.GetEncoding("shift-jis").GetByteCount(text)

        If byteCount <= maxByte Then
            '' バイト数が最大バイト数より小さいとき、"True"をjudgmentへ格納する
            judgment = True
        Else
            '' バイト数が最大バイト数より大きいとき、"False"をjudgmentへ格納する
            judgment = False
        End If

        '' 判定結果を返す
        Return judgment
```

</details>

<a id="C09"></a>

### C09　空値と0の区別

条件：Nothing・DBNull・空文字・空白のみ／0／有効値。

IsNullはNothing・DBNull・Trim後0文字を空と判定。0は空ではない。IsCheckInputはString.IsNullOrEmptyで判定し、空白の扱いが異なる。型変換やNULL補完は呼出先ごとに読む。

<details>
<summary>原文根拠</summary>

[CPCommonUtility.vb:20](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPCommonUtility.vb>)～71

```vb
    ''' Nullの場合は、Trueを返す。※DBNull, Nothing, ""に対応。
    ''' </summary>
    ''' <param name="target">チェックを行う対象のオブジェクト</param>
    ''' <returns>True:Nullの場合</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNull(ByVal target As Object) As Boolean

        Dim judgment As Boolean = False

        If IsDBNull(target) Or IsNothing(target) Then
            judgment = True
        Else
            If target.ToString.Trim.Length = 0 Then
                judgment = True
            End If
        End If

        Return judgment
    End Function

    ''' <summary>
    ''' Nullでなければ、Trueを返す。※DBNull, Nothing, ""に対応。
    ''' </summary>
    ''' <param name="target">チェックを行う対象のオブジェクト</param>
    ''' <returns>False:Nullの場合</returns>
    ''' <remarks></remarks>
    Public Shared Function IsNotNull(ByVal target As Object) As Boolean
        Return (Not IsNull(target))
    End Function

    ''' <summary>
    ''' 必須チェック関数
    ''' </summary>
    ''' <param name="text"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsCheckInput(ByVal text As String) As Boolean

        Dim judgment As Boolean = True      '' 判定格納変数

        If String.IsNullOrEmpty(text) = True Then
            ' null、もしくは空文字列のとき、"False"をjudgmentへ格納する
            judgment = False
        Else
            ' nullではなく、かつ空文字列でないとき、"True"をjudgmentへ格納する
            judgment = True
        End If

        '' 判定結果を返す
        Return judgment

    End Function
```

</details>

<a id="C10"></a>

### C10　変更なし・編集確定・保存確認

条件：画面変更フラグ、DataSet.HasChanges、BindingSource.EndEdit前後、確認Cancel。

CheckExistUpdatedのオーバーロードによって検査・EndEdit・HasChanges判定順が違う。対象なしでは変更なし例外、登録取消では中止。画面のcheckChangedStatus等での独自判定も併せて読む。

<details>
<summary>原文根拠</summary>

[CPCoreUserControl.vb:987](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPCoreUserControl.vb>)～1048

```vb
    Public Function IsExistUpdated(ByRef targetDataSet As DataSet) As Boolean
        If targetDataSet.HasChanges Then
            '' 自動入力チェックを行う
            CPValidateUtilty.ValidateAuto(Me)
            Return True
        End If
        Return False
    End Function

    Public Function IsChangedStatus() As Boolean
        '' Formを取得
        Dim form As CPBaseForm = Me.FindForm
        Return form.IsChangeStatus
    End Function

    ''' <summary>
    ''' 更新対象のデータがあるかどうかをチェックする。
    ''' </summary>
    ''' <param name="targetDataSet">対象のデータセット</param>
    ''' <remarks></remarks>
    Protected Sub CheckExistUpdated(ByRef targetDataSet As DataSet)
        '' Formを取得
        Dim form As CPBaseForm = Me.FindForm

        '' データセットに変更があること、かつ変更フラグがONであること
        If form.IsChangeStatus And targetDataSet.HasChanges Then
            '' 自動入力チェックを行う
            CPValidateUtilty.ValidateAuto(Me)
        Else
            '' 処理を抜けるためExceptionを投げる
            Throw New CPDataNotChangedException
        End If

    End Sub

    ''' <summary>
    ''' 更新対象のデータがあるかどうかをチェックする。
    ''' </summary>
    ''' <param name="targetDataSet">対象のデータセット</param>
    ''' <remarks></remarks>
    Protected Sub CheckExistUpdated(ByRef targetDataSet As DataSet, ByRef bindSource As Object)
        '' Formを取得
        Dim form As CPBaseForm = Me.FindForm

        '' データセットに変更があること、かつ変更フラグがONであること
        If form.IsChangeStatus Then
            '' 自動入力チェックを行う
            CPValidateUtilty.ValidateAuto(Me)

            '' リフレクションによりBindSourceのEndEditを行う
            CPObjectUtility.InvokeMethod("EndEdit", bindSource, Nothing)

            If targetDataSet.HasChanges Then
            Else
                '' 処理を抜けるためExceptionを投げる
                Throw New CPDataNotChangedException
            End If
        Else
            '' 処理を抜けるためExceptionを投げる
            Throw New CPDataNotChangedException
        End If

```

</details>

<a id="C11"></a>

### C11　戻る・閉じる・子画面からの復帰

条件：変更あり／なし、変更破棄確認Cancel／続行、履歴あり／なし、Ownerと返却値。

変更確認取消なら画面を閉じず／戻らず保持。履歴なしの戻るは何もしない。通常の復帰は再描画で初期Loadを省略。子画面から返すイベントにもDAOコミット／例外時ロールバック／終了処理がある。

<details>
<summary>原文根拠</summary>

[CPBaseForm.vb:562](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPBaseForm.vb>)～683

```vb
    Private Sub CPBaseForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) _
            Handles Me.FormClosing

        Try
            '' 変更ステータスをチェックする
            Me.CheckChangeStatus()
        Catch ex As CPProcessCanceledException
            '' キャンセルの場合、処理を抜ける
            e.Cancel = True
            Return
        End Try

        '' 日付情報をクリア
        CPSettingInfo.ClearSystemDate()
        '' Daoをクリア
        CPDaoContainer.SetDao(Nothing)
        '' 履歴情報をクリア
        Me.HistoryScreenId = Nothing
        '' 引継ぎ情報をクリア
        Me.Delivery = Nothing

        '' 呼出元のFormを取得
        Dim originForm As Object = Me.Owner

        If originForm IsNot Nothing Then
            '' 呼出元のFormに表示されている画面の処理を起動する
            If TypeOf originForm Is CPBaseForm Then
                Dim mainForm As CPBaseForm = originForm
                Dim screen As CPBaseUserControl = mainForm.GetCurrentScreen()
                '' イベントを発生させる
                If screen IsNot Nothing Then
                    screen.CalledFromAddForm(Me.SendDataToBaseForm)
                    Me.SendDataToBaseForm = Nothing
                End If
            End If
        End If

        '' ユーザコントロールをすべて破棄する
        For Each component In Me.MainPanel.Controls
            If TypeOf component Is CPBaseUserControl Then
                Dim screen As CPBaseUserControl = component
                ''************** <メモリ開放> 2009/08/21 T.Iwasawa Update Start
                screen.Dispose()
                'screen.Close()
                ''************** <メモリ開放> 2009/08/21 T.Iwasawa Update End
            End If
        Next

        '' コンポーネントをすべて破棄する
        Me.Controls.Clear()

    End Sub

    ''' <summary>
    ''' 戻るボタンを押下したときの動作。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BackButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackButton.Click

        Try
            '' 変更ステータスをチェックする
            Me.CheckChangeStatus()
        Catch ex As CPProcessCanceledException
            '' キャンセルの場合、処理を抜ける
            Return
        End Try

        '' 戻る処理
        Dim currentScreen As CPBaseUserControl = Me.GetCurrentScreen
        If currentScreen IsNot Nothing Then
            currentScreen.ReturnToPreviousScreen()
        End If

    End Sub

    ''' <summary>
    ''' 閉じるボタンを押下したときの動作。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CloseButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseButton.Click

        Me.Close()

    End Sub

    ''' <summary>
    ''' 前の画面に戻る。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ReturnToPreviousScreen()

        '' 履歴がなければ処理を抜ける
        If Me.HistoryScreenId Is Nothing Then
            Return
        End If

        '' 戻り先の画面IDを取得する
        Dim backScreenId As String = Me.HistoryScreenId(Me.HistoryScreenId.Count - 1)

        '' 前の画面に戻る
        Me.DisplayScreen(backScreenId, True, Nothing)
    End Sub

    ''' <summary>
    ''' 変更ステータスをチェックし、変更がある場合は、ダイアログを表示する。キャンセルの場合はCPProcessCanceledExceptionが発生する。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CheckChangeStatus()

        '' 変更ステータスがOnの場合
        If Me.ChangeStatus Then
            '' 確認ダイアログを表示する
            If CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0016) = Windows.Forms.DialogResult.Cancel Then
                Throw New CPProcessCanceledException
            End If

        End If
    End Sub
```

[CPBaseUserControl.vb:89](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPBaseUserControl.vb>)～126

```vb
    Protected Sub ForwardScreen(ByRef targetScreen As String)

        '' 指定した画面に遷移する(遷移先の画面で初期表示処理を実行させる)
        Me.ForwardScreen(targetScreen, Nothing)
    End Sub

    ''' <summary>
    ''' 指定した画面に遷移する。遷移先の画面で初期表示処理が実行される。
    ''' </summary>
    ''' <param name="targetScreen">遷移先の画面ID</param>
    ''' <param name="deliveryData">引継ぎ情報</param>
    ''' <remarks></remarks>
    Protected Sub ForwardScreen(ByRef targetScreen As String, ByVal deliveryData As CPDeliveryData)

        '' 自身が追加されている元のフォームを取得する
        Dim thisForm As CPBaseForm = Me.FindForm

        '' 指定した画面に遷移する(遷移先の画面で初期表示処理を実行させる)
        thisForm.DisplayScreen(targetScreen, False, deliveryData)
    End Sub

    ''' <summary>
    ''' 指定した画面を再描画する。
    ''' </summary>
    ''' <param name="targetScreen">遷移先の画面ID</param>
    ''' <remarks></remarks>
    Protected Sub ReDrawScreen(ByRef targetScreen As String)

        '' 自身が追加されている元のフォームを取得する
        Dim thisForm As CPBaseForm = Me.FindForm

        '' 指定した画面を再表示する(遷移先の画面で初期表示処理を実行させない)
        thisForm.DisplayScreen(targetScreen, True, Nothing)
    End Sub

    ''' <summary>
    ''' 新しくフォームを起動する。
    ''' </summary>
```

[CPBaseUserControl.vb:170](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPBaseUserControl.vb>)～211

```vb
    ''' <summary>
    ''' 新規で立ち上げたFormが閉じられるタイミングで呼び出される。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CalledFromAddForm(ByRef delivery As CPDeliveryData)
        '' イベント引数を作成
        Dim e As CPUserControlCalledFromAddFormEventArgs = New CPUserControlCalledFromAddFormEventArgs()
        e.DeliveryData = delivery

        Try
            '' イベントを起動する
            RaiseEvent CalledFromForwardScreen(e)

            '' Daoがある場合はコミット処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Commit()
            End If

        Catch ex As Exception
            '' Daoがある場合はロールバック処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Rollback()
            End If

            '' Exceptionごとの処理
            CPComponentUtility.HandleException(ex)

        Finally
            '' Daoがある場合はクローズ処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Close()
                dao = Nothing
                CPDaoContainer.SetDao(Nothing)
            End If

            '' システム日付をクリア
            CPSettingInfo.ClearSystemDate()
        End Try
    End Sub
```

</details>

<a id="C12"></a>

### C12　保存・取消・例外とDB更新単位

条件：イベント正常終了／例外、DAO有無、処理件数、DBConcurrencyException、重複制約。

CPButtonの正常終了はDAOをコミットし、DB更新件数>0なら変更フラグを初期化。伝播例外はロールバック後に種類別通知。単なるExit・Return・Falseは例外ではないため、先行DB更新があればコミット可能。外部ファイルやメールはDBトランザクションでは戻らない。排他の発生条件は各XSD更新WHEREと実DB照合が必要。

<details>
<summary>原文根拠</summary>

[CPButton.vb:51](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPButton.vb>)～103

```vb
    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)

        Dim preCursor As Cursor = Cursor.Current
        Cursor.Current = Cursors.WaitCursor

        Try
            '' クリックのイベントを起動する
            MyBase.OnClick(e)

            '' Daoがある場合はコミット処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Commit()
            End If

            '' 更新処理後は変更フラグをリセット
            Dim form As CPBaseForm = Me.FindForm
            If form IsNot Nothing Then
                If form.GetUpdateCount > 0 Then
                    '' ログ出力
                    CPLog.Info(Me.Name & "ボタン実行")

                    form.InitChangeStatus()
                    form.InitUpdateCount()
                End If
            End If

        Catch ex As Exception
            '' Daoがある場合はロールバック処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Rollback()
            End If

            '' Exceptionごとの処理
            CPComponentUtility.HandleException(ex)

        Finally
            '' Daoがある場合はクローズ処理
            Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
            If dao IsNot Nothing Then
                dao.Close()
                dao = Nothing
                CPDaoContainer.SetDao(Nothing)
            End If

            '' システム日付をクリア
            CPSettingInfo.ClearSystemDate()

            Cursor.Current = preCursor
        End Try

    End Sub
```

[CPComponentUtility.vb:21](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPComponentUtility.vb>)～62

```vb
    Public Shared Sub HandleException(ByRef ex As Exception)

        CPLog.Debug(ex.ToString)

        If TypeOf ex Is CPInputCheckException Then
            '' 入力エラーの場合
            Return
        ElseIf TypeOf ex Is CPAuthorityException Then
            '' 権限エラーの場合
            Return
        ElseIf TypeOf ex Is CPProcessCanceledException Then
            '' 処理キャンセルの場合
            Return
        ElseIf TypeOf ex Is CPDataNotChangedException Then
            '' 画面上のデータが変更されていない場合
            CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0010)
            Return
        ElseIf TypeOf ex Is DBConcurrencyException Then
            '' 排他エラーの場合
            CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0009)
            Return
        ElseIf TypeOf ex Is System.Reflection.TargetInvocationException Then
            Dim innerException As Exception = ex.InnerException

            If TypeOf innerException Is DBConcurrencyException Then
                '' 排他エラーの場合
                CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0009)
                Return
            End If

            '' 重複エラーハンドリング
            Dim oex As System.Data.OracleClient.OracleException = CType(innerException, System.Data.OracleClient.OracleException)
            If oex.Code = 1 Then
                CPMessageUtility.DisplayMessage(CPMessageConstant.FWM_0020)
                Return
            End If
        End If

        '' 上記以外は再度throw
        CPLog.Err(ex)
        Throw ex
    End Sub
```

[CPCoreUserControl.vb:423](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPCoreUserControl.vb>)～459

```vb
    Protected Function UpdateAll(ByRef dataTable As DataTable) As Integer
        Dim count As Integer = CPDBUtility.Update(Me.GetTableAdapter(dataTable), dataTable, Me.GetUpdateUser, Me.SystemDate)
        Me.SetUpdateCount(count)
        Return count
    End Function

    ''' <summary>
    ''' 削除/更新/追加を行う。プライマリーキーを自動インクリメントする。
    ''' </summary>
    ''' <param name="dataTable">DataTable</param>
    ''' <param name="primaryKey">対象テーブルのプライマリーキーのID</param>
    ''' <returns>変更が反映された行数</returns>
    ''' <remarks></remarks>
    Protected Function UpdateAll(ByRef dataTable As DataTable, ByVal primaryKey As String) As Integer
        '' 処理件数格納用
        Dim count As Integer = 0

        '' 削除処理
        count = count + Me.DeleteRows(dataTable)

        '' 更新処理
        count = count + Me.UpdateRows(dataTable)

        '' プライマリーキーの最大値＋１をDBから取得する()
        Dim maxValue As Decimal = CPDBUtility.GetMaxValue(dataTable, primaryKey) + 1

        '' 追加対象行を取得する
        Dim insertRow As DataRow() = CPDBUtility.GetInsertRows(dataTable)

        '' 新規対象行をループさせ、行ごとにプライマリーキーをインクリメントし、追加処理を行う
        For Each row As DataRow In insertRow
            '' プライマリーキーをセット
            row.Item(primaryKey) = maxValue
            '' 追加処理
            count = count + Me.InsertRow(row)
            '' プライマリーキーをインクリメント
            maxValue = maxValue + 1
```

</details>

<a id="C13"></a>

### C13　捺印対象・Office失敗とFalse

条件：パス空、拡張子大小文字、.xlsx／その他、起動・Open・Write例外、SaveAs失敗、解放対象有無。

パス空は例外。.xlsxのみOpen・Write・SaveAs。他形式は何も押印せずTrue。SaveAs失敗を捕捉してFalse、それ以外の例外はFinallyで解放後に伝播。最終解放でも例外が起こればFalseだけとみなせない。

<details>
<summary>原文根拠</summary>

[CPExcelManager.vb:542](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/EXCEL/CPExcelManager.vb>)～616

```vb
    Public Function UpdateExcelReport(ByVal targetFilePath As String) As Boolean

        '' Excelアプリケーション操作用
        Dim excelApp As Excel.Application = Nothing
        '' Book管理用
        Dim xlsBooks As Excel.Workbooks = Nothing
        '' テンプレート用Book ※実際のBookを操作
        Dim baseBook As Excel.Workbook = Nothing
        '' テンプレート用Sheet ※すべてのシート
        Dim baseSheets As Excel.Sheets = Nothing

        Try
            '' ファイルの指定がない場合はエラーとする
            If CPCommonUtility.IsNull(targetFilePath) Then
                Throw New CPSpecifiedException("ファイルが指定されていません")
            End If

            '' ファイルの種類がEXCELでない場合は処理を終了する
            Dim fileType As String = System.IO.Path.GetExtension(targetFilePath).ToLower
            If Not fileType.Equals(CPConstant.FILETYPE_EXCEL) Then
                Return True
            End If

            '' Excel操作用オブジェクトを作成
            excelApp = New Excel.Application()
            '' Book操作用オブジェクトを作成
            xlsBooks = excelApp.Workbooks

            '' Excelオブジェクトを取得
            baseBook = xlsBooks.Open(targetFilePath, 0, False)
            'baseBook = xlsBooks.Open(CPSettingInfo.GetExcelTemplatePath & "\" & TemplateFileName, 0, False)
            baseSheets = baseBook.Sheets

            '' シートへの書き込み処理　各サブクラスでオーバーライドされている
            Me.WriteDataInExcel(baseSheets)

            '' Excelを起動するかどうか
            'excelApp.Visible = DisplayReportValue

            '' アラートを回避
            excelApp.DisplayAlerts = False
            '' ファイルを保存して閉じる
            Try
                'baseBook.Close(SaveChanges:=True)
                baseBook.SaveAs(Filename:=targetFilePath)
            Catch ex As Exception
                CPMessageUtility.DisplayMessage(ex.Message)
                Return False
            End Try
            '' アラートの設定を戻す
            excelApp.DisplayAlerts = True

        Finally
            If Not baseBook Is Nothing Then
                baseBook.Close()
            End If
            If Not excelApp Is Nothing Then
                excelApp.Quit()
            End If
            '' オブジェクトをそれぞれ開放
            If Not baseBook Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(baseBook)
            End If
            If Not xlsBooks Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBooks)
            End If
            If Not excelApp Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End If
            '' Sheet/Rangeなどを自動開放
            System.GC.Collect()
        End Try

        Return True

```

[CPConstant.vb:180](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONSTANT/CPConstant.vb>)～187

```vb
    ''' <remarks></remarks>
    Public Const FORMAT_YYYYMM_NO_SLASH As String = "yyyyMM"

    ''' <summary>
    ''' EXCELファイルの拡張子
    ''' </summary>
    ''' <remarks></remarks>
    Public Const FILETYPE_EXCEL As String = ".xlsx"
```

</details>

<a id="C14"></a>

### C14　SQL条件未指定と値の型

条件：置換なし固定条件／置換引数、空値・値あり、通常置換／特殊置換、文字・整数・Decimal・DateTime・文字配列。

値が空ならGetReplaceStringはNothing。固定CONDITIONは採用、単一引数条件は値なしで省略。文字はクォートとエスケープ、数値は値、日付は日単位TO_DATE、配列はカンマ連結。複数引数を同じCONDITIONに持つ場合は有効フラグが引数ごとに更新されるので、一律「一つ空なら必ず省略」と解釈しない。3003Uは開始・終了・差分開始・差分終了が別CONDITIONなので各々省略する。

<details>
<summary>原文根拠</summary>

[CPDBUtility.vb:810](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPDBUtility.vb>)～902

```vb
        '' Conditionタグをループ
        For i As Integer = 0 To whereList.Count - 1
            '' 検索条件
            Dim whereString As String = whereList(i).Condition
            '' Prefix
            Dim prefix As String = whereList(i).Prefix
            '' Suffix
            Dim suffix As String = whereList(i).Suffix
            '' 結果格納用
            Dim resultString As String = Nothing
            '' 検索条件有効フラグ
            Dim isEffectiveFlag As Boolean = False

            '' 検索条件には置換え文字がある場合、":{0}"というような文字が記述されている
            If whereString.IndexOf(":") < 0 Then
                '' 置換え文字がない場合、検索条件有効フラグをON
                isEffectiveFlag = True
                '' 結果文字列に検索条件をセット
                resultString = whereString
            Else
                resultString = whereString

                '' 指定された条件の値をループ
                For j As Integer = 0 To sqlArgs.Length - 1
                    '' 置換え文字をインデックスより作成
                    With Nothing
                        Dim pos As String = ":{" & CType(j, String) & "}"

                        '' 置換え文字があるかどうか
                        If selectSql.IndexOf(pos) < 0 Then
                            '' 置換え文字がない場合
                        Else
                            '' 置換え文字がある場合、検索条件の置換え文字と条件の値を置き換える
                            selectSql = GetReplaceString(selectSql, pos, sqlArgs(j), Nothing, Nothing)
                        End If

                        '' 置換え文字があるかどうか
                        If whereString.IndexOf(pos) < 0 Then
                            '' 置換え文字がない場合
                        Else
                            '' 置換え文字がある場合、検索条件の置換え文字と条件の値を置き換える
                            resultString = GetReplaceString(resultString, pos, sqlArgs(j), prefix, suffix)

                            '' 結果があれば有効フラグをON
                            If resultString IsNot Nothing Then
                                isEffectiveFlag = True
                            End If
                        End If
                    End With

                    With Nothing
                        Dim pos As String = ":${" & CType(j, String) & "}"

                        '' 置換え文字があるかどうか
                        If selectSql.IndexOf(pos) < 0 Then
                            '' 置換え文字がない場合
                        Else
                            '' 置換え文字がある場合、検索条件の置換え文字と条件の値を置き換える
                            selectSql = GetReplaceStringSpecial(selectSql, pos, sqlArgs(j), Nothing, Nothing)
                        End If

                        '' 置換え文字があるかどうか
                        If whereString.IndexOf(pos) < 0 Then
                            '' 置換え文字がない場合
                        Else
                            '' 置換え文字がある場合、検索条件の置換え文字と条件の値を置き換える
                            resultString = GetReplaceStringSpecial(resultString, pos, sqlArgs(j), prefix, suffix)

                            '' 結果があれば有効フラグをON
                            If resultString IsNot Nothing Then
                                isEffectiveFlag = True
                            End If
                        End If
                    End With

                Next

            End If

            If isEffectiveFlag Then
                '' 条件が有効の場合は追加する
                If isWhereStatementFlag Then
                    '' 最初の条件は先頭にWHEREを付加
                    selectSql = selectSql & " WHERE "
                    isWhereStatementFlag = False
                Else
                    '' 次の条件は先頭にANDを付加
                    selectSql = selectSql & " AND "
                End If
                selectSql = selectSql & resultString
            Else
                '' 条件が有効でない場合は、条件から外す
            End If
```

[CPDBUtility.vb:920](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPDBUtility.vb>)～979

```vb
    Private Shared Function GetReplaceString(ByVal sqlString As String, ByVal pos As String, _
           ByVal value As Object, ByVal prefix As String, ByVal suffix As String) As String

        '' 結果格納用
        Dim resultString As String = Nothing

        '' 置換え文字がない場合は処理を抜ける
        If CPCommonUtility.IsNull(value) Then
            Return Nothing
        End If

        '' 置換え文字の型を判断
        If TypeOf value Is String Then
            '' 文字の場合は"''"を付加
            Dim replace As String = prefix & value.ToString.Replace("'", "''") & suffix
            resultString = sqlString.Replace(pos, "'" & replace & "'")
        ElseIf TypeOf value Is Integer Then
            resultString = sqlString.Replace(pos, value)
        ElseIf TypeOf value Is Decimal Then
            resultString = sqlString.Replace(pos, value)
        ElseIf TypeOf value Is DateTime Then
            Dim dateValue As DateTime = value
            Dim year As String = dateValue.Year.ToString
            Dim month As String = dateValue.Month.ToString.PadLeft(2, "0")
            Dim day As String = dateValue.Day.ToString.PadLeft(2, "0")

            resultString = sqlString.Replace(pos, "TO_DATE('" & year & month & day & "','YYYYMMDD')")
        ElseIf TypeOf value Is String() Then
            Dim valueString As String = String.Join(",", value)
            resultString = sqlString.Replace(pos, valueString)

        End If

        Return resultString
    End Function

    ''' <summary>
    ''' 置換え文字を埋め込む
    ''' </summary>
    ''' <param name="sqlString">対象のSQLの文字列</param>
    ''' <param name="pos">置き換える元の記号</param>
    ''' <param name="value">置き換える値</param>
    ''' <param name="prefix">プリフィックス</param>
    ''' <param name="suffix">サフィックス</param>
    ''' <returns>置換え後の文字列</returns>
    ''' <remarks></remarks>
    Private Shared Function GetReplaceStringSpecial(ByVal sqlString As String, ByVal pos As String, _
           ByVal value As Object, ByVal prefix As String, ByVal suffix As String) As String

        '' 結果格納用
        Dim resultString As String = Nothing

        '' 置換え文字がない場合は処理を抜ける
        If CPCommonUtility.IsNull(value) Then
            Return Nothing
        End If

        resultString = sqlString.Replace(pos, value)

        Return resultString
```

</details>

<a id="C15"></a>

### C15　ファイル選択取消と更新イベント

条件：ファイル選択OK／Cancel、DAO有無、イベント例外、更新件数。

CPFileDialogButtonは選択OKの場合だけMyBase.OnClickを実行し、保存イベントへ進む。Cancelではファイル追加処理を呼ばない。正常・例外のDB処理もCPButtonと同系統。存在・同名・承認状態は各画面の条件を併せて確認。

<details>
<summary>原文根拠</summary>

[CPFileDialogButton.vb:52](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPFileDialogButton.vb>)～115

```vb
    Private openDialog As OpenFileDialog

    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        Me.openDialog = New OpenFileDialog
        '' ファイルの種類のフィルターをセット
        'Me.openDialog.Filter = "Microsoft Office Excelブック(*.xls)|*.xls"
        '************** 2009/09/08 M.Ohsuka Update Start
        'Me.openDialog.Filter = "Microsoft Office Excelブック(*.xls)|*.xls|すべてのファイル(*.*)|*.*"
        Me.openDialog.Filter = "すべてのファイル(*.*)|*.*"
        '************** 2009/09/08 M.Ohsuka Update End

        If Me.openDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then

            'MyBase.OnClick(e)

            Try
                '' クリックのイベントを起動する
                MyBase.OnClick(e)

                '' Daoがある場合はコミット処理
                Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
                If dao IsNot Nothing Then
                    dao.Commit()
                End If

                '' 更新処理後は変更フラグをリセット
                Dim form As CPBaseForm = Me.FindForm
                If form IsNot Nothing Then
                    If form.GetUpdateCount > 0 Then
                        '' ログ出力
                        CPLog.Info(Me.Name & "ボタン実行")

                        form.InitChangeStatus()
                        form.InitUpdateCount()
                    End If
                End If

            Catch ex As Exception
                '' Daoがある場合はロールバック処理
                Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
                If dao IsNot Nothing Then
                    dao.Rollback()
                End If

                '' Exceptionごとの処理
                CPComponentUtility.HandleException(ex)

            Finally
                '' Daoがある場合はクローズ処理
                Dim dao As CPDBAccessObject = CPDaoContainer.GetDao
                If dao IsNot Nothing Then
                    dao.Close()
                    dao = Nothing
                    CPDaoContainer.SetDao(Nothing)
                End If

                '' システム日付をクリア
                CPSettingInfo.ClearSystemDate()
            End Try

        End If

    End Sub

```

</details>

## メソッドと業務項目の対応

全549メソッド・プロパティを、業務要約へ対応する処理と、分岐のない保持プロパティへ分類しました。1,054件の構文分岐・反復はすべて業務項目に対応します。ただし、複合条件の全真偽組合せやテストケースを網羅したという意味ではありません。

| 根拠ID | 対応機能ID | 扱い | 分岐・反復数 |
| --- | --- | --- | --- |
| [F1005-02-001](VB条件分岐_原文根拠.md#F1005-02-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-002](VB条件分岐_原文根拠.md#F1005-02-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-003](VB条件分岐_原文根拠.md#F1005-02-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-004](VB条件分岐_原文根拠.md#F1005-02-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-005](VB条件分岐_原文根拠.md#F1005-02-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-006](VB条件分岐_原文根拠.md#F1005-02-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-007](VB条件分岐_原文根拠.md#F1005-02-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-02-008](VB条件分岐_原文根拠.md#F1005-02-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-001](VB条件分岐_原文根拠.md#F1005-03-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-002](VB条件分岐_原文根拠.md#F1005-03-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-003](VB条件分岐_原文根拠.md#F1005-03-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-004](VB条件分岐_原文根拠.md#F1005-03-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-005](VB条件分岐_原文根拠.md#F1005-03-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-006](VB条件分岐_原文根拠.md#F1005-03-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-007](VB条件分岐_原文根拠.md#F1005-03-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-008](VB条件分岐_原文根拠.md#F1005-03-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-009](VB条件分岐_原文根拠.md#F1005-03-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-03-010](VB条件分岐_原文根拠.md#F1005-03-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-001](VB条件分岐_原文根拠.md#F1005-04-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-002](VB条件分岐_原文根拠.md#F1005-04-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-003](VB条件分岐_原文根拠.md#F1005-04-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-004](VB条件分岐_原文根拠.md#F1005-04-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-005](VB条件分岐_原文根拠.md#F1005-04-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-006](VB条件分岐_原文根拠.md#F1005-04-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-007](VB条件分岐_原文根拠.md#F1005-04-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-008](VB条件分岐_原文根拠.md#F1005-04-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-009](VB条件分岐_原文根拠.md#F1005-04-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-010](VB条件分岐_原文根拠.md#F1005-04-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-011](VB条件分岐_原文根拠.md#F1005-04-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-012](VB条件分岐_原文根拠.md#F1005-04-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-013](VB条件分岐_原文根拠.md#F1005-04-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-014](VB条件分岐_原文根拠.md#F1005-04-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-015](VB条件分岐_原文根拠.md#F1005-04-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-016](VB条件分岐_原文根拠.md#F1005-04-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-017](VB条件分岐_原文根拠.md#F1005-04-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-018](VB条件分岐_原文根拠.md#F1005-04-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-019](VB条件分岐_原文根拠.md#F1005-04-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-020](VB条件分岐_原文根拠.md#F1005-04-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-021](VB条件分岐_原文根拠.md#F1005-04-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-04-022](VB条件分岐_原文根拠.md#F1005-04-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-001](VB条件分岐_原文根拠.md#F1005-05-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-002](VB条件分岐_原文根拠.md#F1005-05-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-003](VB条件分岐_原文根拠.md#F1005-05-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-004](VB条件分岐_原文根拠.md#F1005-05-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-005](VB条件分岐_原文根拠.md#F1005-05-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-006](VB条件分岐_原文根拠.md#F1005-05-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-007](VB条件分岐_原文根拠.md#F1005-05-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-008](VB条件分岐_原文根拠.md#F1005-05-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-009](VB条件分岐_原文根拠.md#F1005-05-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-010](VB条件分岐_原文根拠.md#F1005-05-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-011](VB条件分岐_原文根拠.md#F1005-05-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-012](VB条件分岐_原文根拠.md#F1005-05-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-013](VB条件分岐_原文根拠.md#F1005-05-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-014](VB条件分岐_原文根拠.md#F1005-05-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-015](VB条件分岐_原文根拠.md#F1005-05-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-016](VB条件分岐_原文根拠.md#F1005-05-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-017](VB条件分岐_原文根拠.md#F1005-05-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-018](VB条件分岐_原文根拠.md#F1005-05-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-019](VB条件分岐_原文根拠.md#F1005-05-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-020](VB条件分岐_原文根拠.md#F1005-05-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-021](VB条件分岐_原文根拠.md#F1005-05-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-022](VB条件分岐_原文根拠.md#F1005-05-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-05-023](VB条件分岐_原文根拠.md#F1005-05-023) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-001](VB条件分岐_原文根拠.md#F1005-06-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-002](VB条件分岐_原文根拠.md#F1005-06-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-003](VB条件分岐_原文根拠.md#F1005-06-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-004](VB条件分岐_原文根拠.md#F1005-06-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-005](VB条件分岐_原文根拠.md#F1005-06-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-006](VB条件分岐_原文根拠.md#F1005-06-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-007](VB条件分岐_原文根拠.md#F1005-06-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-008](VB条件分岐_原文根拠.md#F1005-06-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-009](VB条件分岐_原文根拠.md#F1005-06-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-010](VB条件分岐_原文根拠.md#F1005-06-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-011](VB条件分岐_原文根拠.md#F1005-06-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-012](VB条件分岐_原文根拠.md#F1005-06-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-013](VB条件分岐_原文根拠.md#F1005-06-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-014](VB条件分岐_原文根拠.md#F1005-06-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-015](VB条件分岐_原文根拠.md#F1005-06-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-016](VB条件分岐_原文根拠.md#F1005-06-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-017](VB条件分岐_原文根拠.md#F1005-06-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-018](VB条件分岐_原文根拠.md#F1005-06-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-019](VB条件分岐_原文根拠.md#F1005-06-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-020](VB条件分岐_原文根拠.md#F1005-06-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-021](VB条件分岐_原文根拠.md#F1005-06-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-022](VB条件分岐_原文根拠.md#F1005-06-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-023](VB条件分岐_原文根拠.md#F1005-06-023) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-024](VB条件分岐_原文根拠.md#F1005-06-024) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-06-025](VB条件分岐_原文根拠.md#F1005-06-025) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-001](VB条件分岐_原文根拠.md#F1005-09-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-002](VB条件分岐_原文根拠.md#F1005-09-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-003](VB条件分岐_原文根拠.md#F1005-09-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-004](VB条件分岐_原文根拠.md#F1005-09-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-005](VB条件分岐_原文根拠.md#F1005-09-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-006](VB条件分岐_原文根拠.md#F1005-09-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-007](VB条件分岐_原文根拠.md#F1005-09-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-008](VB条件分岐_原文根拠.md#F1005-09-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-009](VB条件分岐_原文根拠.md#F1005-09-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-010](VB条件分岐_原文根拠.md#F1005-09-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-011](VB条件分岐_原文根拠.md#F1005-09-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-012](VB条件分岐_原文根拠.md#F1005-09-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-013](VB条件分岐_原文根拠.md#F1005-09-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-014](VB条件分岐_原文根拠.md#F1005-09-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-015](VB条件分岐_原文根拠.md#F1005-09-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-016](VB条件分岐_原文根拠.md#F1005-09-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-017](VB条件分岐_原文根拠.md#F1005-09-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-018](VB条件分岐_原文根拠.md#F1005-09-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-019](VB条件分岐_原文根拠.md#F1005-09-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-020](VB条件分岐_原文根拠.md#F1005-09-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-021](VB条件分岐_原文根拠.md#F1005-09-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-09-022](VB条件分岐_原文根拠.md#F1005-09-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-11-001](VB条件分岐_原文根拠.md#F1005-11-001) | VB1005-01, VB1005-04 | 業務項目に対応 | 3 |
| [F1005-11-002](VB条件分岐_原文根拠.md#F1005-11-002) | VB1005-06, VB1005-07, VB1005-08, VB1005-09, VB1005-10 | 業務項目に対応 | 22 |
| [F1005-11-003](VB条件分岐_原文根拠.md#F1005-11-003) | VB1005-02 | 業務項目に対応 | 6 |
| [F1005-11-004](VB条件分岐_原文根拠.md#F1005-11-004) | VB1005-05 | 業務項目に対応 | 0 |
| [F1005-11-005](VB条件分岐_原文根拠.md#F1005-11-005) | VB1005-05 | 業務項目に対応 | 16 |
| [F1005-11-006](VB条件分岐_原文根拠.md#F1005-11-006) | VB1005-12 | 業務項目に対応 | 3 |
| [F1005-11-007](VB条件分岐_原文根拠.md#F1005-11-007) | VB1005-12, VB1005-13, VB1005-14, VB1005-15 | 業務項目に対応 | 38 |
| [F1005-11-008](VB条件分岐_原文根拠.md#F1005-11-008) | VB1005-11 | 業務項目に対応 | 0 |
| [F1005-11-009](VB条件分岐_原文根拠.md#F1005-11-009) | VB1005-16 | 業務項目に対応 | 3 |
| [F1005-11-010](VB条件分岐_原文根拠.md#F1005-11-010) | VB1005-17 | 業務項目に対応 | 17 |
| [F1005-11-011](VB条件分岐_原文根拠.md#F1005-11-011) | VB1005-16 | 業務項目に対応 | 27 |
| [F1005-11-012](VB条件分岐_原文根拠.md#F1005-11-012) | VB1005-28 | 業務項目に対応 | 4 |
| [F1005-11-013](VB条件分岐_原文根拠.md#F1005-11-013) | VB1005-28 | 業務項目に対応 | 2 |
| [F1005-11-014](VB条件分岐_原文根拠.md#F1005-11-014) | VB1005-28 | 業務項目に対応 | 2 |
| [F1005-11-015](VB条件分岐_原文根拠.md#F1005-11-015) | VB1005-15 | 業務項目に対応 | 1 |
| [F1005-11-016](VB条件分岐_原文根拠.md#F1005-11-016) | VB1005-12 | 業務項目に対応 | 3 |
| [F1005-11-017](VB条件分岐_原文根拠.md#F1005-11-017) | VB1005-15 | 業務項目に対応 | 0 |
| [F1005-11-018](VB条件分岐_原文根拠.md#F1005-11-018) | VB1005-20 | 業務項目に対応 | 5 |
| [F1005-11-019](VB条件分岐_原文根拠.md#F1005-11-019) | VB1005-20 | 業務項目に対応 | 5 |
| [F1005-11-020](VB条件分岐_原文根拠.md#F1005-11-020) | VB1005-29 | 業務項目に対応 | 1 |
| [F1005-11-021](VB条件分岐_原文根拠.md#F1005-11-021) | VB1005-29 | 業務項目に対応 | 1 |
| [F1005-11-022](VB条件分岐_原文根拠.md#F1005-11-022) | VB1005-20 | 業務項目に対応 | 2 |
| [F1005-11-023](VB条件分岐_原文根拠.md#F1005-11-023) | VB1005-20 | 業務項目に対応 | 2 |
| [F1005-11-024](VB条件分岐_原文根拠.md#F1005-11-024) | VB1005-20 | 業務項目に対応 | 2 |
| [F1005-11-025](VB条件分岐_原文根拠.md#F1005-11-025) | VB1005-20 | 業務項目に対応 | 2 |
| [F1005-11-026](VB条件分岐_原文根拠.md#F1005-11-026) | VB1005-20 | 業務項目に対応 | 2 |
| [F1005-11-027](VB条件分岐_原文根拠.md#F1005-11-027) | VB1005-04 | 業務項目に対応 | 0 |
| [F1005-13-001](VB条件分岐_原文根拠.md#F1005-13-001) | VB1005-29 | 業務項目に対応 | 4 |
| [F1005-13-002](VB条件分岐_原文根拠.md#F1005-13-002) | VB1005-03, VB1005-04, VB1005-18 | 業務項目に対応 | 10 |
| [F1005-13-003](VB条件分岐_原文根拠.md#F1005-13-003) | VB1005-05 | 業務項目に対応 | 0 |
| [F1005-13-004](VB条件分岐_原文根拠.md#F1005-13-004) | VB1005-21 | 業務項目に対応 | 9 |
| [F1005-13-005](VB条件分岐_原文根拠.md#F1005-13-005) | VB1005-22 | 業務項目に対応 | 15 |
| [F1005-13-006](VB条件分岐_原文根拠.md#F1005-13-006) | VB1005-23 | 業務項目に対応 | 17 |
| [F1005-13-007](VB条件分岐_原文根拠.md#F1005-13-007) | VB1005-24 | 業務項目に対応 | 19 |
| [F1005-13-008](VB条件分岐_原文根拠.md#F1005-13-008) | VB1005-26 | 業務項目に対応 | 6 |
| [F1005-13-009](VB条件分岐_原文根拠.md#F1005-13-009) | VB1005-05 | 業務項目に対応 | 7 |
| [F1005-13-010](VB条件分岐_原文根拠.md#F1005-13-010) | VB1005-11 | 業務項目に対応 | 29 |
| [F1005-13-011](VB条件分岐_原文根拠.md#F1005-13-011) | VB1005-09 | 業務項目に対応 | 1 |
| [F1005-13-012](VB条件分岐_原文根拠.md#F1005-13-012) | VB1005-07 | 業務項目に対応 | 8 |
| [F1005-13-013](VB条件分岐_原文根拠.md#F1005-13-013) | VB1005-18 | 業務項目に対応 | 2 |
| [F1005-13-014](VB条件分岐_原文根拠.md#F1005-13-014) | VB1005-18 | 業務項目に対応 | 5 |
| [F1005-13-015](VB条件分岐_原文根拠.md#F1005-13-015) | VB1005-25 | 業務項目に対応 | 17 |
| [F1005-13-016](VB条件分岐_原文根拠.md#F1005-13-016) | VB1005-25 | 業務項目に対応 | 2 |
| [F1005-13-017](VB条件分岐_原文根拠.md#F1005-13-017) | VB1005-10 | 業務項目に対応 | 0 |
| [F1005-13-018](VB条件分岐_原文根拠.md#F1005-13-018) | VB1005-27 | 業務項目に対応 | 2 |
| [F1005-13-019](VB条件分岐_原文根拠.md#F1005-13-019) | VB1005-27 | 業務項目に対応 | 3 |
| [F1005-13-020](VB条件分岐_原文根拠.md#F1005-13-020) | VB1005-27 | 業務項目に対応 | 1 |
| [F1005-13-021](VB条件分岐_原文根拠.md#F1005-13-021) | VB1005-22 | 業務項目に対応 | 1 |
| [F1005-13-022](VB条件分岐_原文根拠.md#F1005-13-022) | VB1005-27 | 業務項目に対応 | 0 |
| [F1005-13-023](VB条件分岐_原文根拠.md#F1005-13-023) | VB1005-14 | 業務項目に対応 | 33 |
| [F1005-13-024](VB条件分岐_原文根拠.md#F1005-13-024) | VB1005-19 | 業務項目に対応 | 5 |
| [F1005-13-025](VB条件分岐_原文根拠.md#F1005-13-025) | VB1005-20 | 業務項目に対応 | 3 |
| [F1005-13-026](VB条件分岐_原文根拠.md#F1005-13-026) | VB1005-20 | 業務項目に対応 | 0 |
| [F1005-13-027](VB条件分岐_原文根拠.md#F1005-13-027) | VB1005-20 | 業務項目に対応 | 3 |
| [F1005-13-028](VB条件分岐_原文根拠.md#F1005-13-028) | VB1005-20 | 業務項目に対応 | 0 |
| [F1005-13-029](VB条件分岐_原文根拠.md#F1005-13-029) | VB1005-14 | 業務項目に対応 | 5 |
| [F1005-13-030](VB条件分岐_原文根拠.md#F1005-13-030) | VB1005-26 | 業務項目に対応 | 0 |
| [F1005-13-031](VB条件分岐_原文根拠.md#F1005-13-031) | VB1005-26 | 業務項目に対応 | 1 |
| [F1005-13-032](VB条件分岐_原文根拠.md#F1005-13-032) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-033](VB条件分岐_原文根拠.md#F1005-13-033) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-034](VB条件分岐_原文根拠.md#F1005-13-034) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-035](VB条件分岐_原文根拠.md#F1005-13-035) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-036](VB条件分岐_原文根拠.md#F1005-13-036) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-037](VB条件分岐_原文根拠.md#F1005-13-037) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-038](VB条件分岐_原文根拠.md#F1005-13-038) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-039](VB条件分岐_原文根拠.md#F1005-13-039) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-040](VB条件分岐_原文根拠.md#F1005-13-040) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-041](VB条件分岐_原文根拠.md#F1005-13-041) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-042](VB条件分岐_原文根拠.md#F1005-13-042) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-043](VB条件分岐_原文根拠.md#F1005-13-043) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-044](VB条件分岐_原文根拠.md#F1005-13-044) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-045](VB条件分岐_原文根拠.md#F1005-13-045) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-046](VB条件分岐_原文根拠.md#F1005-13-046) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-047](VB条件分岐_原文根拠.md#F1005-13-047) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-048](VB条件分岐_原文根拠.md#F1005-13-048) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1005-13-049](VB条件分岐_原文根拠.md#F1005-13-049) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-03-001](VB条件分岐_原文根拠.md#F1006-03-001) | VB1006-01 | 業務項目に対応 | 4 |
| [F1006-03-002](VB条件分岐_原文根拠.md#F1006-03-002) | VB1006-02, VB1006-03, VB1006-04 | 業務項目に対応 | 8 |
| [F1006-03-003](VB条件分岐_原文根拠.md#F1006-03-003) | VB1006-04 | 業務項目に対応 | 18 |
| [F1006-11-001](VB条件分岐_原文根拠.md#F1006-11-001) | VB1006-05 | 業務項目に対応 | 25 |
| [F1006-11-002](VB条件分岐_原文根拠.md#F1006-11-002) | VB1006-09, VB1006-10 | 業務項目に対応 | 12 |
| [F1006-11-003](VB条件分岐_原文根拠.md#F1006-11-003) | VB1006-08 | 業務項目に対応 | 18 |
| [F1006-11-004](VB条件分岐_原文根拠.md#F1006-11-004) | VB1006-08 | 業務項目に対応 | 2 |
| [F1006-11-005](VB条件分岐_原文根拠.md#F1006-11-005) | VB1006-11 | 業務項目に対応 | 0 |
| [F1006-11-006](VB条件分岐_原文根拠.md#F1006-11-006) | VB1006-05 | 業務項目に対応 | 3 |
| [F1006-11-007](VB条件分岐_原文根拠.md#F1006-11-007) | VB1006-05 | 業務項目に対応 | 3 |
| [F1006-11-008](VB条件分岐_原文根拠.md#F1006-11-008) | VB1006-06 | 業務項目に対応 | 8 |
| [F1006-11-009](VB条件分岐_原文根拠.md#F1006-11-009) | VB1006-07 | 業務項目に対応 | 6 |
| [F1006-11-010](VB条件分岐_原文根拠.md#F1006-11-010) | VB1006-07 | 業務項目に対応 | 6 |
| [F1006-11-011](VB条件分岐_原文根拠.md#F1006-11-011) | VB1006-07 | 業務項目に対応 | 12 |
| [F1006-11-012](VB条件分岐_原文根拠.md#F1006-11-012) | VB1006-11 | 業務項目に対応 | 1 |
| [F1006-11-013](VB条件分岐_原文根拠.md#F1006-11-013) | VB1006-11 | 業務項目に対応 | 1 |
| [F1006-11-014](VB条件分岐_原文根拠.md#F1006-11-014) | VB1006-11 | 業務項目に対応 | 1 |
| [F1006-13-001](VB条件分岐_原文根拠.md#F1006-13-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-13-002](VB条件分岐_原文根拠.md#F1006-13-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-13-003](VB条件分岐_原文根拠.md#F1006-13-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-13-004](VB条件分岐_原文根拠.md#F1006-13-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-001](VB条件分岐_原文根拠.md#F1006-14-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-002](VB条件分岐_原文根拠.md#F1006-14-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-003](VB条件分岐_原文根拠.md#F1006-14-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-004](VB条件分岐_原文根拠.md#F1006-14-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-005](VB条件分岐_原文根拠.md#F1006-14-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-006](VB条件分岐_原文根拠.md#F1006-14-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-007](VB条件分岐_原文根拠.md#F1006-14-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-14-008](VB条件分岐_原文根拠.md#F1006-14-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-15-001](VB条件分岐_原文根拠.md#F1006-15-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-15-002](VB条件分岐_原文根拠.md#F1006-15-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-15-003](VB条件分岐_原文根拠.md#F1006-15-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-15-004](VB条件分岐_原文根拠.md#F1006-15-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-15-005](VB条件分岐_原文根拠.md#F1006-15-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-15-006](VB条件分岐_原文根拠.md#F1006-15-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-001](VB条件分岐_原文根拠.md#F1006-16-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-002](VB条件分岐_原文根拠.md#F1006-16-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-003](VB条件分岐_原文根拠.md#F1006-16-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-004](VB条件分岐_原文根拠.md#F1006-16-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-005](VB条件分岐_原文根拠.md#F1006-16-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-006](VB条件分岐_原文根拠.md#F1006-16-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-007](VB条件分岐_原文根拠.md#F1006-16-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-16-008](VB条件分岐_原文根拠.md#F1006-16-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-001](VB条件分岐_原文根拠.md#F1006-17-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-002](VB条件分岐_原文根拠.md#F1006-17-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-003](VB条件分岐_原文根拠.md#F1006-17-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-004](VB条件分岐_原文根拠.md#F1006-17-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-005](VB条件分岐_原文根拠.md#F1006-17-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-006](VB条件分岐_原文根拠.md#F1006-17-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-007](VB条件分岐_原文根拠.md#F1006-17-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-008](VB条件分岐_原文根拠.md#F1006-17-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-009](VB条件分岐_原文根拠.md#F1006-17-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-010](VB条件分岐_原文根拠.md#F1006-17-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-011](VB条件分岐_原文根拠.md#F1006-17-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-012](VB条件分岐_原文根拠.md#F1006-17-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-013](VB条件分岐_原文根拠.md#F1006-17-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-014](VB条件分岐_原文根拠.md#F1006-17-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-015](VB条件分岐_原文根拠.md#F1006-17-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-016](VB条件分岐_原文根拠.md#F1006-17-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-017](VB条件分岐_原文根拠.md#F1006-17-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-018](VB条件分岐_原文根拠.md#F1006-17-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-019](VB条件分岐_原文根拠.md#F1006-17-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-020](VB条件分岐_原文根拠.md#F1006-17-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-021](VB条件分岐_原文根拠.md#F1006-17-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-022](VB条件分岐_原文根拠.md#F1006-17-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1006-17-023](VB条件分岐_原文根拠.md#F1006-17-023) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1009-06-001](VB条件分岐_原文根拠.md#F1009-06-001) | VB1009-01 | 業務項目に対応 | 2 |
| [F1009-06-002](VB条件分岐_原文根拠.md#F1009-06-002) | VB1009-02 | 業務項目に対応 | 6 |
| [F1009-06-003](VB条件分岐_原文根拠.md#F1009-06-003) | VB1009-04, VB1009-05 | 業務項目に対応 | 14 |
| [F1010-05-001](VB条件分岐_原文根拠.md#F1010-05-001) | VB1010-01 | 業務項目に対応 | 2 |
| [F1010-05-002](VB条件分岐_原文根拠.md#F1010-05-002) | VB1010-02 | 業務項目に対応 | 6 |
| [F1010-05-003](VB条件分岐_原文根拠.md#F1010-05-003) | VB1010-04, VB1010-05 | 業務項目に対応 | 14 |
| [F1011-06-001](VB条件分岐_原文根拠.md#F1011-06-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F1011-07-001](VB条件分岐_原文根拠.md#F1011-07-001) | VB1011-07 | 業務項目に対応 | 0 |
| [F1011-07-002](VB条件分岐_原文根拠.md#F1011-07-002) | VB1011-07, VB1011-11 | 業務項目に対応 | 5 |
| [F1011-07-003](VB条件分岐_原文根拠.md#F1011-07-003) | VB1011-07 | 業務項目に対応 | 8 |
| [F1011-09-001](VB条件分岐_原文根拠.md#F1011-09-001) | VB1011-01 | 業務項目に対応 | 3 |
| [F1011-09-002](VB条件分岐_原文根拠.md#F1011-09-002) | VB1011-02 | 業務項目に対応 | 0 |
| [F1011-09-003](VB条件分岐_原文根拠.md#F1011-09-003) | VB1011-02 | 業務項目に対応 | 11 |
| [F1011-09-004](VB条件分岐_原文根拠.md#F1011-09-004) | VB1011-02 | 業務項目に対応 | 2 |
| [F1011-09-005](VB条件分岐_原文根拠.md#F1011-09-005) | VB1011-03 | 業務項目に対応 | 3 |
| [F1011-09-006](VB条件分岐_原文根拠.md#F1011-09-006) | VB1011-03 | 業務項目に対応 | 0 |
| [F1011-09-007](VB条件分岐_原文根拠.md#F1011-09-007) | VB1011-04, VB1011-05, VB1011-06, VB1011-08, VB1011-11 | 業務項目に対応 | 19 |
| [F1011-09-008](VB条件分岐_原文根拠.md#F1011-09-008) | VB1011-08 | 業務項目に対応 | 2 |
| [F1011-09-009](VB条件分岐_原文根拠.md#F1011-09-009) | VB1011-08, VB1011-09 | 業務項目に対応 | 3 |
| [F1011-09-010](VB条件分岐_原文根拠.md#F1011-09-010) | VB1011-10 | 業務項目に対応 | 9 |
| [F1011-09-011](VB条件分岐_原文根拠.md#F1011-09-011) | VB1011-03 | 業務項目に対応 | 5 |
| [F1011-09-012](VB条件分岐_原文根拠.md#F1011-09-012) | VB1011-01 | 業務項目に対応 | 0 |
| [F2006-02-001](VB条件分岐_原文根拠.md#F2006-02-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-02-002](VB条件分岐_原文根拠.md#F2006-02-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-02-003](VB条件分岐_原文根拠.md#F2006-02-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-02-004](VB条件分岐_原文根拠.md#F2006-02-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-001](VB条件分岐_原文根拠.md#F2006-03-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-002](VB条件分岐_原文根拠.md#F2006-03-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-003](VB条件分岐_原文根拠.md#F2006-03-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-004](VB条件分岐_原文根拠.md#F2006-03-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-005](VB条件分岐_原文根拠.md#F2006-03-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-006](VB条件分岐_原文根拠.md#F2006-03-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-007](VB条件分岐_原文根拠.md#F2006-03-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-008](VB条件分岐_原文根拠.md#F2006-03-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-009](VB条件分岐_原文根拠.md#F2006-03-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-010](VB条件分岐_原文根拠.md#F2006-03-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-011](VB条件分岐_原文根拠.md#F2006-03-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-012](VB条件分岐_原文根拠.md#F2006-03-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-013](VB条件分岐_原文根拠.md#F2006-03-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-03-014](VB条件分岐_原文根拠.md#F2006-03-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-001](VB条件分岐_原文根拠.md#F2006-04-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-002](VB条件分岐_原文根拠.md#F2006-04-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-003](VB条件分岐_原文根拠.md#F2006-04-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-004](VB条件分岐_原文根拠.md#F2006-04-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-005](VB条件分岐_原文根拠.md#F2006-04-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-006](VB条件分岐_原文根拠.md#F2006-04-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-007](VB条件分岐_原文根拠.md#F2006-04-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-008](VB条件分岐_原文根拠.md#F2006-04-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-009](VB条件分岐_原文根拠.md#F2006-04-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-010](VB条件分岐_原文根拠.md#F2006-04-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-011](VB条件分岐_原文根拠.md#F2006-04-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-012](VB条件分岐_原文根拠.md#F2006-04-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-013](VB条件分岐_原文根拠.md#F2006-04-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-014](VB条件分岐_原文根拠.md#F2006-04-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-04-015](VB条件分岐_原文根拠.md#F2006-04-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-001](VB条件分岐_原文根拠.md#F2006-05-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-002](VB条件分岐_原文根拠.md#F2006-05-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-003](VB条件分岐_原文根拠.md#F2006-05-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-004](VB条件分岐_原文根拠.md#F2006-05-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-005](VB条件分岐_原文根拠.md#F2006-05-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-006](VB条件分岐_原文根拠.md#F2006-05-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-007](VB条件分岐_原文根拠.md#F2006-05-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-008](VB条件分岐_原文根拠.md#F2006-05-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-009](VB条件分岐_原文根拠.md#F2006-05-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-010](VB条件分岐_原文根拠.md#F2006-05-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-011](VB条件分岐_原文根拠.md#F2006-05-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-012](VB条件分岐_原文根拠.md#F2006-05-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-013](VB条件分岐_原文根拠.md#F2006-05-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-014](VB条件分岐_原文根拠.md#F2006-05-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-015](VB条件分岐_原文根拠.md#F2006-05-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-016](VB条件分岐_原文根拠.md#F2006-05-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-017](VB条件分岐_原文根拠.md#F2006-05-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-018](VB条件分岐_原文根拠.md#F2006-05-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-019](VB条件分岐_原文根拠.md#F2006-05-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-020](VB条件分岐_原文根拠.md#F2006-05-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-021](VB条件分岐_原文根拠.md#F2006-05-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-022](VB条件分岐_原文根拠.md#F2006-05-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-023](VB条件分岐_原文根拠.md#F2006-05-023) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-05-024](VB条件分岐_原文根拠.md#F2006-05-024) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-001](VB条件分岐_原文根拠.md#F2006-06-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-002](VB条件分岐_原文根拠.md#F2006-06-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-003](VB条件分岐_原文根拠.md#F2006-06-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-004](VB条件分岐_原文根拠.md#F2006-06-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-005](VB条件分岐_原文根拠.md#F2006-06-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-006](VB条件分岐_原文根拠.md#F2006-06-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-007](VB条件分岐_原文根拠.md#F2006-06-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-008](VB条件分岐_原文根拠.md#F2006-06-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-009](VB条件分岐_原文根拠.md#F2006-06-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-010](VB条件分岐_原文根拠.md#F2006-06-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-011](VB条件分岐_原文根拠.md#F2006-06-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-012](VB条件分岐_原文根拠.md#F2006-06-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-013](VB条件分岐_原文根拠.md#F2006-06-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-014](VB条件分岐_原文根拠.md#F2006-06-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-015](VB条件分岐_原文根拠.md#F2006-06-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-016](VB条件分岐_原文根拠.md#F2006-06-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-017](VB条件分岐_原文根拠.md#F2006-06-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-018](VB条件分岐_原文根拠.md#F2006-06-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-019](VB条件分岐_原文根拠.md#F2006-06-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-020](VB条件分岐_原文根拠.md#F2006-06-020) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-021](VB条件分岐_原文根拠.md#F2006-06-021) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-022](VB条件分岐_原文根拠.md#F2006-06-022) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-023](VB条件分岐_原文根拠.md#F2006-06-023) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-024](VB条件分岐_原文根拠.md#F2006-06-024) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-025](VB条件分岐_原文根拠.md#F2006-06-025) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-026](VB条件分岐_原文根拠.md#F2006-06-026) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-027](VB条件分岐_原文根拠.md#F2006-06-027) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-06-028](VB条件分岐_原文根拠.md#F2006-06-028) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-001](VB条件分岐_原文根拠.md#F2006-08-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-002](VB条件分岐_原文根拠.md#F2006-08-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-003](VB条件分岐_原文根拠.md#F2006-08-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-004](VB条件分岐_原文根拠.md#F2006-08-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-005](VB条件分岐_原文根拠.md#F2006-08-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-006](VB条件分岐_原文根拠.md#F2006-08-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-007](VB条件分岐_原文根拠.md#F2006-08-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-008](VB条件分岐_原文根拠.md#F2006-08-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-009](VB条件分岐_原文根拠.md#F2006-08-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-010](VB条件分岐_原文根拠.md#F2006-08-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-011](VB条件分岐_原文根拠.md#F2006-08-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-012](VB条件分岐_原文根拠.md#F2006-08-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-013](VB条件分岐_原文根拠.md#F2006-08-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-014](VB条件分岐_原文根拠.md#F2006-08-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-015](VB条件分岐_原文根拠.md#F2006-08-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-016](VB条件分岐_原文根拠.md#F2006-08-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-017](VB条件分岐_原文根拠.md#F2006-08-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-018](VB条件分岐_原文根拠.md#F2006-08-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-08-019](VB条件分岐_原文根拠.md#F2006-08-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-10-001](VB条件分岐_原文根拠.md#F2006-10-001) | VB2006-03, VB2006-04, VB2006-05 | 業務項目に対応 | 16 |
| [F2006-10-002](VB条件分岐_原文根拠.md#F2006-10-002) | VB2006-01 | 業務項目に対応 | 3 |
| [F2006-10-003](VB条件分岐_原文根拠.md#F2006-10-003) | VB2006-01 | 業務項目に対応 | 0 |
| [F2006-10-004](VB条件分岐_原文根拠.md#F2006-10-004) | VB2006-01 | 業務項目に対応 | 15 |
| [F2006-10-005](VB条件分岐_原文根拠.md#F2006-10-005) | VB2006-07, VB2006-08, VB2006-09 | 業務項目に対応 | 15 |
| [F2006-10-006](VB条件分岐_原文根拠.md#F2006-10-006) | VB2006-07 | 業務項目に対応 | 3 |
| [F2006-10-007](VB条件分岐_原文根拠.md#F2006-10-007) | VB2006-12 | 業務項目に対応 | 1 |
| [F2006-10-008](VB条件分岐_原文根拠.md#F2006-10-008) | VB2006-06 | 業務項目に対応 | 0 |
| [F2006-10-009](VB条件分岐_原文根拠.md#F2006-10-009) | VB2006-12 | 業務項目に対応 | 5 |
| [F2006-10-010](VB条件分岐_原文根拠.md#F2006-10-010) | VB2006-12 | 業務項目に対応 | 5 |
| [F2006-10-011](VB条件分岐_原文根拠.md#F2006-10-011) | VB2006-10 | 業務項目に対応 | 0 |
| [F2006-10-012](VB条件分岐_原文根拠.md#F2006-10-012) | VB2006-12 | 業務項目に対応 | 2 |
| [F2006-10-013](VB条件分岐_原文根拠.md#F2006-10-013) | VB2006-12 | 業務項目に対応 | 2 |
| [F2006-10-014](VB条件分岐_原文根拠.md#F2006-10-014) | VB2006-12 | 業務項目に対応 | 2 |
| [F2006-10-015](VB条件分岐_原文根拠.md#F2006-10-015) | VB2006-13 | 業務項目に対応 | 0 |
| [F2006-10-016](VB条件分岐_原文根拠.md#F2006-10-016) | VB2006-13 | 業務項目に対応 | 0 |
| [F2006-12-001](VB条件分岐_原文根拠.md#F2006-12-001) | VB2006-14 | 業務項目に対応 | 2 |
| [F2006-12-002](VB条件分岐_原文根拠.md#F2006-12-002) | VB2006-02 | 業務項目に対応 | 5 |
| [F2006-12-003](VB条件分岐_原文根拠.md#F2006-12-003) | VB2006-14, VB2006-18 | 業務項目に対応 | 9 |
| [F2006-12-004](VB条件分岐_原文根拠.md#F2006-12-004) | VB2006-15 | 業務項目に対応 | 11 |
| [F2006-12-005](VB条件分岐_原文根拠.md#F2006-12-005) | VB2006-15 | 業務項目に対応 | 9 |
| [F2006-12-006](VB条件分岐_原文根拠.md#F2006-12-006) | VB2006-16 | 業務項目に対応 | 7 |
| [F2006-12-007](VB条件分岐_原文根拠.md#F2006-12-007) | VB2006-14 | 業務項目に対応 | 1 |
| [F2006-12-008](VB条件分岐_原文根拠.md#F2006-12-008) | VB2006-03 | 業務項目に対応 | 4 |
| [F2006-12-009](VB条件分岐_原文根拠.md#F2006-12-009) | VB2006-04 | 業務項目に対応 | 0 |
| [F2006-12-010](VB条件分岐_原文根拠.md#F2006-12-010) | VB2006-05 | 業務項目に対応 | 12 |
| [F2006-12-011](VB条件分岐_原文根拠.md#F2006-12-011) | VB2006-09, VB2006-11 | 業務項目に対応 | 20 |
| [F2006-12-012](VB条件分岐_原文根拠.md#F2006-12-012) | VB2006-06 | 業務項目に対応 | 24 |
| [F2006-12-013](VB条件分岐_原文根拠.md#F2006-12-013) | VB2006-12 | 業務項目に対応 | 3 |
| [F2006-12-014](VB条件分岐_原文根拠.md#F2006-12-014) | VB2006-12 | 業務項目に対応 | 0 |
| [F2006-12-015](VB条件分岐_原文根拠.md#F2006-12-015) | VB2006-12 | 業務項目に対応 | 3 |
| [F2006-12-016](VB条件分岐_原文根拠.md#F2006-12-016) | VB2006-12 | 業務項目に対応 | 0 |
| [F2006-12-017](VB条件分岐_原文根拠.md#F2006-12-017) | VB2006-10 | 業務項目に対応 | 0 |
| [F2006-12-018](VB条件分岐_原文根拠.md#F2006-12-018) | VB2006-12 | 業務項目に対応 | 0 |
| [F2006-12-019](VB条件分岐_原文根拠.md#F2006-12-019) | VB2006-14 | 業務項目に対応 | 1 |
| [F2006-12-020](VB条件分岐_原文根拠.md#F2006-12-020) | VB2006-17 | 業務項目に対応 | 1 |
| [F2006-12-021](VB条件分岐_原文根拠.md#F2006-12-021) | VB2006-17 | 業務項目に対応 | 5 |
| [F2006-12-022](VB条件分岐_原文根拠.md#F2006-12-022) | VB2006-18 | 業務項目に対応 | 3 |
| [F2006-12-023](VB条件分岐_原文根拠.md#F2006-12-023) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-024](VB条件分岐_原文根拠.md#F2006-12-024) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-025](VB条件分岐_原文根拠.md#F2006-12-025) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-026](VB条件分岐_原文根拠.md#F2006-12-026) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-027](VB条件分岐_原文根拠.md#F2006-12-027) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-028](VB条件分岐_原文根拠.md#F2006-12-028) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-029](VB条件分岐_原文根拠.md#F2006-12-029) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-030](VB条件分岐_原文根拠.md#F2006-12-030) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-031](VB条件分岐_原文根拠.md#F2006-12-031) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-032](VB条件分岐_原文根拠.md#F2006-12-032) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-033](VB条件分岐_原文根拠.md#F2006-12-033) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-034](VB条件分岐_原文根拠.md#F2006-12-034) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2006-12-035](VB条件分岐_原文根拠.md#F2006-12-035) | VB2006-12 | 業務項目に対応 | 3 |
| [F2006-12-036](VB条件分岐_原文根拠.md#F2006-12-036) | VB2006-13 | 業務項目に対応 | 0 |
| [F2006-12-037](VB条件分岐_原文根拠.md#F2006-12-037) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-07-001](VB条件分岐_原文根拠.md#F2007-07-001) | VB2007-05 | 業務項目に対応 | 15 |
| [F2007-07-002](VB条件分岐_原文根拠.md#F2007-07-002) | VB2007-05 | 業務項目に対応 | 3 |
| [F2007-07-003](VB条件分岐_原文根拠.md#F2007-07-003) | VB2007-05 | 業務項目に対応 | 3 |
| [F2007-07-004](VB条件分岐_原文根拠.md#F2007-07-004) | VB2007-05 | 業務項目に対応 | 3 |
| [F2007-07-005](VB条件分岐_原文根拠.md#F2007-07-005) | VB2007-06, VB2007-08 | 業務項目に対応 | 19 |
| [F2007-07-006](VB条件分岐_原文根拠.md#F2007-07-006) | VB2007-08 | 業務項目に対応 | 1 |
| [F2007-07-007](VB条件分岐_原文根拠.md#F2007-07-007) | VB2007-07 | 業務項目に対応 | 7 |
| [F2007-07-008](VB条件分岐_原文根拠.md#F2007-07-008) | VB2007-07 | 業務項目に対応 | 5 |
| [F2007-07-009](VB条件分岐_原文根拠.md#F2007-07-009) | VB2007-07 | 業務項目に対応 | 6 |
| [F2007-07-010](VB条件分岐_原文根拠.md#F2007-07-010) | VB2007-07 | 業務項目に対応 | 11 |
| [F2007-07-011](VB条件分岐_原文根拠.md#F2007-07-011) | VB2007-07 | 業務項目に対応 | 1 |
| [F2007-07-012](VB条件分岐_原文根拠.md#F2007-07-012) | VB2007-07 | 業務項目に対応 | 1 |
| [F2007-07-013](VB条件分岐_原文根拠.md#F2007-07-013) | VB2007-07 | 業務項目に対応 | 1 |
| [F2007-09-001](VB条件分岐_原文根拠.md#F2007-09-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-09-002](VB条件分岐_原文根拠.md#F2007-09-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-10-001](VB条件分岐_原文根拠.md#F2007-10-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-10-002](VB条件分岐_原文根拠.md#F2007-10-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-10-003](VB条件分岐_原文根拠.md#F2007-10-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-10-004](VB条件分岐_原文根拠.md#F2007-10-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-10-005](VB条件分岐_原文根拠.md#F2007-10-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-11-001](VB条件分岐_原文根拠.md#F2007-11-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-11-002](VB条件分岐_原文根拠.md#F2007-11-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-11-003](VB条件分岐_原文根拠.md#F2007-11-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-11-004](VB条件分岐_原文根拠.md#F2007-11-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-11-005](VB条件分岐_原文根拠.md#F2007-11-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-001](VB条件分岐_原文根拠.md#F2007-12-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-002](VB条件分岐_原文根拠.md#F2007-12-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-003](VB条件分岐_原文根拠.md#F2007-12-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-004](VB条件分岐_原文根拠.md#F2007-12-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-005](VB条件分岐_原文根拠.md#F2007-12-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-006](VB条件分岐_原文根拠.md#F2007-12-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-007](VB条件分岐_原文根拠.md#F2007-12-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-008](VB条件分岐_原文根拠.md#F2007-12-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-12-009](VB条件分岐_原文根拠.md#F2007-12-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-001](VB条件分岐_原文根拠.md#F2007-13-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-002](VB条件分岐_原文根拠.md#F2007-13-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-003](VB条件分岐_原文根拠.md#F2007-13-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-004](VB条件分岐_原文根拠.md#F2007-13-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-005](VB条件分岐_原文根拠.md#F2007-13-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-006](VB条件分岐_原文根拠.md#F2007-13-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-13-007](VB条件分岐_原文根拠.md#F2007-13-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-001](VB条件分岐_原文根拠.md#F2007-15-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-002](VB条件分岐_原文根拠.md#F2007-15-002) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-003](VB条件分岐_原文根拠.md#F2007-15-003) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-004](VB条件分岐_原文根拠.md#F2007-15-004) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-005](VB条件分岐_原文根拠.md#F2007-15-005) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-006](VB条件分岐_原文根拠.md#F2007-15-006) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-007](VB条件分岐_原文根拠.md#F2007-15-007) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-008](VB条件分岐_原文根拠.md#F2007-15-008) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-009](VB条件分岐_原文根拠.md#F2007-15-009) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-010](VB条件分岐_原文根拠.md#F2007-15-010) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-011](VB条件分岐_原文根拠.md#F2007-15-011) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-012](VB条件分岐_原文根拠.md#F2007-15-012) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-013](VB条件分岐_原文根拠.md#F2007-15-013) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-014](VB条件分岐_原文根拠.md#F2007-15-014) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-015](VB条件分岐_原文根拠.md#F2007-15-015) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-016](VB条件分岐_原文根拠.md#F2007-15-016) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-017](VB条件分岐_原文根拠.md#F2007-15-017) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-018](VB条件分岐_原文根拠.md#F2007-15-018) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-15-019](VB条件分岐_原文根拠.md#F2007-15-019) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F2007-17-001](VB条件分岐_原文根拠.md#F2007-17-001) | VB2007-01 | 業務項目に対応 | 12 |
| [F2007-17-002](VB条件分岐_原文根拠.md#F2007-17-002) | VB2007-02, VB2007-03 | 業務項目に対応 | 8 |
| [F2007-17-003](VB条件分岐_原文根拠.md#F2007-17-003) | VB2007-04 | 業務項目に対応 | 2 |
| [F2007-17-004](VB条件分岐_原文根拠.md#F2007-17-004) | VB2007-04 | 業務項目に対応 | 3 |
| [F2007-17-005](VB条件分岐_原文根拠.md#F2007-17-005) | VB2007-04 | 業務項目に対応 | 1 |
| [F2007-17-006](VB条件分岐_原文根拠.md#F2007-17-006) | VB2007-04 | 業務項目に対応 | 1 |
| [F2007-17-007](VB条件分岐_原文根拠.md#F2007-17-007) | VB2007-04 | 業務項目に対応 | 1 |
| [F3002-06-001](VB条件分岐_原文根拠.md#F3002-06-001) | VB3002-01 | 業務項目に対応 | 0 |
| [F3002-06-002](VB条件分岐_原文根拠.md#F3002-06-002) | VB3002-02, VB3002-04 | 業務項目に対応 | 7 |
| [F3003-05-001](VB条件分岐_原文根拠.md#F3003-05-001) | VB3003-01 | 業務項目に対応 | 0 |
| [F3003-05-002](VB条件分岐_原文根拠.md#F3003-05-002) | VB3003-05 | 業務項目に対応 | 2 |
| [F3003-05-003](VB条件分岐_原文根拠.md#F3003-05-003) | VB3003-05 | 業務項目に対応 | 2 |
| [F3003-05-004](VB条件分岐_原文根拠.md#F3003-05-004) | VB3003-02, VB3003-03 | 業務項目に対応 | 10 |
| [F3003-05-005](VB条件分岐_原文根拠.md#F3003-05-005) | VB3003-03 | 業務項目に対応 | 5 |
| [F3003-05-006](VB条件分岐_原文根拠.md#F3003-05-006) | VB3003-01 | 業務項目に対応 | 2 |
| [F3004-15-001](VB条件分岐_原文根拠.md#F3004-15-001) | VB3004-01 | 業務項目に対応 | 1 |
| [F3004-15-002](VB条件分岐_原文根拠.md#F3004-15-002) | VB3004-02 | 業務項目に対応 | 1 |
| [F3004-15-003](VB条件分岐_原文根拠.md#F3004-15-003) | VB3004-03 | 業務項目に対応 | 5 |
| [F3004-15-004](VB条件分岐_原文根拠.md#F3004-15-004) | VB3004-04 | 業務項目に対応 | 1 |
| [F3004-15-005](VB条件分岐_原文根拠.md#F3004-15-005) | VB3004-05 | 業務項目に対応 | 1 |
| [F3004-15-006](VB条件分岐_原文根拠.md#F3004-15-006) | VB3004-09 | 業務項目に対応 | 3 |
| [F3004-15-007](VB条件分岐_原文根拠.md#F3004-15-007) | VB3004-06 | 業務項目に対応 | 1 |
| [F3004-15-008](VB条件分岐_原文根拠.md#F3004-15-008) | VB3004-07 | 業務項目に対応 | 2 |
| [F3004-15-009](VB条件分岐_原文根拠.md#F3004-15-009) | VB3004-08 | 業務項目に対応 | 1 |
| [F3004-15-010](VB条件分岐_原文根拠.md#F3004-15-010) | VB3004-10 | 業務項目に対応 | 1 |
| [F3004-15-011](VB条件分岐_原文根拠.md#F3004-15-011) | VB3004-11 | 業務項目に対応 | 2 |
| [F3004-15-012](VB条件分岐_原文根拠.md#F3004-15-012) | VB3004-11 | 業務項目に対応 | 3 |
| [F3004-15-013](VB条件分岐_原文根拠.md#F3004-15-013) | VB3004-12 | 業務項目に対応 | 1 |
| [F3004-15-014](VB条件分岐_原文根拠.md#F3004-15-014) | VB3004-13 | 業務項目に対応 | 1 |
| [F3004-15-015](VB条件分岐_原文根拠.md#F3004-15-015) | VB3004-14 | 業務項目に対応 | 1 |
| [F3004-15-016](VB条件分岐_原文根拠.md#F3004-15-016) | VB3004-15 | 業務項目に対応 | 1 |
| [F3004-15-017](VB条件分岐_原文根拠.md#F3004-15-017) | VB3004-16 | 業務項目に対応 | 1 |
| [F3004-15-018](VB条件分岐_原文根拠.md#F3004-15-018) | VB3004-17 | 業務項目に対応 | 1 |
| [F3004-15-019](VB条件分岐_原文根拠.md#F3004-15-019) | VB3004-18 | 業務項目に対応 | 1 |
| [F3004-15-020](VB条件分岐_原文根拠.md#F3004-15-020) | VB3004-19 | 業務項目に対応 | 1 |
| [F3004-15-021](VB条件分岐_原文根拠.md#F3004-15-021) | VB3004-20 | 業務項目に対応 | 1 |
| [F3004-15-022](VB条件分岐_原文根拠.md#F3004-15-022) | VB3004-21 | 業務項目に対応 | 3 |
| [F3004-15-023](VB条件分岐_原文根拠.md#F3004-15-023) | VB3004-21 | 業務項目に対応 | 0 |
| [F3004-15-024](VB条件分岐_原文根拠.md#F3004-15-024) | VB3004-21 | 業務項目に対応 | 2 |
| [F3005-04-001](VB条件分岐_原文根拠.md#F3005-04-001) | ― | 条件分岐のない受渡し・保持プロパティ（原文索引） | 0 |
| [F3005-06-001](VB条件分岐_原文根拠.md#F3005-06-001) | VB3005-01 | 業務項目に対応 | 2 |
| [F3005-06-002](VB条件分岐_原文根拠.md#F3005-06-002) | VB3005-01 | 業務項目に対応 | 3 |
| [F3005-06-003](VB条件分岐_原文根拠.md#F3005-06-003) | VB3005-01 | 業務項目に対応 | 0 |
| [F3005-06-004](VB条件分岐_原文根拠.md#F3005-06-004) | VB3005-02, VB3005-03, VB3005-04 | 業務項目に対応 | 6 |
| [F3005-06-005](VB条件分岐_原文根拠.md#F3005-06-005) | VB3005-02 | 業務項目に対応 | 0 |
| [F3005-06-006](VB条件分岐_原文根拠.md#F3005-06-006) | VB3005-04 | 業務項目に対応 | 1 |
| [F3005-06-007](VB条件分岐_原文根拠.md#F3005-06-007) | VB3005-05 | 業務項目に対応 | 1 |
| [F3005-06-008](VB条件分岐_原文根拠.md#F3005-06-008) | VB3005-06 | 業務項目に対応 | 4 |
| [F3005-06-009](VB条件分岐_原文根拠.md#F3005-06-009) | VB3005-07 | 業務項目に対応 | 5 |
| [F3007-08-001](VB条件分岐_原文根拠.md#F3007-08-001) | VB3007-01 | 業務項目に対応 | 4 |
| [F3007-08-002](VB条件分岐_原文根拠.md#F3007-08-002) | VB3007-02 | 業務項目に対応 | 10 |
| [F3007-08-003](VB条件分岐_原文根拠.md#F3007-08-003) | VB3007-04 | 業務項目に対応 | 8 |
| [F3007-08-004](VB条件分岐_原文根拠.md#F3007-08-004) | VB3007-05, VB3007-06 | 業務項目に対応 | 9 |
| [F3007-08-005](VB条件分岐_原文根拠.md#F3007-08-005) | VB3007-07, VB3007-08 | 業務項目に対応 | 14 |
| [F3008-06-001](VB条件分岐_原文根拠.md#F3008-06-001) | VB3008-01 | 業務項目に対応 | 0 |
| [F3008-06-002](VB条件分岐_原文根拠.md#F3008-06-002) | VB3008-01 | 業務項目に対応 | 0 |
| [F3008-06-003](VB条件分岐_原文根拠.md#F3008-06-003) | VB3008-01 | 業務項目に対応 | 2 |
| [F3008-06-004](VB条件分岐_原文根拠.md#F3008-06-004) | VB3008-03, VB3008-04, VB3008-05 | 業務項目に対応 | 6 |

## 入力キーと共通定義

以下は検査キーの候補です。明示ValidateKeyを優先し、未指定はバインド項目を使用。リンク・チェック列など検査対象外の型、ReadOnly・Immediate・非活性・非表示はC08の実行条件を併せて判定します。空の設定を既定Falseと推測せず、定義なしは「未指定」と記載。型が対象外の候補も含むので行数は必須項目数ではありません。

共通定義：[CPValidate.xml](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CPValidate.xml>)。Designerの実行時上書きは元の原文索引へ。

| 画面 | コントロール | 検査キー候補 | 必須／即時通過／読取専用（明示値） | 共通型・長さ・範囲・マスク | 原本 |
| --- | --- | --- | --- | --- | --- |
| 1005 | IRAIJIGYOSYO\_NKTextBox | IRAIJIGYOSYO\_NK | False／False／True | name=依頼事業所名; type=String; length=200; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | IRAITANTOSYATextBox | IRAITANTOSYA | False／False／True | name=依頼担当者; type=String; length=50; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | KEIYAKU\_NOTextBox | KEIYAKU\_NO | False／False／未指定 | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | JIDOKOSIN\_FLGCheckBox | JIDOKOSIN\_FLG | 未指定／未指定／未指定 | name=自動更新フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | JIKAIKOSIN\_DTTextBox | JIKAIKOSIN\_DT | False／False／True | name=次回更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | KEIYAKU\_DTTextBox | KEIYAKU\_DT | False／False／未指定 | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | KEIYAKUMANRYO\_DTTextBox | KEIYAKUMANRYO\_DT | False／False／未指定 | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | ENTYOKEIYAKUMANRYO\_DTTextBox | ENTYOKEIYAKUMANRYO\_DT | False／False／未指定 | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | KAIYAKU\_DTTextBox | KAIYAKU\_DT | False／False／未指定 | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | PACK\_FLGCheckBox | PACK\_FLG | 未指定／未指定／未指定 | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | PACKKEIYAKUNAIYOCPTextBox | PACKKEIYAKUNAIYO | False／False／未指定 | name=契約内容; type=String; length=400; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | BIKOTextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | CREATED\_DTCPTextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | CREATED\_BYCPTextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | LASTUPDATE\_DTCPTextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | LASTUPDATE\_BYCPTextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | SHOKAI\_KEIYAKU\_DTCPTextBox | SHOKAI\_KEIYAKU\_DT | False／False／未指定 | name=初回契約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | KOSINNAIYOTextBox | KOSINNAIYO | False／False／未指定 | name=更新内容; type=String; length=4000; min=; max=; mask= | [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) |
| 1005 | ShiharaiBikoTextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SIHARAIKAISUTextBox | KAISU | False／False／未指定 | name=回数; type=Decimal; length=2; min=0; max=12; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKOSEI\_NK\_TENKEN\_TextBox | KIKIKOSEI\_NK | False／False／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIHINMEI\_NK\_TENKEN\_TextBox | KIKIHINMEI\_NK | False／False／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKATASHIKI\_TENKEN\_TextBox | KIKIKATASHIKI | False／False／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TENKENKAISU\_TENKEN\_TextBox | TENKENKAISU | False／False／未指定 | name=点検回数; type=Decimal; length=2; min=0; max=12; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI1\_TENKEN\_CheckBox | TSUKI1 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI2\_TENKEN\_CheckBox | TSUKI2 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI3\_TENKEN\_CheckBox | TSUKI3 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI4\_TENKEN\_CheckBox | TSUKI4 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI5\_TENKEN\_CheckBox | TSUKI5 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI6\_TENKEN\_CheckBox | TSUKI6 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI7\_TENKEN\_CheckBox | TSUKI7 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI8\_TENKEN\_CheckBox | TSUKI8 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI9\_TENKEN\_CheckBox | TSUKI9 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI10\_TENKEN\_CheckBox | TSUKI10 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI11\_TENKEN\_CheckBox | TSUKI11 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI12\_TENKEN\_CheckBox | TSUKI12 | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TENKENKANOYOBI\_TENKEN\_ComboBox | TENKENKANOYOBI | False／False／未指定 | name=点検可能曜日; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | YAKANTAIOUMU\_TENKEN\_CheckBox | YAKANTAIOUMU | 未指定／False／未指定 | name=夜間対応有無; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | BIKO\_TENKEN\_TextBox | BIKO | 未指定／未指定／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CREATED\_DT\_TENKEN\_TextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CREATED\_BY\_TENKEN\_TextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LASTUPDATE\_DT\_TENKEN\_TextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LASTUPDATE\_BY\_TENKEN\_TextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TK\_TENKEN\_ID\_TENKEN\_TextBox | TK\_TENKEN\_ID | False／False／True | name=取引先契約点検ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKOSEI\_ID\_TENKEN\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIMEISAI\_ID\_TENKEN\_TextBox | KIKIMEISAI\_ID | False／True／True | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | OYAKIKIBUNRUI\_CD\_TENKEN\_TextBox | OYAKIKIBUNRUI\_CD | False／True／True | name=親機器分類コード; type=String; length=20; min=; max=; mask=HANKAKU | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAE\_HYOJIJUN | MAE\_HYOJIJUN | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAF\_HYOJIJUN | MAF\_HYOJIJUN | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KAISI\_DTTextBox | KAISI\_DT | True／False／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SYURYO\_DTTextBox | SYURYO\_DT | True／False／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKOSEI\_NK\_KIKIJOHO\_TextBox | KIKIKOSEI\_NK | False／False／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SEIZOMAKERNK\_KIKIJOHO\_TextBox | SEIZOMAKER\_NK | False／False／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIHINMEINK\_KIKIJOHO\_TextBox | KIKIHINMEI\_NK | False／False／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKATASHIKI\_KIKIJOHO\_TextBox | KIKIKATASHIKI | False／False／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SURYONM\_KIKIJOHO\_TextBox | SURYO\_NM | False／False／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HYOJUNKIN\_KIKIJOHO\_TextBox | HYOJUN\_KIN | False／False／True | name=標準価格; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HYOJUNKEI\_KIN\_KIKIJOHO\_TextBox | HYOJUNKEI\_KIN | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SIKIRIKIN\_KIKIJOHO\_TextBox | SIKIRI\_KIN | False／False／True | name=仕切り価格; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SIKIRIKEI\_KIN\_KIKIJOHO\_TextBox | SIKIRIKEI\_KIN | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TEHAISEIBAN\_KIKIJOHO\_TextBox | TEHAISEIBAN | False／False／True | name=手配製番; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TORIHOSYUJIKAN\_ID\_KIKIJOHO\_ComboBox | TORIHOSYUJIKAN\_ID | False／False／True | name=取引先保守時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TENKENUMU\_KIKIJOHO\_ComboBox | TENKENUMU | False／False／True | name=点検有無; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HOSYUHOHO\_KIKIJOHO\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SERVICEKEITAI\_KIKIJOHO\_ComboBox | SERVICEKEITAI | False／False／True | name=サービス形態; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKB\_SUPPORT\_ID\_OLD\_KIKIJOHO\_TextBox | TKB\_SUPPORT\_ID\_OLD | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAB\_SUPPORT\_ID\_KIKIJOHO\_TextBox | MAB\_SUPPORT\_ID | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | BIKO\_KIKIJOHO\_TextBox | BIKO | False／False／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKKIKIKOSEIID\_KIKIJOHO\_TextBox | TK\_KIKIKOSEI\_ID | False／False／True | name=取引先契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKKIKANID\_KIKIJOHO\_TextBox | TK\_KIKAN\_ID | False／False／True | name=取引先契約期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKOSEIID\_KIKIJOHO\_TextBox | KIKIKOSEI\_ID | False／False／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TK\_KIKIMEISAI\_ID\_KIKIJOHO\_TextBox | TK\_KIKIMEISAI\_ID | False／False／True | name=取引先契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TK\_TANKA\_ID\_KIKIJOHO\_TextBox | TK\_TANKA\_ID | False／False／True | name=取引先契約単価ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAE\_HYOJIJUN\_KIKIJOHO\_TextBox | MAE\_HYOJIJUN | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAF\_HYOJIJUN\_KIKIJOHO\_TextBox | MAF\_HYOJIJUN | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KEIYAKUJIKANTAITextBox | KEIYAKUJIKANTAI | False／False／未指定 | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HOSYUHOHOCombobox | HOSYUHOHO | False／False／未指定 | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HYOJUNGOKEI\_KINCPTextBox | HYOJUNGOKEI\_KIN | False／False／True | name=標準価格合計; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SIKIRISYOKEI\_KINCPTextBox | SIKIRISYOKEI\_KIN | False／False／True | name=仕切り価格小計; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SYUSSEINEBIKI\_KINCPTextBox | SYUSSEINEBIKI\_KIN | False／False／True | name=出精値引き; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SIKIRIGOKEI\_KINCPTextBox | SIKIRIGOKEI\_KIN | False／False／True | name=仕切り価格合計; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | BIKOTextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TENPUFILE\_NK\_TENPU\_LinkBox | TENPUFILE\_NK | 未指定／未指定／True | name=添付ファイル名; type=String; length=2000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | SHONINJOTAI\_TENPU\_ComboBox | SHONINJOTAI | False／False／True | name=承認状態; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | DIRECTORY\_TENPU\_TextBox | DIRECTORY | False／False／未指定 | name=ディレクトリ; type=String; length=2000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LOCAL\_PATH\_TENPU\_TextBox | LOCAL\_PATH | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TK\_TENPU\_ID\_TENPU\_TextBox | TK\_TENPU\_ID | False／False／True | name=取引先契約添付ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TK\_KIKAN\_ID\_TENPU\_TextBox | TK\_KIKAN\_ID | 未指定／未指定／True | name=取引先契約期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CREATED\_DTCPTextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CREATED\_BYCPTextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LASTUPDATE\_DTCPTextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LASTUPDATE\_BYCPTextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CpDataGridTextBoxColumn1 | KIKIKOSEI\_NK | False／False／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CpDataGridTextBoxColumn2 | KIKIHINMEI\_NK | False／False／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CpDataGridTextBoxColumn3 | KIKIKATASHIKI | False／False／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HOSHUGAISHA\_ID\_TENKEN\_ComboBox | HOSHUGAISHA\_ID | False／False／未指定 | name=保守会社ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | YAKANHOSHUGAISHA\_ID\_TENKEN\_ComboBox | YAKANHOSHUGAISHA\_ID | False／False／未指定 | name=夜間保守会社ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENIDDataGridViewTextBoxColumn | TK\_TENKEN\_ID | 未指定／未指定／未指定 | name=取引先契約点検ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKKIKIKOSEIIDDataGridViewTextBoxColumn | TK\_KIKIKOSEI\_ID | 未指定／未指定／未指定 | name=取引先契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TENKENKAISUDataGridViewTextBoxColumn | TENKENKAISU | 未指定／未指定／未指定 | name=点検回数; type=Decimal; length=2; min=0; max=12; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TENKENKANOYOBIDataGridViewTextBoxColumn | TENKENKANOYOBI | 未指定／未指定／未指定 | name=点検可能曜日; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | YAKANTAIOUMUDataGridViewTextBoxColumn | YAKANTAIOUMU | 未指定／未指定／未指定 | name=夜間対応有無; type=String; length=1; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | HOSHUGAISHAIDDataGridViewTextBoxColumn | HOSHUGAISHA\_ID | 未指定／未指定／未指定 | name=保守会社ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | BIKODataGridViewTextBoxColumn | BIKO | 未指定／未指定／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CREATEDDTDataGridViewTextBoxColumn | CREATED\_DT | 未指定／未指定／未指定 | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | CREATEDBYDataGridViewTextBoxColumn | CREATED\_BY | 未指定／未指定／未指定 | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LASTUPDATEDTDataGridViewTextBoxColumn | LASTUPDATE\_DT | 未指定／未指定／未指定 | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | LASTUPDATEBYDataGridViewTextBoxColumn | LASTUPDATE\_BY | 未指定／未指定／未指定 | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID1DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID1 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID2DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID2 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID3DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID3 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID4DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID4 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID5DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID5 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID6DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID6 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID7DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID7 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID8DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID8 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID9DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID9 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID10DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID10 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID11DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID11 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKTENKENMEISAIID12DataGridViewTextBoxColumn | TK\_TENKENMEISAI\_ID12 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI1DataGridViewTextBoxColumn | TSUKI1 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI2DataGridViewTextBoxColumn | TSUKI2 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI3DataGridViewTextBoxColumn | TSUKI3 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI4DataGridViewTextBoxColumn | TSUKI4 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI5DataGridViewTextBoxColumn | TSUKI5 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI6DataGridViewTextBoxColumn | TSUKI6 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI7DataGridViewTextBoxColumn | TSUKI7 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI8DataGridViewTextBoxColumn | TSUKI8 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI9DataGridViewTextBoxColumn | TSUKI9 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI10DataGridViewTextBoxColumn | TSUKI10 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI11DataGridViewTextBoxColumn | TSUKI11 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TSUKI12DataGridViewTextBoxColumn | TSUKI12 | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKOSEINKDataGridViewTextBoxColumn | KIKIKOSEI\_NK | 未指定／未指定／未指定 | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKOSEIIDDataGridViewTextBoxColumn | KIKIKOSEI\_ID | 未指定／未指定／未指定 | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIHINMEINKDataGridViewTextBoxColumn | KIKIHINMEI\_NK | 未指定／未指定／未指定 | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIMEISAIIDDataGridViewTextBoxColumn | KIKIMEISAI\_ID | 未指定／未指定／未指定 | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KIKIKATASHIKIDataGridViewTextBoxColumn | KIKIKATASHIKI | 未指定／未指定／未指定 | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | YAKANHOSHUGAISHAIDDataGridViewTextBoxColumn | YAKANHOSHUGAISHA\_ID | 未指定／未指定／未指定 | name=夜間保守会社ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | TKKIKANIDDataGridViewTextBoxColumn | TK\_KIKAN\_ID | 未指定／未指定／未指定 | name=取引先契約期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | OYAKIKIBUNRUICDDataGridViewTextBoxColumn | OYAKIKIBUNRUI\_CD | 未指定／未指定／未指定 | name=親機器分類コード; type=String; length=20; min=; max=; mask=HANKAKU | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAEHYOJIJUNDataGridViewTextBoxColumn | MAE\_HYOJIJUN | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | MAFHYOJIJUNDataGridViewTextBoxColumn | MAF\_HYOJIJUN | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1005 | KEIYAKU\_NOTextBox | KEIYAKU\_NO | False／False／未指定 | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) |
| 1006 | KAISI\_DTTextBox | KAISI\_DT | True／False／未指定 | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SYURYO\_DTTextBox | SYURYO\_DT | True／False／未指定 | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KIKIKOSEI\_NK\_TANKA\_TextBox | KIKIKOSEI\_NK | False／False／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TEHAISEIBAN\_TANKA\_TextBox | TEHAISEIBAN | False／False／True | name=手配製番; type=String; length=50; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SEIZOMAKER\_NK\_TANKA\_TextBox | SEIZOMAKER\_NK | False／False／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KIKIHINMEI\_NK\_TANKA\_TextBox | KIKIHINMEI\_NK | False／False／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KIKIKATASHIKI\_TANKA\_TextBox | KIKIKATASHIKI | False／False／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SURYO\_NM\_TANKA\_TextBox | SURYO\_NM | False／False／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | HYOJUN\_KIN\_TANKA\_TextBox | HYOJUN\_KIN | False／False／True | name=標準価格; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SIKIRI\_KIN\_TANKA\_TextBox | SIKIRI\_KIN | False／False／True | name=仕切り価格; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | PACK\_FLG\_TANKA\_CheckBox | PACK\_FLG | 未指定／False／True | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KEIYAKUNAIYO\_TANKA\_TextBox | KEIYAKUNAIYO | False／False／True | name=契約内容; type=String; length=400; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KEIYAKU\_NO\_TANKA\_TextBox | KEIYAKU\_NO | False／False／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TORIHOSYUJIKAN\_ID\_TANKA\_ComboBox | TORIHOSYUJIKAN\_ID | False／True／True | name=取引先保守時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TENKENUMU\_TANKA\_ComboBox | TENKENUMU | False／True／True | name=点検有無; type=String; length=1; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | HOSYUHOHO\_TANKA\_ComboBox | HOSYUHOHO | False／True／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SERVICEKEITAI\_TANKA\_ComboBox | SERVICEKEITAI | False／True／True | name=サービス形態; type=String; length=1; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TMF\_BIKO\_TANKA\_TextBox | TMF\_BIKO | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_KIKAN\_ID\_TANKA\_TextBox | TM\_KIKAN\_ID | False／False／True | name=取引先見積期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_TANKA\_ID\_TANKA\_TextBox | TM\_TANKA\_ID | False／False／True | name=取引先見積単価ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | CHECK\_FLG\_MITSUMORI\_TextBox | CHECK\_FLG | 未指定／False／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_MITSUMORI\_NO\_MITSUMORI\_TextBox | TM\_IRAI\_NO | False／False／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KEIYAKUJIKANTAI\_MITSUMORI\_TextBox | KEIYAKUJIKANTAI | 未指定／未指定／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | HOSYUHOHO\_MITSUMORI\_ComboBox | HOSYUHOHO | False／True／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TENKENUMU\_MITSUMORI\_ComboBox | TENKENUMU | False／True／True | name=点検有無; type=String; length=1; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TENKENKANOYOBI\_MITSUMORI\_ComboBox | TENKENKANOYOBI | False／True／True | name=点検可能曜日; type=String; length=1; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | YAKANTAIOUMU\_MITSUMORI\_ComboBox | YAKANTAIOUMU | False／True／True | name=夜間対応有無; type=String; length=1; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_IRAI\_ID\_MITSUMORI\_TextBox | TM\_IRAI\_ID | False／False／True | name=取引先見積依頼ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TMKEIYAKUJIKAN\_ID\_MITSUMORI\_TextBox | TM\_KEIYAKUJIKAN\_ID | False／False／True | name=取引先見積契約時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | MCM\_TM\_KIKANDataGridView | TM\_KEIYAKUJIKAN\_ID | 未指定／未指定／未指定 | name=取引先見積契約時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_MITSUMORI\_NO\_KIKAN\_TextBox | TM\_MITSUMORI\_NO | False／False／True | name=取引先見積NO; type=String; length=20; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | KAISI\_DT\_KIKAN\_TextBox | KAISI\_DT | False／False／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SYURYO\_DT\_KIKAN\_TextBox | SYURYO\_DT | False／False／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | HYOJUNGOKEI\_KIN\_KIKAN\_TextBox | HYOJUNGOKEI\_KIN | False／False／True | name=標準価格合計; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SIKIRISYOKEI\_KIN\_KIKAN\_TextBox | SIKIRISYOKEI\_KIN | False／False／True | name=仕切り価格小計; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SYUSSEINEBIKI\_KIN\_KIKAN\_TextBox | SYUSSEINEBIKI\_KIN | False／False／True | name=出精値引き; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | SIKIRIGOKEI\_KIN\_KIKAN\_TextBox | SIKIRIGOKEI\_KIN | False／False／True | name=仕切り価格合計; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_KEIYAKUJIKAN\_ID\_\_KIKAN\_TextBox | TM\_KEIYAKUJIKAN\_ID | False／False／True | name=取引先見積契約時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | TM\_KIKAN\_ID\_KIKAN\_TextBox | TM\_KIKAN\_ID | False／False／True | name=取引先見積期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) |
| 1006 | CHECK\_FLG\_KIKIKOSEI\_CheckBox | CHECK\_FLG | 未指定／False／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_IRAI\_NO\_KIKIKOSEI\_TextBox | TM\_IRAI\_NO | False／False／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKOSEI\_NK\_KIKIKOSEI\_TextBox | KIKIKOSEI\_NK | False／False／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SET\_NM\_KIKIKOSEI\_TextBox | SET\_NM | False／True／True | name=セット数; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TANI\_KIKIKOSEI\_TextBox | TANI | False／True／True | name=単位; type=String; length=10; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TEHAISEIBAN\_KIKIKOSEI\_TextBox | TEHAISEIBAN | False／True／True | name=手配製番; type=String; length=50; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | BIKODataGrid\_KIKIKOSEI\_TextBox | BIKO | False／True／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKOSEI\_ID\_KIKIKOSEI\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_KIKIKOSEI\_ID\_KIKIKOSEI\_TextBox | TM\_KIKIKOSEI\_ID | False／False／True | name=取引先見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TK\_KIKIKOSEI\_ID\_KIKIKOSEI\_TextBox | TK\_KIKIKOSEI\_ID | False／False／True | name=取引先契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MITSUMORICHECK\_FLG\_KIKIKOSEI\_TextBox | MITSUMORICHECK\_FLG | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | CHECK\_FLG\_KIKIMEISAI\_CheckBox | CHECK\_FLG | 未指定／False／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_IRAI\_NO\_KIKIMEISAI\_TextBox | TM\_IRAI\_NO | False／False／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SEIZOMAKER\_NK\_KIKIMEISAI\_TextBox | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIHINMEI\_NK\_KIKIMEISAI\_TextBox | KIKIHINMEI\_NK | False／True／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKATASHIKI\_KIKIMEISAI\_TextBox | KIKIKATASHIKI | False／True／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SURYO\_NM\_KIKIMEISAI\_TextBox | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | BIKO\_KIKIMEISAI\_TextBox | BIKO | False／True／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIMEISAI\_ID\_KIKIMEISAI\_TextBox | KIKIMEISAI\_ID | False／True／True | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKOSEI\_ID\_KIKIMEISAI\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_KIKIKOSEI\_ID\_KIKIMEISAI\_TextBox | TM\_KIKIKOSEI\_ID | False／False／True | name=取引先見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_KIKIMEISAI\_ID\_KIKIMEISAI\_TextBox | TM\_KIKIMEISAI\_ID | False／False／True | name=取引先見積機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TK\_KIKIKOSEI\_ID\_KIKIMEISAI\_TextBox | TK\_KIKIKOSEI\_ID | False／False／True | name=取引先契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TK\_KIKIMEISAI\_ID\_KIKIMEISAI\_TextBox | TK\_KIKIMEISAI\_ID | False／False／True | name=取引先契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MITSUMORICHECK\_FLG\_KIKIMEISAI\_TextBox | MITSUMORICHECK\_FLG | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | CHECK\_FLG\_OLD\_KIKIKOSEI\_TestBox | CHECK\_FLG\_OLD | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | CHECK\_FLG\_KOTAIMEI\_CheckBox | CHECK\_FLG | 未指定／False／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_IRAI\_NO\_KOTAIMEI\_TextBox | TM\_IRAI\_NO | False／False／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAF\_KIKIHINMEI\_NK\_KOTAIMEI\_TextBox | KIKIHINMEI\_NK | False／False／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAF\_KIKIKATASHIKI\_KOTAIMEI\_TextBox | KIKIKATASHIKI | False／False／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAF\_SURYO\_NM\_KOTAIMEI\_TextBox | SURYO\_NM | False／False／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAG\_KOTAI\_NK\_KOTAIMEI\_TextBox | KOTAI\_NK | False／False／True | name=個体名; type=String; length=60; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAG\_SERIAL\_NO\_KOTAIMEI\_TextBox | SERIAL\_NO | False／False／True | name=シリアル番号; type=String; length=100; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAG\_SETCHIBASYO\_KOTAIMEI\_TextBox | SETCHIBASYO | False／False／True | name=設置場所; type=String; length=60; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAK\_TORIHIKISAKI\_NK\_KOTAIMEI\_TextBox | TORIHIKISAKI\_NK | False／False／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAD\_BRANDSYOSAI\_NK\_KOTAIMEI\_TextBox | BRANDSYOSAI\_NK | False／False／True | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAG\_ITIJINONYU\_DT\_KOTAIMEI\_TextBox | ITIJINONYU\_DT | False／False／True | name=一次納入日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MAG\_ENCHOKEIYAKUKIGEN\_DT\_KOTAIMEI\_TextBox | ENCHOKEIYAKUKIGEN\_DT | False／False／True | name=延長契約期限; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKOSEI\_ID\_KOTAIMEI\_TextBox | KIKIKOSEI\_ID | False／False／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIMEISAI\_ID\_KOTAIMEI\_TextBox | KIKIMEISAI\_ID | False／False／True | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KOTAIKANRI\_ID\_KOTAIMEI\_TextBox | KOTAIKANRI\_ID | False／False／True | name=個体管理ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_KIKIKOSEI\_ID\_KOTAIMEI\_TextBox | TM\_KIKIKOSEI\_ID | False／False／True | name=取引先見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_KIKIMEISAI\_ID\_KOTAIMEI\_TextBox | TM\_KIKIMEISAI\_ID | False／False／True | name=取引先見積機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TM\_KOTAIMEISAI\_ID\_KOTAIMEI\_TextBox | TM\_KOTAIMEISAI\_ID | False／False／True | name=取引先見積個体明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TK\_KIKIMEISAI\_ID\_KOTAIMEI\_TextBox | TK\_KIKIMEISAI\_ID | False／False／True | name=取引先契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TK\_KIKIKOSEI\_ID\_KOTAIMEI\_TextBox | TK\_KIKIKOSEI\_ID | False／False／True | name=取引先契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TK\_KOTAIMEISAI\_ID\_KOTAIMEI\_TextBox | TK\_KOTAIMEISAI\_ID | False／False／True | name=取引先契約個体明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MITSUMORICHECK\_FLG\_KOTAIMEI\_TextBox | MITSUMORICHECK\_FLG | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | CHECK\_FLG\_OLD\_KOTAIMEI\_TextBox | CHECK\_FLG\_OLD | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | CHECKFLGDataGridViewTextBoxColumn | CHECK\_FLG | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | MITSUMORICHECKFLGDataGridViewTextBoxColumn | MITSUMORICHECK\_FLG | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | CHECKFLGOLDDataGridViewTextBoxColumn | CHECK\_FLG\_OLD | 未指定／未指定／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKOSEIIDDataGridViewTextBoxColumn | KIKIKOSEI\_ID | 未指定／未指定／未指定 | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIMEISAIIDDataGridViewTextBoxColumn | KIKIMEISAI\_ID | 未指定／未指定／未指定 | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KOTAIKANRIIDDataGridViewTextBoxColumn | KOTAIKANRI\_ID | 未指定／未指定／未指定 | name=個体管理ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | ATSUKAIKIKIIDDataGridViewTextBoxColumn | ATSUKAIKIKI\_ID | 未指定／未指定／未指定 | name=取扱機器ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KOTAINKDataGridViewTextBoxColumn | KOTAI\_NK | 未指定／未指定／未指定 | name=個体名; type=String; length=60; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SERIALNODataGridViewTextBoxColumn | SERIAL\_NO | 未指定／未指定／未指定 | name=シリアル番号; type=String; length=100; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | ITIJINONYUDTDataGridViewTextBoxColumn | ITIJINONYU\_DT | 未指定／未指定／未指定 | name=一次納入日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SETCHIBASYODataGridViewTextBoxColumn | SETCHIBASYO | 未指定／未指定／未指定 | name=設置場所; type=String; length=60; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIHINMEINKDataGridViewTextBoxColumn | KIKIHINMEI\_NK | 未指定／未指定／未指定 | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KIKIKATASHIKIDataGridViewTextBoxColumn | KIKIKATASHIKI | 未指定／未指定／未指定 | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SURYONMDataGridViewTextBoxColumn | SURYO\_NM | 未指定／未指定／未指定 | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | SETNMDataGridViewTextBoxColumn | SET\_NM | 未指定／未指定／未指定 | name=セット数; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TORIHIKISAKINKDataGridViewTextBoxColumn | TORIHIKISAKI\_NK | 未指定／未指定／未指定 | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | BRANDSYOSAINKDataGridViewTextBoxColumn | BRANDSYOSAI\_NK | 未指定／未指定／未指定 | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TEKKYODTDataGridViewTextBoxColumn | TEKKYO\_DT | 未指定／未指定／未指定 | name=撤去日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | KEIYAKUKIGENDTDataGridViewTextBoxColumn | KEIYAKUKIGEN\_DT | 未指定／未指定／未指定 | name=契約期限; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | ENCHOKEIYAKUKIGENDTDataGridViewTextBoxColumn | ENCHOKEIYAKUKIGEN\_DT | 未指定／未指定／未指定 | name=延長契約期限; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TMIRAINODataGridViewTextBoxColumn | TM\_IRAI\_NO | 未指定／未指定／未指定 | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TMKIKIKOSEIIDDataGridViewTextBoxColumn | TM\_KIKIKOSEI\_ID | 未指定／未指定／未指定 | name=取引先見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TMKIKIMEISAIIDDataGridViewTextBoxColumn | TM\_KIKIMEISAI\_ID | 未指定／未指定／未指定 | name=取引先見積機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TMKOTAIMEISAIIDDataGridViewTextBoxColumn | TM\_KOTAIMEISAI\_ID | 未指定／未指定／未指定 | name=取引先見積個体明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TKKIKIKOSEIIDDataGridViewTextBoxColumn | TK\_KIKIKOSEI\_ID | 未指定／未指定／未指定 | name=取引先契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TKKIKIMEISAIIDDataGridViewTextBoxColumn | TK\_KIKIMEISAI\_ID | 未指定／未指定／未指定 | name=取引先契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1006 | TKKOTAIMEISAIIDDataGridViewTextBoxColumn | TK\_KOTAIMEISAI\_ID | 未指定／未指定／未指定 | name=取引先契約個体明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) |
| 1009 | TM\_IRAI\_ID\_MCM\_1003\_V\_TextBox | TM\_IRAI\_ID | False／False／True | name=取引先見積依頼ID; type=Decimal; length=; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TK\_KEIYAKU\_ID\_MCM\_1003\_V\_TextBox | TK\_KEIYAKU\_ID | False／False／True | name=取引先契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | JOTAI\_MCM\_1003\_V\_TextBox | JOTAI | False／False／True | name=状態; type=String; length=1; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TM\_IRAI\_NO\_MCM\_1003\_V\_Link | TM\_IRAI\_NO | 未指定／未指定／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | MITSUMORI\_DT\_MCM\_1003\_V\_TextBox | MITSUMORI\_DT | False／False／True | name=見積日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TM\_KEIYAKUJIKAN\_ID\_MCM\_1003\_V\_TextBox | TM\_KEIYAKUJIKAN\_ID | False／False／True | name=取引先見積契約時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KEIYAKUJIKANTAI\_MCM\_1003\_V\_TextBox | KEIYAKUJIKANTAI | False／False／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | SHONINJOTAI | SHONINJOTAI | False／False／True | name=承認状態; type=String; length=1; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KEIYAKU\_NO\_MCM\_1003\_V\_Link | KEIYAKU\_NO | 未指定／未指定／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | FUSEI\_MCM\_1003\_V\_TextBox | FUSEI | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TORIHIKISAKI\_ID\_MCM\_1003\_V\_TextBox | TORIHIKISAKI\_ID | False／False／True | name=取引先ID; type=Decimal; length=; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TORIHIKISAKI\_CD\_MCM\_1003\_V\_TextBox | TORIHIKISAKI\_CD | False／False／True | name=取引先コード; type=String; length=20; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TORIHIKISAKI\_NK\_MCM\_1003\_V\_TextBox | TORIHIKISAKI\_NK | False／False／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | HOSYUHOHO\_MCM\_1003\_V\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TENKENUMU\_MCM\_1003\_V\_TextBox | TENKENUMU | False／False／True | name=点検有無; type=String; length=1; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | TENKENKANOYOBI\_MCM\_1003\_V\_TextBox | TENKENKANOYOBI | False／False／True | name=点検可能曜日; type=String; length=1; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | YAKANTAIOUMU\_MCM\_1003\_V\_TextBox | YAKANTAIOUMU | False／False／True | name=夜間対応有無; type=String; length=1; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KEIYAKU\_DT\_MCM\_1003\_V\_TextBox | KEIYAKU\_DT | False／False／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KAIYAKU\_DT\_MCM\_1003\_V\_TextBox | KAIYAKU\_DT | False／False／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | NONYUSAKI\_ID\_MCM\_1003\_V\_TextBox | NONYUSAKI\_ID | False／False／True | name=納入先ID; type=Decimal; length=; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | NONYUSAKI\_CD\_MCM\_1003\_V\_TextBox | NONYUSAKI\_CD | False／False／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | NONYUSAKI\_NK\_MCM\_1003\_V\_TextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KYUNONYUSAKI\_NK\_MCM\_1003\_V\_TextBox | KYUNONYUSAKI\_NK | False／False／True | name=旧納入先名; type=String; length=80; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | NONYUSAKIKOJO\_NK\_MCM\_1003\_V\_TextBox | NONYUSAKIKOJO\_NK | False／False／True | name=納入先工場名; type=String; length=80; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | SUPPORT\_ID\_MCM\_1003\_V\_TextBox | SUPPORT\_ID | False／False／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | PLANT\_NK\_MCM\_1003\_V\_TextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | PLANT\_ID\_MCM\_1003\_V\_TextBox | PLANT\_ID | False／False／True | name=プラントID; type=Decimal; length=; min=; max=; mask= | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KEIYAKU\_KEIYAKU\_MCM\_1003\_V\_Link | KEIYAKU\_KEIYAKU | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1009 | KEIYAKU\_DEL\_MCM\_1003\_V\_Link | KEIYAKU\_DEL | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) |
| 1010 | TM\_IRAI\_ID\_MCM\_1003\_V\_TextBox | TM\_IRAI\_ID | False／False／True | name=取引先見積依頼ID; type=Decimal; length=; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TK\_KEIYAKU\_ID\_MCM\_1003\_V\_TextBox | TK\_KEIYAKU\_ID | False／False／True | name=取引先契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | JOTAI\_MCM\_1003\_V\_TextBox | JOTAI | False／False／True | name=状態; type=String; length=1; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TM\_IRAI\_NO\_MCM\_1003\_V\_Link | TM\_IRAI\_NO | 未指定／未指定／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | MITSUMORI\_DT\_MCM\_1003\_V\_TextBox | MITSUMORI\_DT | False／False／True | name=見積日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TM\_KEIYAKUJIKAN\_ID\_MCM\_1003\_V\_TextBox | TM\_KEIYAKUJIKAN\_ID | False／False／True | name=取引先見積契約時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KEIYAKUJIKANTAI\_MCM\_1003\_V\_TextBox | KEIYAKUJIKANTAI | False／False／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | SHONINJOTAI | SHONINJOTAI | False／False／True | name=承認状態; type=String; length=1; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KEIYAKU\_NO\_MCM\_1003\_V\_Link | KEIYAKU\_NO | 未指定／未指定／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | FUSEI\_MCM\_1003\_V\_TextBox | FUSEI | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | JIGYOSHO\_CD\_1003\_V\_TextBox | JIGYOSHO\_CD | False／False／True | name=事業所コード; type=String; length=10; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | MEISHO4\_NK\_1003\_V\_TextBox | MEISHO4\_NK | False／False／True | name=名称４; type=String; length=100; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TORIHIKISAKI\_ID\_MCM\_1003\_V\_TextBox | TORIHIKISAKI\_ID | False／False／True | name=取引先ID; type=Decimal; length=; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TORIHIKISAKI\_CD\_MCM\_1003\_V\_TextBox | TORIHIKISAKI\_CD | False／False／True | name=取引先コード; type=String; length=20; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TORIHIKISAKI\_NK\_MCM\_1003\_V\_TextBox | TORIHIKISAKI\_NK | False／False／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | HOSYUHOHO\_MCM\_1003\_V\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TENKENUMU\_MCM\_1003\_V\_TextBox | TENKENUMU | False／False／True | name=点検有無; type=String; length=1; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | TENKENKANOYOBI\_MCM\_1003\_V\_TextBox | TENKENKANOYOBI | False／False／True | name=点検可能曜日; type=String; length=1; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | YAKANTAIOUMU\_MCM\_1003\_V\_TextBox | YAKANTAIOUMU | False／False／True | name=夜間対応有無; type=String; length=1; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KEIYAKU\_DT\_MCM\_1003\_V\_TextBox | KEIYAKU\_DT | False／False／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KAIYAKU\_DT\_MCM\_1003\_V\_TextBox | KAIYAKU\_DT | False／False／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | NONYUSAKI\_ID\_MCM\_1003\_V\_TextBox | NONYUSAKI\_ID | False／False／True | name=納入先ID; type=Decimal; length=; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | NONYUSAKI\_CD\_MCM\_1003\_V\_TextBox | NONYUSAKI\_CD | False／False／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | NONYUSAKI\_NK\_MCM\_1003\_V\_TextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KYUNONYUSAKI\_NK\_MCM\_1003\_V\_TextBox | KYUNONYUSAKI\_NK | False／False／True | name=旧納入先名; type=String; length=80; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | NONYUSAKIKOJO\_NK\_MCM\_1003\_V\_TextBox | NONYUSAKIKOJO\_NK | False／False／True | name=納入先工場名; type=String; length=80; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | SUPPORT\_ID\_MCM\_1003\_V\_TextBox | SUPPORT\_ID | False／False／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | PLANT\_NK\_MCM\_1003\_V\_TextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | PLANT\_ID\_MCM\_1003\_V\_TextBox | PLANT\_ID | False／False／True | name=プラントID; type=Decimal; length=; min=; max=; mask= | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KEIYAKU\_KEIYAKU\_MCM\_1003\_V\_Link | KEIYAKU\_KEIYAKU | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1010 | KEIYAKU\_DEL\_MCM\_1003\_V\_Link | KEIYAKU\_DEL | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) |
| 1011 | TENPUFILE\_NK\_TENPU\_Link | TENPUFILE\_NK | 未指定／未指定／True | name=添付ファイル名; type=String; length=2000; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | DIRECTORY\_TENPU\_TextBox | DIRECTORY | False／False／未指定 | name=ディレクトリ; type=String; length=2000; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | TENPU\_ID\_TENPU\_TextBox | TENPU\_ID | False／True／未指定 | name=添付ID; type=Decimal; length=; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | CHECK\_FLG\_MITSUMORI\_TextBox | CHECK\_FLG | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SYOUNIN\_JOTAI\_MITSUMORI\_ComboBox | SYOUNIN\_JOTAI | False／False／True | name=承認状態; type=String; length=1; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SHORUI\_NO\_MITSUMORI\_Link | SHORUI\_NO | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | KAISI\_DT\_MITSUMORI\_TextBox | KAISI\_DT | False／False／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | KEIYAKUJIKANTAI\_MCM\_UM\_MITSUMORI\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | TORIHIKISAKI\_NK\_MCM\_UM\_MITSUMORI\_TextBox | TORIHIKISAKI\_NK | False／True／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | NONYUSAKI\_NK\_MCM\_UM\_MITSUMORI\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | KYUNONYUSAKI\_NK\_MCM\_UM\_MITSUMORI\_TextBox | KYUNONYUSAKI\_NK | False／True／True | name=旧納入先名; type=String; length=80; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | NONYUSAKIKOJO\_NK\_MCM\_UM\_MITSUMORI\_TextBox | NONYUSAKIKOJO\_NK | False／True／True | name=納入先工場名; type=String; length=80; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SUPPORT\_ID\_MCM\_UM\_MITSUMORI\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | PLANT\_NK\_MCM\_UM\_MITSUMORI\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | KINGAKU\_MITSUMORI\_TextBox | KINGAKU | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | IRAITANTO\_NK\_MITSUMORI\_TextBox | IRAITANTO\_NK | False／False／True | name=依頼元担当者; type=String; length=100; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | JOTAI\_MITSUMORI\_ComboBox | JOTAI | False／False／True | name=状態; type=String; length=1; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SHINSA\_DT | SHINSA\_DT | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SHINSA\_BY | SHINSA\_BY | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SYOUNIN\_DT | SYOUNIN\_DT | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | SYOUNIN\_BY | SYOUNIN\_BY | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | RELATION\_ID\_MITSUMORI\_TextBox | RELATION\_ID | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 1011 | TK\_KEIYAKU\_ID\_MITSUMORI\_TextBox | TK\_KEIYAKU\_ID | False／False／True | name=取引先契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) |
| 2006 | KEIYAKU\_NOTextBox | KEIYAKU\_NO | False／False／未指定 | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KEIYAKU\_DTTextBox | KEIYAKU\_DT | True／False／未指定 | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | SYOKAI\_KEIYAKU\_DTCPTextBox | SHOKAI\_KEIYAKU\_DT | True／False／未指定 | name=初回契約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | AUTO\_FLGCheckBox | AUTO\_FLG | 未指定／未指定／未指定 | name=自動更新フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KEIYAKUMANRYO\_DTTextBox | KEIYAKUMANRYO\_DT | False／False／未指定 | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | ENTYOKEIYAKUMANRYO\_DTCPTextBox | ENTYOKEIYAKUMANRYO\_DT | False／False／未指定 | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAIYAKU\_DTCPTextBox | KAIYAKU\_DT | False／False／未指定 | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | BIKOCPTextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | SYOKAI\_MANRYOU\_DTCPTextBox | JIKAIKOSIN\_DT | False／False／True | name=次回更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | CREATED\_DTCPTextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | CREATED\_BYCPTextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | LASTUPDATE\_DTCPTextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | LASTUPDATE\_BYCPTextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAISI\_DT\_SEIBAN\_TextBox | KAISI\_DT | False／False／未指定 | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | SYURYO\_DT\_SEIBAN\_TextBox | SYURYO\_DT | False／False／未指定 | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | HARD\_SEIBAN\_SEIBAN\_TextBox | HARD\_SEIBAN | False／False／未指定 | name=ハード製番; type=String; length=20; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | SOFT\_SEIBAN\_SEIBAN\_TextBox | SOFT\_SEIBAN | False／False／未指定 | name=ソフト製番; type=String; length=20; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | BIKO\_SEIBAN\_TextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAKUNIN\_KBN\_SEIBAN\_ComboBox | KAKUNIN\_KBN | False／False／True | name=確認区分; type=String; length=1; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAKUNIN\_DT\_SEIBAN\_TextBox | KAKUNIN\_DT | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAKUNIN\_BY\_SEIBAN\_TextBox | KAKUNIN\_BY | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAKUNINIRAI\_DT\_SEIBAN\_TextBox | KAKUNINIRAI\_DT | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | KAKUNINIRAI\_BY\_SEIBAN\_TextBox | KAKUNINIRAI\_BY | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | CREATED\_DT\_SEIBAN\_TextBox | CREATED\_DT | False／True／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | CREATED\_BY\_SEIBAN\_TextBox | CREATED\_BY | False／True／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | LASTUPDATE\_DT\_SEIBAN\_TextBox | LASTUPDATE\_DT | False／True／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | LASTUPDATE\_BY\_SEIBAN\_TextBox | LASTUPDATE\_BY | False／True／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | UK\_SEIBAN\_ID\_SEIBAN\_TextBox | UK\_SEIBAN\_ID | False／True／True | name=ユーザー契約製番ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | UK\_KEIYAKU\_ID\_SEIBAN\_TextBox | UK\_KEIYAKU\_ID | False／True／True | name=ユーザー契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) |
| 2006 | NONYUSAKI\_CDTextBox | NONYUSAKI\_CD | False／False／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | NONYUSAKI\_NKTextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | SUPPORT\_IDTextBox | SUPPORT\_ID | False／False／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | PLANT\_NKTextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KEIYAKUJIKANTAITextBox | KEIYAKUJIKANTAI | False／False／未指定 | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KAISI\_DTTextBox | KAISI\_DT | True／False／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | SYURYO\_DTTextBox | SYURYO\_DT | True／False／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | HOSYU\_GKINTextBox | HOSYU\_GKIN | False／False／未指定 | name=保守金額合計; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BRAND\_NK\_BRAND\_TextBox | BRAND\_NK | False／True／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BRANDSYOSAI\_NK\_BRAND\_TextBox | BRANDSYOSAI\_NK | False／True／True | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KEIYAKUJIKANTAI\_KEIYAKU\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | HOSYUHOHO\_BRAND\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | SOFT\_FLG\_BRAND\_CheckBox | SOFT\_FLG | 未指定／False／未指定 | name=ソフトウェア契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | DREMOS\_FLG\_BRAND\_CheckBox | DREMOS\_FLG | 未指定／False／未指定 | name=DREMOS契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | DREMOS\_NM\_BRAND\_TextBox | DREMOS\_NM | False／False／未指定 | name=DREMOS契約台数; type=Decimal; length=4; min=0; max=9999; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | REMOTE\_FLG\_BRAND\_CheckBox | REMOTE\_FLG | 未指定／False／未指定 | name=リモート保守フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | REMOTERENRAKUSAKI\_BRAND\_TextBox | REMOTERENRAKUSAKI | False／False／未指定 | name=リモート連絡先; type=String; length=200; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BIKO\_BRAND\_TextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_DT\_BRAND\_TextBox | CREATED\_DT | False／True／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_BY\_BRAND\_TextBox | CREATED\_BY | False／True／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_DT\_BRAND\_TextBox | LASTUPDATE\_DT | False／True／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_BY\_BRAND\_TextBox | LASTUPDATE\_BY | False／True／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_BRAND\_ID\_BRAND\_TextBox | UK\_BRAND\_ID | False／True／True | name=ユーザー契約ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_KIKAN\_ID\_BRAND\_TextBox | UK\_KIKAN\_ID | False／True／True | name=ユーザー契約期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BRANDKOSEI\_ID\_BRAND\_TextBox | BRANDKOSEI\_ID | False／True／True | name=ブランド構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KIKIKOSEI\_NK\_KIKIKOSEI\_TextBox | KIKIKOSEI\_NK | False／True／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | SET\_NM\_KIKIKOSEI\_TextBox | SET\_NM | False／False／True | name=セット数; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | TANI\_KIKIKOSEI\_TextBox | TANI | False／False／True | name=単位; type=String; length=10; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CONTROLLER\_FLG\_KIKIKOSEI\_CheckBox | CONTROLLER\_FLG | 未指定／False／True | name=コントローラフラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | TEHAISEIBAN\_KIKIKOSEI\_TextBox | TEHAISEIBAN | False／True／True | name=手配製番; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BIKO\_KIKIKOSEI\_TextBox | BIKO | False／True／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_DT\_KIKIKOSEI\_TextBox | CREATED\_DT | False／True／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_BY\_KIKIKOSEI\_TextBox | CREATED\_BY | False／True／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_DT\_KIKIKOSEI\_TextBox | LASTUPDATE\_DT | False／True／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_BY\_KIKIKOSEI\_TextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_KIKIKOSEI\_ID\_KIKIKOSEI\_TextBox | UK\_KIKIKOSEI\_ID | False／True／True | name=ユーザー契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_BRAND\_ID\_KIKIKOSEI\_TextBox | UK\_BRAND\_ID | False／True／True | name=ユーザー契約ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KIKIKOSEI\_ID\_KIKIKOSEI\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | HYOJIJUN\_KIKIKOSEI\_TextBox | HYOJIJUN | False／True／True | name=表示順; type=Decimal; length=6; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | SEIZOMAKER\_NK\_MEISAI\_TextBox | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KIKIHINMEI\_NK\_MEISAI\_TextBox | KIKIHINMEI\_NK | False／True／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KIKIKATASHIKI\_MEISAI\_TextBox | KIKIKATASHIKI | False／False／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | SURYO\_NM\_MEISAI\_TextBox | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | PACK\_FLG\_MEISAI\_CheckBox | PACK\_FLG | 未指定／False／未指定 | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KEIYAKUNAIYO\_MEISAI\_TextBox | KEIYAKUNAIYO | False／False／未指定 | name=契約内容; type=String; length=400; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KEIYAKU\_NO\_MEISAI\_TextBox | KEIYAKU\_NO | False／False／未指定 | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | DAIFUKUHOSYUJIKAN\_ID\_MEISAI\_ComboBox | DAIFUKUHOSYUJIKAN\_ID | False／False／未指定 | name=ダイフク保守時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | TENKENKAISU\_MEISAI\_TextBox | TENKENKAISU | False／False／未指定 | name=点検回数; type=Decimal; length=2; min=0; max=12; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | TENKENYOBI\_MEISAI\_ComboBox | TENKENYOBI | False／False／未指定 | name=点検曜日; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | HOSYUHOHO\_MEISAI\_TextBox | HOSYUHOHO | False／False／未指定 | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BIKO\_MEISAI\_TextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_DT\_MEISAI\_TextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_BY\_MEISAI\_TextBox | CREATED\_BY | False／True／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_DT\_MEISAI\_TextBox | LASTUPDATE\_DT | False／True／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_BY\_MEISAI\_TextBox | LASTUPDATE\_BY | False／True／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_KIKIMEISAI\_ID\_MEISAI\_TextBox | UK\_KIKIMEISAI\_ID | False／False／True | name=ユーザー契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_KIKIKOSEI\_ID\_MEISAI\_TextBox | UK\_KIKIKOSEI\_ID | False／True／True | name=ユーザー契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | KIKIMEISAI\_ID\_MEISAI\_TextBox | KIKIMEISAI\_ID | False／True／True | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | HYOJIJUN\_MEISAI\_TextBox | HYOJIJUN | False／True／True | name=表示順; type=Decimal; length=6; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | NAIYO\_TENKEN\_TextBox | NAIYO | False／False／未指定 | name=内容; type=String; length=4000; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M01\_TENKEN\_CheckBox | M01 | 未指定／False／未指定 | name=1月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M02\_TENKEN\_CheckBox | M02 | 未指定／False／未指定 | name=2月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M03\_TENKEN\_CheckBox | M03 | 未指定／False／未指定 | name=3月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M04\_TENKEN\_CheckBox | M04 | 未指定／False／未指定 | name=4月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M05\_TENKEN\_CheckBox | M05 | 未指定／False／未指定 | name=5月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M06\_TENKEN\_CheckBox | M06 | 未指定／False／未指定 | name=6月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M07\_TENKEN\_CheckBox | M07 | 未指定／False／未指定 | name=7月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M08\_TENKEN\_CheckBox | M08 | 未指定／False／未指定 | name=8月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M09\_TENKEN\_CheckBox | M09 | 未指定／False／未指定 | name=9月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M10\_TENKEN\_CheckBox | M10 | 未指定／False／未指定 | name=10月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M11\_TENKEN\_CheckBox | M11 | 未指定／False／未指定 | name=11月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | M12\_TENKEN\_CheckBox | M12 | 未指定／False／未指定 | name=12月; type=String; length=1; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BIKO\_TENKEN\_TextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_DT\_TENKEN\_TextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_BY\_TENKEN\_TextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_DT\_TENKEN\_TextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_BY\_TENKEN\_TextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_KIKAN\_ID\_TENKEN\_TextBox | UK\_KIKAN\_ID | False／False／True | name=ユーザー契約期間ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | UK\_TENKEN\_ID\_TENKEN\_TextBox | UK\_TENKEN\_ID | False／False／True | name=ユーザー契約点検ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_DTTextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | CREATED\_BYTextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_DTTextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | LASTUPDATE\_BYTextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | BIKOTextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | IRAITENPO\_IDCombobox | IRAITENPO\_ID | False／True／未指定 | name=依頼元店舗ID; type=Decimal; length=; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2006 | IRAITANTO\_NKTextBox | IRAITANTO\_NK | False／False／未指定 | name=依頼元担当者; type=String; length=100; min=; max=; mask= | [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_NO\_BRAND\_TextBox | UM\_MITSUMORI\_NO | False／True／True | name=店舗見積NO; type=String; length=20; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRANDSYOSAI\_NK\_BRAND\_TextBox | BRANDSYOSAI\_NK | False／True／True | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KEIYAKUJIKANTAI\_BRAND\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SOFTHOSYUHOHO\_BRAND\_TextBox | SOFTHOSYUHOHO | False／True／True | name=ソフト保守方法; type=String; length=4000; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | DREMOS\_FLG\_BRAND\_CheckBox | DREMOS\_FLG | 未指定／True／True | name=DREMOS契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | REMOTE\_FLG\_BRAND\_CheckBox | REMOTE\_FLG | 未指定／True／True | name=リモート保守フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SOFT\_FLG\_BRAND\_CheckBox | SOFT\_FLG | 未指定／True／True | name=ソフトウェア契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRANDKOSEI\_ID\_BRAND\_TextBox | BRANDKOSEI\_ID | False／True／True | name=ブランド構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRAND\_ID\_BRAND\_TextBox | BRAND\_ID | False／True／True | name=ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_KOSEI\_CheckBox | CHECK\_FLG | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_NO\_KOSEI\_TextBox | UM\_MITSUMORI\_NO | False／True／True | name=店舗見積NO; type=String; length=20; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIKOSEI\_NK\_KOSEI\_TextBox | KIKIKOSEI\_NK | False／True／True | name=機器構成名; type=String; length=100; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SET\_NM\_KOSEI\_TextBox | SET\_NM | False／True／True | name=セット数; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | TANI\_KOSEI\_TextBox | TANI | False／True／True | name=単位; type=String; length=10; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | TEHAISEIBAN\_KOSEI\_TextBox | TEHAISEIBAN | False／True／True | name=手配製番; type=String; length=50; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BIKO\_KOSEI\_TextBox | BIKO | False／True／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIKOSEI\_ID\_KOSEI\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRANDKOSEI\_ID\_KOSEI\_TextBox | BRANDKOSEI\_ID | False／True／True | name=ブランド構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIHON\_BRAND\_ID\_KOSEI\_TextBox | UM\_KIHON\_BRAND\_ID | False／True／True | name=店舗基本見積ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIKIKOSEI\_ID\_KOSEI\_TextBox | UM\_KIKIKOSEI\_ID | False／True／True | name=店舗見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MITSUMORICHECK\_FLG\_KIKIKOSEI\_TextBox | MITSUMORICHECK\_FLG | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_OLD\_KIKIKOSEI\_TextBox | CHECK\_FLG\_OLD | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MAE\_HYOJIJUN\_KOSEI\_TextBox | MAE\_HYOJIJUN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UK\_KIKIKOSEI\_ID\_KOSEI\_TextBox | UK\_KIKIKOSEI\_ID | False／True／未指定 | name=ユーザー契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_MEISAI\_CheckBox | CHECK\_FLG | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_NO\_MEISAI\_TextBox | UM\_MITSUMORI\_NO | False／True／True | name=店舗見積NO; type=String; length=20; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SEIZOMAKER\_NK\_MEISAI\_TextBox | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIHINMEI\_NK\_MEISAI\_TextBox | KIKIHINMEI\_NK | False／True／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIKATASHIKI\_MEISAI\_TextBox | KIKIKATASHIKI | False／True／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SURYO\_NM\_MEISAI\_TextBox | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BIKO\_MEISAI\_TextBox | BIKO | False／True／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | HYOJIJUN\_MEISAI\_TextBox | HYOJIJUN | False／True／True | name=表示順; type=Decimal; length=6; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRANDKOSEI\_ID\_MEISAI\_TextBox | BRANDKOSEI\_ID | False／True／True | name=ブランド構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIMEISAI\_ID\_MEISAI\_TextBox | KIKIMEISAI\_ID | False／True／True | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIKOSEI\_ID\_MEISAI\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | ATSUKAIKIKI\_ID\_MEISAI\_TextBox | ATSUKAIKIKI\_ID | False／True／True | name=取扱機器ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SEIZOMAKER\_ID\_MEISAI\_TextBox | SEIZOMAKER\_ID | False／True／True | name=製造メーカーID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIHON\_BRAND\_ID\_MEISAI\_TextBox | UM\_KIHON\_BRAND\_ID | False／True／True | name=店舗基本見積ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIKIKOSEI\_ID\_MEISAI\_TextBox | UM\_KIKIKOSEI\_ID | False／True／True | name=店舗見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIKIMEISAI\_ID\_MEISAI\_TextBox | UM\_KIKIMEISAI\_ID | False／True／True | name=店舗見積機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UK\_KIKIKOSEI\_ID\_MEISAI\_TextBox | UK\_KIKIKOSEI\_ID | False／True／True | name=ユーザー契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UK\_KIKIMEISAI\_ID\_MEISAI\_TextBox | UK\_KIKIMEISAI\_ID | False／True／True | name=ユーザー契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MITSUMORICHECK\_FLG\_MEISAI\_TextBox | MITSUMORICHECK\_FLG | False／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_OLD\_MEISAI\_TextBox | CHECK\_FLG\_OLD | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MAE\_HYOJIJUN\_MEISAI\_TextBox | MAE\_HYOJIJUN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MAF\_HYOJIJUN\_MEISAI\_TextBox | MAF\_HYOJIJUN | False／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_KOTAI\_CheckBox | CHECK\_FLG | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_NO\_KOTAI\_TextBox | UM\_MITSUMORI\_NO | False／True／True | name=店舗見積NO; type=String; length=20; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIHINMEI\_NK\_KOTAI\_TextBox | KIKIHINMEI\_NK | False／True／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIKATASHIKI\_KOSEI\_TextBox | KIKIKATASHIKI | False／True／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SURYO\_NM\_KOSEI\_TextBox | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KOTAI\_NK\_KOTAI\_TextBox | KOTAI\_NK | False／True／True | name=個体名; type=String; length=60; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SERIAL\_NO\_KOSEI\_TextBox | SERIAL\_NO | False／True／True | name=シリアル番号; type=String; length=100; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SETCHIBASYO\_KOTAI\_TextBox | SETCHIBASYO | False／True／True | name=設置場所; type=String; length=60; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | TORIHIKISAKI\_NK\_KOTAI\_TextBox | TORIHIKISAKI\_NK | False／True／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | ITIJINONYU\_DT\_KOSEI\_TextBox | ITIJINONYU\_DT | False／True／True | name=一次納入日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MAG\_ENCHOKEIYAKUKIGEN\_DT\_KOSEI\_TextBox | MAG\_ENCHOKEIYAKUKIGEN\_DT | False／True／True | name=延長契約期限; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRANDKOSEI\_ID\_KOTAI\_TextBox | BRANDKOSEI\_ID | False／True／True | name=ブランド構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIKOSEI\_ID\_KOTAI\_TextBox | KIKIKOSEI\_ID | False／True／True | name=機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KIKIMEISAI\_ID\_KOTAI\_TextBox | KIKIMEISAI\_ID | False／True／True | name=機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | KOTAIKANRI\_ID\_KOTAI\_TextBox | KOTAIKANRI\_ID | False／True／True | name=個体管理ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | ATSUKAIKIKI\_ID\_KOTAI\_TextBox | ATSUKAIKIKI\_ID | False／True／True | name=取扱機器ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | SET\_NM\_KOTAI\_TextBox | SET\_NM | False／True／True | name=セット数; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRAND\_ID\_KOSEI\_TextBox | BRAND\_ID | False／True／True | name=ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRAND\_NK\_KOTAI\_TextBox | BRAND\_NK | False／True／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | BRANDSYOSAI\_NK\_KOTAI\_TextBox | BRANDSYOSAI\_NK | False／True／True | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIHON\_BRAND\_ID\_KOTAI\_TextBox | UM\_KIHON\_BRAND\_ID | False／True／True | name=店舗基本見積ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIKIKOSEI\_ID\_KOTAI\_TextBox | UM\_KIKIKOSEI\_ID | False／True／True | name=店舗見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KIKIMEISAI\_ID\_KOTAI\_TextBox | UM\_KIKIMEISAI\_ID | False／True／True | name=店舗見積機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UM\_KOTAIMEISAI\_ID\_KOTAI\_TextBox | UM\_KOTAIMEISAI\_ID | False／True／True | name=店舗見積個体明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UK\_KIKIKOSEI\_ID\_KOTAI\_TextBox | UK\_KIKIKOSEI\_ID | False／True／True | name=ユーザー契約機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UK\_KIKIMEISAI\_ID\_KOTAI\_TextBox | UK\_KIKIMEISAI\_ID | False／True／未指定 | name=ユーザー契約機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | UK\_KOTAIMEISAI\_ID\_KOTAI\_TextBox | UK\_KOTAIMEISAI\_ID | False／True／True | name=ユーザー契約個体明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | MITSUMORICHECK\_FLG\_KOTAI\_TextBox | MITSUMORICHECK\_FLG | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_OLD\_KOTAI\_TextBox | CHECK\_FLG\_OLD | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) |
| 2007 | CHECK\_FLG\_BRAND\_CheckBox | CHECK\_FLG | 未指定／False／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BRAND\_NK\_BRAND\_TextBox | BRAND\_NK | False／False／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BRANDSYOSAI\_NK\_BRAND\_TextBox | BRANDSYOSAI\_NK | False／True／True | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KEIYAKUJIKANTAI\_BRAND\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | HOSYUHOHO\_BRAND\_TextBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | DREMOS\_FLG\_BRAND\_CheckBox | DREMOS\_FLG | 未指定／False／True | name=DREMOS契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | REMOTE\_FLG\_BRAND\_CheckBox | REMOTE\_FLG | 未指定／False／True | name=リモート保守フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | HARDHOSYU\_KIN\_BRAND\_TextBox | HARDHOSYU\_KIN | False／False／True | name=ハード保守費用; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | HOSEISOFTHOSHU\_KIN\_BRAND\_TextBox | HOSEISOFTHOSHU\_KIN | False／False／True | name=補正ソフトウェア保守費; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | CHOSEI\_KIN\_BRAND\_TextBox | CHOSEI\_KIN | False／False／True | name=調整費; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | HOSHU\_KIN\_BRAND\_TextBox | HOSHU\_KIN | False／False／True | name=保守費; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BIKO\_BRAND\_TextBox | BIKO | False／False／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_KIHON\_BRAND\_ID\_BRAND\_TextBox | UM\_KIHON\_BRAND\_ID | False／False／True | name=店舗基本見積ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_ID\_BRAND\_TextBox | UM\_MITSUMORI\_ID | False／False／True | name=店舗見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_KIHON\_MITSUMORI\_ID\_BRAND\_TextBox | UM\_KIHON\_MITSUMORI\_ID | False／False／True | name=店舗基本見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BRANDKOSEI\_ID\_BRAND\_TextBox | BRANDKOSEI\_ID | False／False／True | name=ブランド構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | CHECK\_FLG\_MITSUMORI\_CheckBox | CHECK\_FLG | 未指定／False／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_NO\_MITSUMORI\_TextBox | UM\_MITSUMORI\_NO | False／True／True | name=店舗見積NO; type=String; length=20; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | MITSUMORI\_DT\_MITSUMORI\_TextBox | MITSUMORI\_DT | False／False／True | name=見積日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KAISI\_DT\_MITSUMORI\_TextBox | KAISI\_DT | False／True／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | SYURYO\_DT\_SYURYO\_TextBox | SYURYO\_DT | False／False／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | MITSUMORI\_GKIN\_MITSUMORI\_TextBox | MITSUMORI\_GKIN | False／False／True | name=見積合計; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_ID\_MITSUMORI\_TextBox | UM\_MITSUMORI\_ID | False／False／True | name=店舗見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_KIHON\_MITSUMORI\_ID\_MITSUMORI\_TextBox | UM\_KIHON\_MITSUMORI\_ID | False／False／True | name=店舗基本見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BRAND\_NK\_KIKIKOSEI\_Text\_Box | BRAND\_NK | False／False／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BRANDSYOSAI\_NK\_KIKIKOSEI | BRANDSYOSAI\_NK | False／False／True | name=ブランド詳細名; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | TEHAISEIBAN\_KIKIKOSEI\_TextBox | TEHAISEIBAN | False／False／True | name=手配製番; type=String; length=50; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | SEIZOMAKER\_NK\_KIKIMEISAI\_TextBox | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KIKIHINMEI\_NK\_KIKIMEISAI\_TextBox | KIKIHINMEI\_NK | False／True／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KIKIKATASHIKI\_KIKIMEISAI\_TextBox | KIKIKATASHIKI | False／True／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | SURYO\_NM\_KIKIMEISAI\_TextBox | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | HYOJUN\_KIN\_KIKIKOSEI\_TextBox | HYOJUN\_KIN | False／False／True | name=標準価格; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | PACK\_FLG\_KIKIKOSEI\_TextBox | PACK\_FLG | 未指定／False／True | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KEIYAKUNAIYO\_KIKIKOSEI\_TextBox | KEIYAKUNAIYO | False／False／True | name=契約内容; type=String; length=400; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KEIYAKU\_NO\_KIKIKOSEI | KEIYAKU\_NO | False／False／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | DAIFUKUHOSYUJIKAN\_ID\_KIKIKOSEI\_TextBox | DAIFUKUHOSYUJIKAN\_ID | False／False／True | name=ダイフク保守時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | TENKENKAISU\_KIKIKOSEI | TENKENKAISU | False／False／True | name=点検回数; type=Decimal; length=2; min=0; max=12; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | TENKENYOBI\_KIKIKOSEI\_ComboBox | TENKENYOBI | False／False／True | name=点検曜日; type=String; length=1; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | HOSYUHOHO\_KIKIKOSEI\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | SERVICEKEITAI\_KIKIKOSEI\_TextBox | SERVICEKEITAI | False／False／True | name=サービス形態; type=String; length=1; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | BIKO\_KIKIMEISAI\_TextBox | BIKO | False／True／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_MITSUMORI\_ID\_KIKIKOSEI\_TextBox | UM\_MITSUMORI\_ID | False／True／True | name=店舗見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_KIHON\_BRAND\_ID\_KIKIMEISAI\_TextBox | UM\_KIHON\_BRAND\_ID | False／True／True | name=店舗基本見積ブランドID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_KIKIKOSEI\_ID\_KIKIMEISAI\_TextBox | UM\_KIKIKOSEI\_ID | False／True／True | name=店舗見積機器構成ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | UM\_KIKIMEISAI\_ID\_KIKIMEISAI\_TextBox | UM\_KIKIMEISAI\_ID | False／True／True | name=店舗見積機器明細ID; type=Decimal; length=; min=; max=; mask= | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | KAISI\_DTTextBox | KAISI\_DT | True／False／未指定 | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 2007 | SYURYO\_DTTextBox | KAISI\_DT | True／False／未指定 | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox1 | STANDARD\_MONTH | True／True／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_TENPORYAKU\_NK\_VOA\_TextBox | MAR\_TENPORYAKU\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_TANTO\_NK\_VOA\_TextBox | MAR\_TANTO\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_CD\_VOA\_TextBox | NONYUSAKI\_CD | False／True／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_NK\_VOA\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SUPPORT\_ID\_VOA\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PLANT\_NK\_VOA\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_NO\_VOA\_TextBox | KEIYAKU\_NO | False／True／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_DT\_VOA\_TextBox | KEIYAKU\_DT | False／True／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KAIYAKU\_DT\_VOA\_TextBox | KAIYAKU\_DT | False／True／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUMANRYO\_DT\_VOA\_TextBox | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | ENTYOKEIYAKUMANRYO\_DT\_VOA\_TextBox | ENTYOKEIYAKUMANRYO\_DT | False／True／True | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | HOSYUHOHO\_VOA\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUJIKANTAI\_VOA\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox2 | STANDARD\_MONTH | True／True／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TORIHIKISAKI\_CD\_VOB\_TextBox | TORIHIKISAKI\_CD | False／True／True | name=取引先コード; type=String; length=20; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TORIHIKISAKI\_NK\_VOB\_TextBox | TORIHIKISAKI\_NK | False／True／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_CD\_VOB\_TextBox | NONYUSAKI\_CD | False／True／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_NK\_VOB\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SUPPORT\_ID\_VOB\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PLANT\_NK\_VOB\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PACK\_FLG\_VOB\_CheckBox | PACK\_FLG | 未指定／False／True | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_NO\_VOB\_TextBox | KEIYAKU\_NO | False／True／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PACKKEIYAKUNAIYO\_VOB\_TextBox | PACKKEIYAKUNAIYO | False／True／True | name=契約内容; type=String; length=400; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUDTDataGridViewTextBoxColumn | KEIYAKU\_DT | False／True／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KAIYAKU\_DT\_VOB\_TextBox | KAIYAKU\_DT | False／True／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUMANRYO\_DT\_VOB\_TextBox | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | ENTYOKEIYAKUMANRYO\_DT\_VOB\_TextBox | ENTYOKEIYAKUMANRYO\_DT | False／True／True | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | HOSYUHOHO\_VOB\_ComboBox | HOSYUHOHO | False／True／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUJIKANTAI\_VOB\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox3 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TARGET\_PERIODTextBox3 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_HYOJIJUN\_VOC\_TextBox | HYOJIJUN | False／True／True | name=表示順; type=Decimal; length=6; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_TENPORYAKU\_NK\_VOC\_TextBox | TENPORYAKU\_NK | False／True／True | name=店舗略名; type=String; length=100; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_SUPPORT\_ID\_VOC\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_NONYUSAKI\_NK\_VOC\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_NONYUBUSYO\_NK\_VOC\_TextBox | NONYUBUSYO\_NK | False／True／True | name=納入部署; type=String; length=60; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKE\_BRAND\_NK\_VOC\_TextBox | BRAND\_NK | False／True／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKD\_SEIZOMAKER\_NK\_VOC\_TextBox | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKDKIKIHINMEINKDataGridViewTextBoxColumn | KIKIHINMEI\_NK | False／True／True | name=機器品名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKDKIKIKATASHIKIDataGridViewTextBoxColumn | KIKIKATASHIKI | False／True／True | name=機器型式; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SURYONMDataGridViewTextBoxColumn | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TANKAKINDataGridViewTextBoxColumn | TANKA\_KIN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TOTALKINDataGridViewTextBoxColumn | TOTAL\_KIN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKA\_KEIYAKU\_NO\_VOC\_TextBox | TKA\_KEIYAKU\_NO | False／True／True | name=契約NO; type=String; length=20; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUMANRYODTDataGridViewTextBoxColumn1 | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_KIKAN\_VOC\_TextBox | KEIYAKU\_KIKAN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MITSUMORI\_WORK\_VOC\_TextBox | MITSUMORI\_WORK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MITSUMORI\_END\_DT\_VOC\_TextBox | MITSUMORI\_END\_DT | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox4 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TARGET\_PERIODTextBox4 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_HYOJIJUN\_VOD\_TextBox | MAR\_HYOJIJUN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_TENPORYAKU\_NK\_VOD\_TextBox | MAR\_TENPORYAKU\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_SUPPORT\_ID\_VOD\_TextBox | TKB\_SUPPORT\_ID | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_NONYUSAKI\_NK\_VOD\_TextBox | TKB\_NONYUSAKI\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_NONYUBUSYO\_NK\_VOD\_TextBox | TKB\_NONYUBUSYO\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKE\_BRAND\_NK\_VOD\_TextBox | TKE\_BRAND\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKD\_SEIZOMAKER\_NK\_VOD\_TextBox | TKD\_SEIZOMAKER\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKD\_KIKIHINMEI\_NK\_VOD\_TextBox | TKD\_KIKIHINMEI\_NK | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKD\_KIKIKATASHIKI\_VOD\_TextBox | TKD\_KIKIKATASHIKI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SURYO\_NM\_VOD\_TextBox | SURYO\_NM | False／True／True | name=数量; type=Decimal; length=4; min=1; max=9999; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TANKA\_KIN\_VOD\_TextBox | TANKA\_KIN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TOTAL\_KIN\_VOD\_TextBox | TOTAL\_KIN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKA\_KEIYAKU\_NO\_VOD\_TextBox | TKA\_KEIYAKU\_NO | False／True／True | name=契約NO; type=String; length=20; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUMANRYO\_DT\_VOD\_TextBox | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_KIKAN\_VOD\_TextBox | KEIYAKU\_KIKAN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox5 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TARGET\_PERIODTextBox5 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_CD\_VOE\_TextBox | NONYUSAKI\_CD | False／True／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_NK\_VOE\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SUPPORT\_ID\_VOE\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PLANT\_NK\_VOE\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKA\_KEIYAKU\_NO\_VOE\_TextBox | KEIYAKU\_NO | False／True／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKA\_KEIYAKUMANRYO\_DT\_VOE\_TextBox | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKA\_ENTYOKEIYAKUMANRYO\_DT\_VOE\_TextBox | ENTYOKEIYAKUMANRYO\_DT | False／True／True | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_TORIHIKISAKI\_CD\_VOE\_TextBox | TORIHIKISAKI\_CD | False／True／True | name=取引先コード; type=String; length=20; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TKB\_TORIHIKISAKI\_NK\_VOE\_TextBox | TORIHIKISAKI\_NK | False／True／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKA\_KEIYAKU\_NO\_VOE\_TextBox | KEIYAKU\_NO | False／True／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKA\_KEIYAKUMANRYO\_DT\_VOE\_TextBox | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKA\_ENTYOKEIYAKUMANRYO\_DT\_VOE\_TextBox | ENTYOKEIYAKUMANRYO\_DT | False／True／True | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox6 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TARGET\_PERIODTextBox6 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | MAR\_TENPORYAKU\_NK\_VOG\_TextBox | TENPORYAKU\_NK | False／True／True | name=店舗略名; type=String; length=100; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKB\_SUPPORT\_ID\_VOG\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKB\_NONYUSAKI\_NK\_VOG\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKB\_PLANT\_NK\_VOG\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UKB\_HOSYU\_GKIN\_VOG\_TextBox | UKB\_HOSYU\_GKIN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SYURYO\_DT\_VOG\_TextBox | SYURYO\_DT | False／True／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | HARD\_SEIBAN\_VOG\_TextBox | HARD\_SEIBAN | False／True／True | name=ハード製番; type=String; length=20; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox7 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TARGET\_PERIODTextBox7 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TOROKU\_ID\_SAGYOYOTEI\_TextBox | TOROKU\_ID | 未指定／未指定／True | name=登録ID; type=Decimal; length=; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | YOTEI\_DT\_SAGYOYOTEI\_TextBox | YOTEI\_DT | False／False／True | name=予定日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KANRYO\_FLG\_SAGYOYOTEI\_CheckBox | KANRYO\_FLG | 未指定／False／True | name=完了フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NAIYO\_SAGYOYOTEI\_TextBox | NAIYO | False／False／True | name=内容; type=String; length=4000; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | BIKO\_SAGYOYOTEI\_TextBox | BIKO | False／False／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CREATED\_DT\_SAGYOYOTEI\_TextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CREATED\_BY\_SAGYOYOTEI\_TextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | LASTUPDATE\_DT\_SAGYOYOTEI\_TextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | LASTUPDATE\_BY\_SAGYOYOTEI\_TextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | STANDARD\_MONTHTextBox8 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SUPPORT\_ID\_MCM\_ALEAT8\_V\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_NK\_MCM\_ALEAT8\_V\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TENPORYAKU\_NK\_MCM\_ALEAT8\_V\_TextBox | TENPORYAKU\_NK | False／True／True | name=店舗略名; type=String; length=100; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PLANT\_NK\_MCM\_ALEAT8\_V\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | BRAND\_NK\_MCM\_ALEAT8\_V\_TextBox | BRAND\_NK | False／True／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PC\_TENKEN\_MCM\_ALEAT8\_V\_TextBox | PC\_TENKEN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | DENTI\_SYURUI\_MCM\_ALEAT8\_V\_TextBox | DENTI\_SYURUI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SYURUI\_MCM\_ALEAT8\_V\_TextBox | SYURUI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | GOUKI\_MCM\_ALEAT8\_V\_TextBox | GOUKI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SEIZOMAKER\_NK\_MCM\_ALEAT8\_V\_TextBox | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KATASHIKI\_MCM\_ALEAT8\_V\_TextBox | KATASHIKI | False／True／True | name=型式; type=String; length=60; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | BATTERY\_KATASHIKI\_MCM\_ALEAT8\_V\_TextBox | BATTERY\_KATASHIKI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | ITIJINONYU\_DT\_MCM\_ALEAT8\_V\_TextBox | ITIJINONYU\_DT | False／True／True | name=一次納入日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NEXT\_DT\_MCM\_ALEAT8\_V\_TextBox | NEXT\_DT | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | UK\_KEIYAKU\_MCM\_ALEAT8\_V\_TextBox | UK\_KEIYAKU | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TK\_KEIYAKU\_MCM\_ALEAT8\_V\_TextBox | TK\_KEIYAKU | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CT\_KEIYAKU\_MCM\_ALEAT8\_V\_TextBox | CT\_KEIYAKU | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TARGET\_PERIODTextBox8 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KAISI\_DTTextBox | KAISI\_DT | True／True／未指定 | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TORIHIKISAKI\_CD\_VOI\_TextBox | TORIHIKISAKI\_CD | False／True／True | name=取引先コード; type=String; length=20; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | TORIHIKISAKI\_NK\_VOI\_TextBox | TORIHIKISAKI\_NK | False／True／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_CD\_VOＩ\_TextBox | NONYUSAKI\_CD | False／True／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | NONYUSAKI\_NK\_VOI\_TextBox | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SUPPORT\_ID\_VOI\_TextBox | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PLANT\_NK\_VOＩ\_TextBox | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PACK\_FLG\_VOI\_CheckBox | PACK\_FLG | 未指定／False／True | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_NO\_VOＩ\_TextBox | KEIYAKU\_NO | False／True／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | PACKKEIYAKUNAIYO\_VOI\_TextBox | PACKKEIYAKUNAIYO | False／True／True | name=契約内容; type=String; length=400; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | BIKO\_VOI\_TextBox | BIKO | False／False／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKU\_DT\_VOI\_TextBox | KEIYAKU\_DT | False／True／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KAIYAKU\_DT\_VOＩ\_TextBox | KAIYAKU\_DT | False／True／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUMANRYO\_DT\_VOＩ\_TextBox | KEIYAKUMANRYO\_DT | False／True／True | name=契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | ENTYOKEIYAKUMANRYO\_DT\_VOＩ\_TextBox | ENTYOKEIYAKUMANRYO\_DT | False／True／True | name=延長契約満了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KAISI\_DT\_VOＩ\_TextBox | KAISI\_DT | False／True／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SYURYO\_DT\_VOＩ\_TextBox | SYURYO\_DT | False／True／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | HOSYUHOHO\_VOＩ\_ComboBox | HOSYUHOHO | False／True／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | KEIYAKUJIKANTAI\_VOＩ\_TextBox | KEIYAKUJIKANTAI | False／True／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | HYOJUNGOKEI\_KIN\_VOＩ\_TextBox | HYOJUNGOKEI\_KIN | False／True／True | name=標準価格合計; type=Decimal; length=; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SIKIRISYOKEI\_KIN\_VOＩ\_TextBox | SIKIRISYOKEI\_KIN | False／True／True | name=仕切り価格小計; type=Decimal; length=; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SYUSSEINEBIKI\_KIN\_VOＩ\_TextBox | SYUSSEINEBIKI\_KIN | False／True／True | name=出精値引き; type=Decimal; length=; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | SIKIRIGOKEI\_KIN\_VOＩ\_TextBox | SIKIRIGOKEI\_KIN | False／True／True | name=仕切り価格合計; type=Decimal; length=; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpTextBox1 | STANDARD\_MONTH | True／False／未指定 | name=基準月; type=DateTime; length=; min=; max=; mask=yyyy/MM | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn1 | SUPPORT\_ID | False／True／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn2 | NONYUSAKI\_NK | False／True／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn3 | TENPORYAKU\_NK | False／True／True | name=店舗略名; type=String; length=100; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn4 | PLANT\_NK | False／True／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn5 | BRAND\_NK | False／True／True | name=ブランド名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn6 | PC\_TENKEN | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn7 | DENTI\_SYURUI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn8 | SYURUI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn9 | GOUKI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn10 | SEIZOMAKER\_NK | False／True／True | name=製造メーカー名; type=String; length=80; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn11 | KATASHIKI | False／True／True | name=型式; type=String; length=60; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn12 | BATTERY\_KATASHIKI | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn13 | ITIJINONYU\_DT | False／True／True | name=一次納入日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn14 | NEXT\_DT | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn15 | UK\_KEIYAKU | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn16 | TK\_KEIYAKU | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpDataGridTextBoxColumn17 | CT\_KEIYAKU | False／True／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3004 | CpTextBox2 | TARGET\_PERIOD | True／False／未指定 | name=対象期間; type=Decimal; length=2; min=; max=; mask= | [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) |
| 3005 | TOROKU\_IDTextBox | TOROKU\_ID | False／False／True | name=登録ID; type=Decimal; length=; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | YOTEI\_DTTextBox | YOTEI\_DT | True／False／未指定 | name=予定日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | KANRYO\_FLGCheckBox | KANRYO\_FLG | 未指定／未指定／未指定 | name=完了フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | NAIYOTextBox | NAIYO | False／False／未指定 | name=内容; type=String; length=4000; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | BIKOTextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | CREATED\_DTTextBox | CREATED\_DT | False／False／True | name=作成日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | CREATED\_BYTextBox | CREATED\_BY | False／False／True | name=作成者; type=String; length=50; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | LASTUPDATE\_DTTextBox | LASTUPDATE\_DT | False／False／True | name=更新日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | LASTUPDATE\_BYTextBox | LASTUPDATE\_BY | False／False／True | name=更新者; type=String; length=50; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | TENPUFILE\_NK\_TENPU\_LinkBox | TENPUFILE\_NK | 未指定／未指定／未指定 | name=添付ファイル名; type=String; length=2000; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | BIKO\_ID\_TENPU\_TextBox | BIKO | False／False／未指定 | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | DIRECTORY\_TENPU\_TextBox | DIRECTORY | False／False／未指定 | name=ディレクトリ; type=String; length=2000; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | MO\_TENPU\_ID\_TENPU\_TextBox | TENPU\_ID | False／False／未指定 | name=添付ID; type=Decimal; length=; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3005 | MO\_TOROKU\_ID\_TENPU\_TextBox | TOROKU\_ID | False／False／未指定 | name=登録ID; type=Decimal; length=; min=; max=; mask= | [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) |
| 3007 | JOTAI\_2004V\_TextBox | JOTAI | False／False／True | name=状態; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | UM\_MITSUMORI\_NO\_2004V\_Link | UM\_MITSUMORI\_NO | 未指定／未指定／True | name=店舗見積NO; type=String; length=20; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | SYOUNIN\_JOTAI\_2004V\_TextBox | SYOUNIN\_JOTAI | False／False／True | name=承認状態; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_NO\_2004V\_Link | KEIYAKU\_NO | 未指定／未指定／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | SOFUMEISHO4\_NK\_2004V\_TextBox | SOFUMEISHO4\_NK | False／False／True | name=送付先名称４; type=String; length=100; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_CD\_2004V\_TextBox | NONYUSAKI\_CD | False／False／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_NK\_2004V\_TextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KYUNONYUSAKI\_NK\_2004V\_TextBox | KYUNONYUSAKI\_NK | False／False／True | name=旧納入先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKIKOJO\_NK\_2004V\_TextBox | NONYUSAKIKOJO\_NK | False／False／True | name=納入先工場名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | SUPPORT\_ID\_2004V\_TextBox | SUPPORT\_ID | False／False／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PLANT\_NK\_2004V\_TextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKUJIKANTAI\_2004V\_TextBox | KEIYAKUJIKANTAI | False／False／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | HOSYUHOHO\_2004V\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | MITSUMORILEVEL\_2004V\_ComboBox | MITSUMORILEVEL | False／False／True | name=見積レベル; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | MITSUMORI\_DT\_2004V\_TextBox | MITSUMORI\_DT | False／False／True | name=見積日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | MITSUMORISAKUSEISYA\_NK\_2004V\_TextBox | MITSUMORISAKUSEISYA\_NK | False／False／True | name=見積作成者; type=String; length=50; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_DT\_2004V\_TextBox | KEIYAKU\_DT | False／False／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KAIYAKU\_DT\_2004V\_TextBox | KAIYAKU\_DT | False／False／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_KEIYAKU\_2004V\_Link | KEIYAKU\_KEIYAKU | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_DEL\_2004V\_Link | KEIYAKU\_DEL | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PLANT\_ID\_2004V\_TextBox | PLANT\_ID | False／False／True | name=プラントID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_ID\_2004V\_TextBox | NONYUSAKI\_ID | False／False／True | name=納入先ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | UM\_KIHON\_MITSUMORI\_ID\_2004V\_TextBox | UM\_KIHON\_MITSUMORI\_ID | False／False／True | name=店舗基本見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | UK\_KEIYAKU\_ID\_2004V\_TextBox | UK\_KEIYAKU\_ID | False／False／True | name=ユーザー契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TM\_IRAI\_ID\_UVA\_TextBox | TM\_IRAI\_ID | False／False／True | name=取引先見積依頼ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TK\_KEIYAKU\_ID\_UVA\_TextBox | TK\_KEIYAKU\_ID | False／False／True | name=取引先契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | JOTAI\_UVA\_TextBox | JOTAI | False／False／True | name=状態; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TM\_IRAI\_NO\_UVA\_Link | TM\_IRAI\_NO | 未指定／未指定／True | name=取引先見積依頼NO; type=String; length=50; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | MITSUMORI\_DT\_UVA\_TextBox | MITSUMORI\_DT | False／False／True | name=見積日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TM\_KEIYAKUJIKAN\_ID\_UVA\_TextBox | TM\_KEIYAKUJIKAN\_ID | False／False／True | name=取引先見積契約時間ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKUJIKANTAI\_UVA\_TextBox | KEIYAKUJIKANTAI | False／False／True | name=保守契約時間帯; type=Decimal; length=2; min=0; max=24; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | SHONINJOTAI | SHONINJOTAI | False／False／True | name=承認状態; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_NO\_UVA\_Link | KEIYAKU\_NO | 未指定／未指定／True | name=契約番号; type=String; length=50; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PACK\_FLG\_UVA\_ComboBox | PACK\_FLG | False／False／True | name=パック契約フラグ; type=Decimal; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | FUSEI\_UVA\_TextBox | FUSEI | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TORIHIKISAKI\_ID\_UVA\_TextBox | TORIHIKISAKI\_ID | False／False／True | name=取引先ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TORIHIKISAKI\_CD\_UVA\_TextBox | TORIHIKISAKI\_CD | False／False／True | name=取引先コード; type=String; length=20; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TORIHIKISAKI\_NK\_UVA\_TextBox | TORIHIKISAKI\_NK | False／False／True | name=取引先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_ID\_UVA\_TextBox | NONYUSAKI\_ID | False／False／True | name=納入先ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_CD\_UVA\_TextBox | NONYUSAKI\_CD | False／False／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_NK\_UVA\_TextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KYUNONYUSAKI\_NK\_UVA\_TextBox | KYUNONYUSAKI\_NK | False／False／True | name=旧納入先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKIKOJO\_NK\_UVA\_TextBox | NONYUSAKIKOJO\_NK | False／False／True | name=納入先工場名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | HOSYUHOHO\_UVA\_ComboBox | HOSYUHOHO | False／False／True | name=保守方法; type=String; length=2; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TENKENUMU\_UVA\_TextBox | TENKENUMU | False／False／True | name=点検有無; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TENKENKANOYOBI\_UVA\_TextBox | TENKENKANOYOBI | False／False／True | name=点検可能曜日; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | YAKANTAIOUMU\_UVA\_TextBox | YAKANTAIOUMU | False／False／True | name=夜間対応有無; type=String; length=1; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_DT\_UVA\_TextBox | KEIYAKU\_DT | False／False／True | name=契約日付; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KAIYAKU\_DT\_UVA\_TextBox | KAIYAKU\_DT | False／False／True | name=解約日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PLANT\_ID\_UVA\_TextBox | PLANT\_ID | False／False／True | name=プラントID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | SUPPORT\_ID\_UVA\_TextBox | SUPPORT\_ID | False／False／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PLANT\_NK\_UVA\_TextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_KEIYAKU\_UVA\_Link | KEIYAKU\_KEIYAKU | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KEIYAKU\_DEL\_UVA\_Link | KEIYAKU\_DEL | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | UM\_KIHON\_MITSUMORI\_ID\_UVA\_TextBox | UM\_KIHON\_MITSUMORI\_ID | False／False／True | name=店舗基本見積ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_CD\_NONYUSAKI\_Link | NONYUSAKI\_CD | 未指定／未指定／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_NK\_NONYUSAKI\_TextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KYUNONYUSAKI\_NK\_NONYUSAKI\_TextBox | KYUNONYUSAKI\_NK | False／False／True | name=旧納入先名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKIKOJO\_NK\_NONYUSAKI\_TextBox | NONYUSAKIKOJO\_NK | False／False／True | name=納入先工場名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | SUPPORT\_ID\_NONYUSAKI\_Link | SUPPORT\_ID | 未指定／未指定／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PLANT\_NK\_NONYUSAKI\_TextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYU\_DT\_NONYUSAKI\_TextBox | NONYU\_DT | False／False／True | name=納入日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | TEKKYO\_DT\_NONYUSAKI\_TextBox | TEKKYO\_DT | False／False／True | name=撤去日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | HOSYUSYUSOKU\_DT\_NONYUSAKI\_TextBox | HOSYUSYUSOKU\_DT | False／False／True | name=保守終息日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | KIKIKOSEILINK\_NONYUSAKI\_Link | KIKIKOSEILINK | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | NONYUSAKI\_ID\_NONYUSAKI\_TextBox | NONYUSAKI\_ID | False／False／True | name=納入先ID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3007 | PLANT\_ID\_NONYUSAKI\_TextBox | PLANT\_ID | False／False／True | name=プラントID; type=Decimal; length=; min=; max=; mask= | [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) |
| 3008 | CHECK\_FLG\_SEIBAN\_CheckBox | CHECK\_FLG | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KAKUNINKBNDataGridViewTextBoxColumn | KAKUNIN\_KBN | False／False／True | name=確認区分; type=String; length=1; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KAKUNINIRAI\_DT | KAKUNINIRAI\_DT | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KAKUNINIRAI\_BY | KAKUNINIRAI\_BY | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | IRAIMEISHO4\_NK\_SEIBAN\_TextBox | IRAIMEISHO4\_NK | False／False／True | name=依頼元名称４; type=String; length=100; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SUPPORT\_ID\_SEIBAN\_TextBox | SUPPORT\_ID | False／False／True | name=サポートID; type=String; length=7; min=; max=; mask=HANKAKU | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | NONYUSAKI\_NK\_SEIBAN\_TextBox | NONYUSAKI\_NK | False／False／True | name=納入先名; type=String; length=80; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | PLANT\_NK\_SEIBAN\_TextBox | PLANT\_NK | False／False／True | name=プラント名; type=String; length=80; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | HOSYU\_GKIN\_SEIBAN\_TextBox | HOSYU\_GKIN | False／False／True | name=保守金額合計; type=Decimal; length=; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KAISIDT\_SEIBAN\_TextBox | KAISI\_DT | False／False／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SYURYODTDataGridViewTextBoxColumn | SYURYO\_DT | False／False／True | name=終了日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | HARDSEIBANDataGridViewTextBoxColumn | HARD\_SEIBAN | False／False／True | name=ハード製番; type=String; length=20; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SOFTSEIBANDataGridViewTextBoxColumn | SOFT\_SEIBAN | False／False／True | name=ソフト製番; type=String; length=20; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | BIKODataGridViewTextBoxColumn | BIKO | False／False／True | name=備考; type=String; length=4000; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | UKSEIBANIDDataGridViewTextBoxColumn | UK\_SEIBAN\_ID | 未指定／未指定／未指定 | name=ユーザー契約製番ID; type=Decimal; length=; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | UKKEIYAKUIDDataGridViewTextBoxColumn | UK\_KEIYAKU\_ID | 未指定／未指定／未指定 | name=ユーザー契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | CHECK\_FLG\_MITSUMORI\_TextBox | CHECK\_FLG | 未指定／True／未指定 | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SHORUI\_NO\_MITSUMORI\_Link | SHORUI\_NO | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KAISI\_DT\_MITSUMORI\_TextBox | KAISI\_DT | False／False／True | name=開始日; type=DateTime; length=; min=; max=; mask=yyyy/MM/dd | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KINGAKU\_MITSUMORI\_TextBox | KINGAKU | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | IRAITANTO\_NK\_MITSUMORI\_TextBox | IRAITANTO\_NK | False／False／True | name=依頼元担当者; type=String; length=100; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | JOTAI\_MITSUMORI\_ComboBox | JOTAI | False／False／True | name=状態; type=String; length=1; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | RELATION\_ID\_MITSUMORI\_TextBox | RELATION\_ID | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | KIKIKOSEILINK\_MITSUMORI\_Link | KIKIKOSEILINK | 未指定／未指定／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SHINSA\_DT | SHINSA\_DT | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SHINSA\_BY | SHINSA\_BY | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SYOUNIN\_DT | SYOUNIN\_DT | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | SYOUNIN\_BY | SYOUNIN\_BY | False／False／True | 定義なし（監査・対象外型・実行時検査キーを確認） | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | TK\_KEIYAKU\_ID\_MITSUMORI\_TextBox | TK\_KEIYAKU\_ID | False／False／True | name=取引先契約ID; type=Decimal; length=; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | NONYUSAKI\_CD\_MITSUMORI\_TextBox | NONYUSAKI\_CD | 未指定／未指定／True | name=納入先コード; type=String; length=12; min=; max=; mask=HANKAKU | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |
| 3008 | PLANT\_ID\_MITSUMORI\_TextBox | PLANT\_ID | False／False／True | name=プラントID; type=Decimal; length=; min=; max=; mask= | [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) |

## DataSetの型・NULL・桁数・キーと関連

XSDのテーブル・列とキー／関連の定義を記録します。これはVBのDataSet定義であり、現在DBの物理制約ではありません。既定minOccurs=1、明示0は省略可能。SQL更新・削除WHEREは既存SQL原文で確認してください。DataSet制約による例外や親子削除・採番の競合は実行確認が必要です。

### 1005　Mcm1005uDataSet.xsd

[Mcm1005uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_TK\_KEIYAKU | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | IRAIJIGYOSYO\_NK | {"Generator\_UserColumnName": "IRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "IRAIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "IRAIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KEIYAKU | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUKAISI\_DT | {"Generator\_UserColumnName": "KEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUSYURYO\_DT | {"Generator\_UserColumnName": "KEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUSYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUSYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnJIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "nullValue": "\_throw", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "Generator\_ColumnVarNameInTable": "columnJOTAI", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KEIYAKU | JIDOKOSIN\_FLG | {"Generator\_UserColumnName": "JIDOKOSIN\_FLG", "Generator\_ColumnPropNameInRow": "JIDOKOSIN\_FLG", "Generator\_ColumnVarNameInTable": "columnJIDOKOSIN\_FLG", "Generator\_ColumnPropNameInTable": "JIDOKOSIN\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKU\_NO | {"Generator\_UserColumnName": "PACKKEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | SHOKAI\_KEIYAKU\_DT | {"Generator\_UserColumnName": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnSHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "SHOKAI\_KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KOSINNAIYO | {"Generator\_UserColumnName": "KOSINNAIYO", "Generator\_ColumnVarNameInTable": "columnKOSINNAIYO", "Generator\_ColumnPropNameInRow": "KOSINNAIYO", "Generator\_ColumnPropNameInTable": "KOSINNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_TK\_KIKAN | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKAN | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | HYOJUNGOKEI\_KIN | {"Generator\_UserColumnName": "HYOJUNGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUNGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SIKIRISYOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRISYOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRISYOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SYUSSEINEBIKI\_KIN | {"Generator\_UserColumnName": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInRow": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInTable": "SYUSSEINEBIKI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SIKIRIGOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRIGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRIGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_TK\_KIKAN | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_TK\_KIKAN | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KIKAN | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KIKAN | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TK\_KIKAN | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | TORITEL\_NO | {"Generator\_UserColumnName": "TORITEL\_NO", "Generator\_ColumnPropNameInRow": "TORITEL\_NO", "Generator\_ColumnVarNameInTable": "columnTORITEL\_NO", "Generator\_ColumnPropNameInTable": "TORITEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | TORIFAX\_NO | {"Generator\_UserColumnName": "TORIFAX\_NO", "Generator\_ColumnPropNameInRow": "TORIFAX\_NO", "Generator\_ColumnVarNameInTable": "columnTORIFAX\_NO", "Generator\_ColumnPropNameInTable": "TORIFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | TORIJIGYOSYO\_NK | {"Generator\_UserColumnName": "TORIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "TORIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnTORIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "TORIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KIKAN | TORISYUTANTOSYA\_NK | {"Generator\_UserColumnName": "TORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "TORISYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnTORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "TORISYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KIKAN | TORIASSISTANT\_NK | {"Generator\_UserColumnName": "TORIASSISTANT\_NK", "Generator\_ColumnPropNameInRow": "TORIASSISTANT\_NK", "Generator\_ColumnVarNameInTable": "columnTORIASSISTANT\_NK", "Generator\_ColumnPropNameInTable": "TORIASSISTANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TK\_KIKAN | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TK\_KIKAN | YUKO\_FLG | {"Generator\_UserColumnName": "YUKO\_FLG", "Generator\_ColumnVarNameInTable": "columnYUKO\_FLG", "Generator\_ColumnPropNameInRow": "YUKO\_FLG", "Generator\_ColumnPropNameInTable": "YUKO\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKAN | YUBIN\_NO | {"Generator\_UserColumnName": "YUBIN\_NO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "YUBIN\_NO", "Generator\_ColumnPropNameInTable": "YUBIN\_NOColumn", "Generator\_ColumnVarNameInTable": "columnYUBIN\_NO", "minOccurs": "0"} | [["maxLength", {"value": "8"}]] |
| MCM\_TK\_KIKAN | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIJOHO | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKIJOHO | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIJOHO | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKIJOHO | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKIJOHO | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKIJOHO | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KIKIJOHO | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TK\_KIKIJOHO | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KIKIJOHO | HYOJUN\_KIN | {"Generator\_UserColumnName": "HYOJUN\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUN\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUN\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUN\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKIJOHO | TKB\_SUPPORT\_ID\_OLD | {"Generator\_UserColumnName": "TKB\_SUPPORT\_ID\_OLD", "Generator\_ColumnPropNameInRow": "TKB\_SUPPORT\_ID\_OLD", "Generator\_ColumnVarNameInTable": "columnTKB\_SUPPORT\_ID\_OLD", "Generator\_ColumnPropNameInTable": "TKB\_SUPPORT\_ID\_OLDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_TK\_KIKIJOHO | MAB\_SUPPORT\_ID | {"Generator\_UserColumnName": "MAB\_SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "MAB\_SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnMAB\_SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "MAB\_SUPPORT\_IDColumn"} | [["maxLength", {"value": "7"}]] |
| MCM\_TK\_KIKIJOHO | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKIJOHO | MAF\_HYOJIJUN | {"Generator\_UserColumnName": "MAF\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAF\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAF\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAF\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | SIKIRI\_KIN | {"Generator\_UserColumnName": "SIKIRI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_TK\_KIKIJOHO | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | HYOJUNKEI\_KIN | {"Generator\_UserColumnName": "HYOJUNKEI\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUNKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUNKEI\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUNKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIJOHO | SIKIRIKEI\_KIN | {"Generator\_UserColumnName": "SIKIRIKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRIKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRIKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRIKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAI | TK\_SIHARAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAI | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_SIHARAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAI | KAISU | {"Generator\_UserColumnName": "KAISU", "Generator\_ColumnPropNameInRow": "KAISU", "Generator\_ColumnVarNameInTable": "columnKAISU", "Generator\_ColumnPropNameInTable": "KAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKEN\_ID | {"Generator\_UserColumnName": "TK\_TENKEN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TENKEN\_ID", "Generator\_ColumnPropNameInRow": "TK\_TENKEN\_ID", "Generator\_ColumnPropNameInTable": "TK\_TENKEN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENKEN | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENKEN | TENKENKAISU | {"Generator\_UserColumnName": "TENKENKAISU", "Generator\_ColumnVarNameInTable": "columnTENKENKAISU", "Generator\_ColumnPropNameInRow": "TENKENKAISU", "Generator\_ColumnPropNameInTable": "TENKENKAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TENKENKANOYOBI | {"Generator\_UserColumnName": "TENKENKANOYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENKANOYOBI", "Generator\_ColumnPropNameInRow": "TENKENKANOYOBI", "Generator\_ColumnPropNameInTable": "TENKENKANOYOBIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_TENKEN | YAKANTAIOUMU | {"Generator\_UserColumnName": "YAKANTAIOUMU", "Generator\_ColumnVarNameInTable": "columnYAKANTAIOUMU", "Generator\_ColumnPropNameInRow": "YAKANTAIOUMU", "Generator\_ColumnPropNameInTable": "YAKANTAIOUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_TENKEN | HOSHUGAISHA\_ID | {"Generator\_UserColumnName": "HOSHUGAISHA\_ID", "Generator\_ColumnVarNameInTable": "columnHOSHUGAISHA\_ID", "Generator\_ColumnPropNameInRow": "HOSHUGAISHA\_ID", "Generator\_ColumnPropNameInTable": "HOSHUGAISHA\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_TENKEN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKEN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID1 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID1", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID1", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID1", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID1Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID2 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID2", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID2", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID2", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID2Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID3 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID3", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID3", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID3", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID3Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID4 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID4", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID4", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID4", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID4Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID5 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID5", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID5", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID5", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID5Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID6 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID6", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID6", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID6", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID6Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID7 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID7", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID7", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID7", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID7Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID8 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID8", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID8", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID8", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID8Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID9 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID9", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID9", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID9", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID9Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID10 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID10", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID10", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID10", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID10Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID11 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID11", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID11", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID11", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID11Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_TENKENMEISAI\_ID12 | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID12", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID12", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID12", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_ID12Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI1 | {"Generator\_UserColumnName": "TSUKI1", "Generator\_ColumnPropNameInRow": "TSUKI1", "Generator\_ColumnVarNameInTable": "columnTSUKI1", "Generator\_ColumnPropNameInTable": "TSUKI1Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI2 | {"Generator\_UserColumnName": "TSUKI2", "Generator\_ColumnPropNameInRow": "TSUKI2", "Generator\_ColumnVarNameInTable": "columnTSUKI2", "Generator\_ColumnPropNameInTable": "TSUKI2Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI3 | {"Generator\_UserColumnName": "TSUKI3", "Generator\_ColumnPropNameInRow": "TSUKI3", "Generator\_ColumnVarNameInTable": "columnTSUKI3", "Generator\_ColumnPropNameInTable": "TSUKI3Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI4 | {"Generator\_UserColumnName": "TSUKI4", "Generator\_ColumnPropNameInRow": "TSUKI4", "Generator\_ColumnVarNameInTable": "columnTSUKI4", "Generator\_ColumnPropNameInTable": "TSUKI4Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI5 | {"Generator\_UserColumnName": "TSUKI5", "Generator\_ColumnPropNameInRow": "TSUKI5", "Generator\_ColumnVarNameInTable": "columnTSUKI5", "Generator\_ColumnPropNameInTable": "TSUKI5Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI6 | {"Generator\_UserColumnName": "TSUKI6", "Generator\_ColumnPropNameInRow": "TSUKI6", "Generator\_ColumnVarNameInTable": "columnTSUKI6", "Generator\_ColumnPropNameInTable": "TSUKI6Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI7 | {"Generator\_UserColumnName": "TSUKI7", "Generator\_ColumnPropNameInRow": "TSUKI7", "Generator\_ColumnVarNameInTable": "columnTSUKI7", "Generator\_ColumnPropNameInTable": "TSUKI7Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI8 | {"Generator\_UserColumnName": "TSUKI8", "Generator\_ColumnPropNameInRow": "TSUKI8", "Generator\_ColumnVarNameInTable": "columnTSUKI8", "Generator\_ColumnPropNameInTable": "TSUKI8Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI9 | {"Generator\_UserColumnName": "TSUKI9", "Generator\_ColumnPropNameInRow": "TSUKI9", "Generator\_ColumnVarNameInTable": "columnTSUKI9", "Generator\_ColumnPropNameInTable": "TSUKI9Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI10 | {"Generator\_UserColumnName": "TSUKI10", "Generator\_ColumnPropNameInRow": "TSUKI10", "Generator\_ColumnVarNameInTable": "columnTSUKI10", "Generator\_ColumnPropNameInTable": "TSUKI10Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI11 | {"Generator\_UserColumnName": "TSUKI11", "Generator\_ColumnPropNameInRow": "TSUKI11", "Generator\_ColumnVarNameInTable": "columnTSUKI11", "Generator\_ColumnPropNameInTable": "TSUKI11Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TSUKI12 | {"Generator\_UserColumnName": "TSUKI12", "Generator\_ColumnPropNameInRow": "TSUKI12", "Generator\_ColumnVarNameInTable": "columnTSUKI12", "Generator\_ColumnPropNameInTable": "TSUKI12Column", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn"} | [["maxLength", {"value": "100"}]] |
| MCM\_TK\_TENKEN | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENKEN | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENKEN | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN | YAKANHOSHUGAISHA\_ID | {"Generator\_UserColumnName": "YAKANHOSHUGAISHA\_ID", "Generator\_ColumnPropNameInRow": "YAKANHOSHUGAISHA\_ID", "Generator\_ColumnVarNameInTable": "columnYAKANHOSHUGAISHA\_ID", "Generator\_ColumnPropNameInTable": "YAKANHOSHUGAISHA\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN | OYAKIKIBUNRUI\_CD | {"Generator\_UserColumnName": "OYAKIKIBUNRUI\_CD", "Generator\_ColumnPropNameInRow": "OYAKIKIBUNRUI\_CD", "Generator\_ColumnVarNameInTable": "columnOYAKIKIBUNRUI\_CD", "Generator\_ColumnPropNameInTable": "OYAKIKIBUNRUI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TK\_TENKEN | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENKEN | MAF\_HYOJIJUN | {"Generator\_UserColumnName": "MAF\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAF\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAF\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAF\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENPU | TK\_TENPU\_ID | {"Generator\_UserColumnName": "TK\_TENPU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TENPU\_ID", "Generator\_ColumnPropNameInRow": "TK\_TENPU\_ID", "Generator\_ColumnPropNameInTable": "TK\_TENPU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENPU | TENPUFILE\_NK | {"Generator\_UserColumnName": "TENPUFILE\_NK", "Generator\_ColumnVarNameInTable": "columnTENPUFILE\_NK", "Generator\_ColumnPropNameInRow": "TENPUFILE\_NK", "Generator\_ColumnPropNameInTable": "TENPUFILE\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_TK\_TENPU | DIRECTORY | {"Generator\_UserColumnName": "DIRECTORY", "Generator\_ColumnVarNameInTable": "columnDIRECTORY", "Generator\_ColumnPropNameInRow": "DIRECTORY", "Generator\_ColumnPropNameInTable": "DIRECTORYColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_TK\_TENPU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_TENPU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | LOCAL\_PATH | {"Generator\_UserColumnName": "LOCAL\_PATH", "Generator\_ColumnVarNameInTable": "columnLOCAL\_PATH", "Generator\_ColumnPropNameInRow": "LOCAL\_PATH", "Generator\_ColumnPropNameInTable": "LOCAL\_PATHColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_DT | {"Generator\_UserColumnName": "ON\_DT", "Generator\_ColumnPropNameInRow": "ON\_DT", "Generator\_ColumnVarNameInTable": "columnON\_DT", "Generator\_ColumnPropNameInTable": "ON\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_BY | {"Generator\_UserColumnName": "ON\_BY", "Generator\_ColumnPropNameInRow": "ON\_BY", "Generator\_ColumnVarNameInTable": "columnON\_BY", "Generator\_ColumnPropNameInTable": "ON\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_DT | {"Generator\_UserColumnName": "OFF\_DT", "Generator\_ColumnPropNameInRow": "OFF\_DT", "Generator\_ColumnVarNameInTable": "columnOFF\_DT", "Generator\_ColumnPropNameInTable": "OFF\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_BY | {"Generator\_UserColumnName": "OFF\_BY", "Generator\_ColumnPropNameInRow": "OFF\_BY", "Generator\_ColumnVarNameInTable": "columnOFF\_BY", "Generator\_ColumnPropNameInTable": "OFF\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | ON\_FLG | {"Generator\_UserColumnName": "ON\_FLG", "Generator\_ColumnPropNameInRow": "ON\_FLG", "Generator\_ColumnVarNameInTable": "columnON\_FLG", "Generator\_ColumnPropNameInTable": "ON\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_FLG | {"Generator\_UserColumnName": "OFF\_FLG", "Generator\_ColumnPropNameInRow": "OFF\_FLG", "Generator\_ColumnVarNameInTable": "columnOFF\_FLG", "Generator\_ColumnPropNameInTable": "OFF\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TSUKI | {"Generator\_UserColumnName": "TSUKI", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnVarNameInTable": "columnTSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TSUKI\_MM | {"Generator\_UserColumnName": "TSUKI\_MM", "Generator\_ColumnPropNameInRow": "TSUKI\_MM", "Generator\_ColumnVarNameInTable": "columnTSUKI\_MM", "Generator\_ColumnPropNameInTable": "TSUKI\_MMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | TK\_TENKEN\_ID | {"Generator\_UserColumnName": "TK\_TENKEN\_ID", "Generator\_ColumnPropNameInRow": "TK\_TENKEN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TENKEN\_ID", "Generator\_ColumnPropNameInTable": "TK\_TENKEN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKENMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKENMEISAI | TK\_TENKENMEISAI\_ID | {"Generator\_UserColumnName": "TK\_TENKENMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TENKENMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_TENKENMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_TENKENMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENKENMEISAI | ON\_FLG | {"Generator\_UserColumnName": "ON\_FLG", "Generator\_ColumnVarNameInTable": "columnON\_FLG", "Generator\_ColumnPropNameInRow": "ON\_FLG", "Generator\_ColumnPropNameInTable": "ON\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | ON\_DT | {"Generator\_UserColumnName": "ON\_DT", "Generator\_ColumnVarNameInTable": "columnON\_DT", "Generator\_ColumnPropNameInRow": "ON\_DT", "Generator\_ColumnPropNameInTable": "ON\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | ON\_BY | {"Generator\_UserColumnName": "ON\_BY", "Generator\_ColumnVarNameInTable": "columnON\_BY", "Generator\_ColumnPropNameInRow": "ON\_BY", "Generator\_ColumnPropNameInTable": "ON\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKENMEISAI | OFF\_FLG | {"Generator\_UserColumnName": "OFF\_FLG", "Generator\_ColumnVarNameInTable": "columnOFF\_FLG", "Generator\_ColumnPropNameInRow": "OFF\_FLG", "Generator\_ColumnPropNameInTable": "OFF\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | OFF\_DT | {"Generator\_UserColumnName": "OFF\_DT", "Generator\_ColumnVarNameInTable": "columnOFF\_DT", "Generator\_ColumnPropNameInRow": "OFF\_DT", "Generator\_ColumnPropNameInTable": "OFF\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | OFF\_BY | {"Generator\_UserColumnName": "OFF\_BY", "Generator\_ColumnVarNameInTable": "columnOFF\_BY", "Generator\_ColumnPropNameInRow": "OFF\_BY", "Generator\_ColumnPropNameInTable": "OFF\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKENMEISAI | TSUKI | {"Generator\_UserColumnName": "TSUKI", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnVarNameInTable": "columnTSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKENMEISAI | TSUKI\_MM | {"Generator\_UserColumnName": "TSUKI\_MM", "Generator\_ColumnPropNameInRow": "TSUKI\_MM", "Generator\_ColumnVarNameInTable": "columnTSUKI\_MM", "Generator\_ColumnPropNameInTable": "TSUKI\_MMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIKOSEI | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKIKOSEI | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIKOSEI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIKOSEI | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIKOSEI | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIKOSEI | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnVarNameInTable": "columnBIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KIKIKOSEI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIKOSEI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIKOSEI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIKOSEI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIKOSEI | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_TK\_KIKIKOSEI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKIMEISAI | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKIMEISAI | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | SEIZOMAKER\_ID | {"Generator\_UserColumnName": "SEIZOMAKER\_ID", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_ID", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_ID", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKIMEISAI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKIMEISAI | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnVarNameInTable": "columnBIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KIKIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKIMEISAI | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKIMEISAI | OYAKIKIBUNRUI\_CD | {"Generator\_UserColumnName": "OYAKIKIBUNRUI\_CD", "Generator\_ColumnVarNameInTable": "columnOYAKIKIBUNRUI\_CD", "Generator\_ColumnPropNameInRow": "OYAKIKIBUNRUI\_CD", "Generator\_ColumnPropNameInTable": "OYAKIKIBUNRUI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TK\_KOTAIMEISAI | TK\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KOTAIMEISAI | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | KOTAIKANRI\_ID | {"Generator\_UserColumnName": "KOTAIKANRI\_ID", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_ID", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_ID", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | KOTAI\_NK | {"Generator\_UserColumnName": "KOTAI\_NK", "Generator\_ColumnVarNameInTable": "columnKOTAI\_NK", "Generator\_ColumnPropNameInRow": "KOTAI\_NK", "Generator\_ColumnPropNameInTable": "KOTAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KOTAIMEISAI | SERIAL\_NO | {"Generator\_UserColumnName": "SERIAL\_NO", "Generator\_ColumnVarNameInTable": "columnSERIAL\_NO", "Generator\_ColumnPropNameInRow": "SERIAL\_NO", "Generator\_ColumnPropNameInTable": "SERIAL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_TK\_KOTAIMEISAI | ITIJINONYU\_DT | {"Generator\_UserColumnName": "ITIJINONYU\_DT", "Generator\_ColumnVarNameInTable": "columnITIJINONYU\_DT", "Generator\_ColumnPropNameInRow": "ITIJINONYU\_DT", "Generator\_ColumnPropNameInTable": "ITIJINONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | SETCHIBASYO | {"Generator\_UserColumnName": "SETCHIBASYO", "Generator\_ColumnVarNameInTable": "columnSETCHIBASYO", "Generator\_ColumnPropNameInRow": "SETCHIBASYO", "Generator\_ColumnPropNameInTable": "SETCHIBASYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KOTAIMEISAI | TEKKYOBI\_DT | {"Generator\_UserColumnName": "TEKKYOBI\_DT", "Generator\_ColumnVarNameInTable": "columnTEKKYOBI\_DT", "Generator\_ColumnPropNameInRow": "TEKKYOBI\_DT", "Generator\_ColumnPropNameInTable": "TEKKYOBI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | KEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "KEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KOTAIMEISAI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KOTAIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KOTAIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KOTAIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KOTAIMEISAI | ENCHOKEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "ENCHOKEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KOTAIMEISAI | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TANKA | TK\_TANKA\_ID | {"Generator\_UserColumnName": "TK\_TANKA\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TANKA\_ID", "Generator\_ColumnPropNameInRow": "TK\_TANKA\_ID", "Generator\_ColumnPropNameInTable": "TK\_TANKA\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TANKA | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | TM\_TANKA\_ID | {"Generator\_UserColumnName": "TM\_TANKA\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_TANKA\_ID", "Generator\_ColumnPropNameInRow": "TM\_TANKA\_ID", "Generator\_ColumnPropNameInTable": "TM\_TANKA\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | KEIYAKUNAIYO | {"Generator\_UserColumnName": "KEIYAKUNAIYO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "KEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "KEIYAKUNAIYOColumn", "Generator\_ColumnVarNameInTable": "columnKEIYAKUNAIYO", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TK\_TANKA | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TANKA | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_TANKA | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TK\_TANKA | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_TANKA | HYOJUN\_KIN | {"Generator\_UserColumnName": "HYOJUN\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUN\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUN\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUN\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_TANKA | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TANKA | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TANKA | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | SIKIRI\_KIN | {"Generator\_UserColumnName": "SIKIRI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | TM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TANKA | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TANKA | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | HYOJUNKEI\_KIN | {"Generator\_UserColumnName": "HYOJUNKEI\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUNKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUNKEI\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUNKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | SIKIRIKEI\_KIN | {"Generator\_UserColumnName": "SIKIRIKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRIKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRIKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRIKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TANKA | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_PLANT | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_PLANT | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_PLANT | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn"} | [["maxLength", {"value": "7"}]] |
| MCM\_MA\_PLANT | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_PLANT | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_PLANT | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_MA\_PLANT | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_PLANT | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_PLANT | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn"} | [["maxLength", {"value": "12"}]] |
| MCM\_MA\_PLANT | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_PLANT | JUSYO1\_NK | {"Generator\_UserColumnName": "JUSYO1\_NK", "Generator\_ColumnPropNameInRow": "JUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "JUSYO1\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_PLANT | JUSYO2\_NK | {"Generator\_UserColumnName": "JUSYO2\_NK", "Generator\_ColumnPropNameInRow": "JUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "JUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_KEIYAKUJIKAN | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_KEIYAKUJIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_KIKIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KOTAIKANRI\_ID | {"Generator\_UserColumnName": "KOTAIKANRI\_ID", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_ID", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_ID", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | ATSUKAIKIKIKOSEI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KOTAI\_NK | {"Generator\_UserColumnName": "KOTAI\_NK", "Generator\_ColumnVarNameInTable": "columnKOTAI\_NK", "Generator\_ColumnPropNameInRow": "KOTAI\_NK", "Generator\_ColumnPropNameInTable": "KOTAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | SERIAL\_NO | {"Generator\_UserColumnName": "SERIAL\_NO", "Generator\_ColumnVarNameInTable": "columnSERIAL\_NO", "Generator\_ColumnPropNameInRow": "SERIAL\_NO", "Generator\_ColumnPropNameInTable": "SERIAL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | ITIJINONYU\_DT | {"Generator\_UserColumnName": "ITIJINONYU\_DT", "Generator\_ColumnVarNameInTable": "columnITIJINONYU\_DT", "Generator\_ColumnPropNameInRow": "ITIJINONYU\_DT", "Generator\_ColumnPropNameInTable": "ITIJINONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | SETCHIBASYO | {"Generator\_UserColumnName": "SETCHIBASYO", "Generator\_ColumnVarNameInTable": "columnSETCHIBASYO", "Generator\_ColumnPropNameInRow": "SETCHIBASYO", "Generator\_ColumnPropNameInTable": "SETCHIBASYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | TEKKYO\_DT | {"Generator\_UserColumnName": "TEKKYO\_DT", "Generator\_ColumnVarNameInTable": "columnTEKKYO\_DT", "Generator\_ColumnPropNameInRow": "TEKKYO\_DT", "Generator\_ColumnPropNameInTable": "TEKKYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "KEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | ENCHOKEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "ENCHOKEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UPSKOKAN\_DT | {"Generator\_UserColumnName": "UPSKOKAN\_DT", "Generator\_ColumnVarNameInTable": "columnUPSKOKAN\_DT", "Generator\_ColumnPropNameInRow": "UPSKOKAN\_DT", "Generator\_ColumnPropNameInTable": "UPSKOKAN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KEIYAKUMANRYOYOTEI\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYOYOTEI\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYOYOTEI\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYOYOTEI\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYOYOTEI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KEIYAKU" />
      <xs:field xpath="mstns:TK_KEIYAKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KIKAN" />
      <xs:field xpath="mstns:TK_KIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KIKIJOHO_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KIKIJOHO" />
      <xs:field xpath="mstns:TK_KIKIKOSEI_ID" />
      <xs:field xpath="mstns:TK_KIKIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_SIHARAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_SIHARAI" />
      <xs:field xpath="mstns:TK_SIHARAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_TENKEN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_TENKEN" />
      <xs:field xpath="mstns:TK_TENKEN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_TENPU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_TENPU" />
      <xs:field xpath="mstns:TK_TENPU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_SIHARAIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_SIHARAIMEISAI" />
      <xs:field xpath="mstns:TK_SIHARAIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_TENKENMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_TENKENMEISAI" />
      <xs:field xpath="mstns:TK_TENKENMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KIKIKOSEI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KIKIKOSEI" />
      <xs:field xpath="mstns:TK_KIKIKOSEI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KIKIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KIKIMEISAI" />
      <xs:field xpath="mstns:TK_KIKIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KOTAIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KOTAIMEISAI" />
      <xs:field xpath="mstns:TK_KOTAIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_TANKA_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_TANKA" />
      <xs:field xpath="mstns:TK_TANKA_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_PLANT_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_PLANT" />
      <xs:field xpath="mstns:PLANT_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_KEIYAKUJIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_KEIYAKUJIKAN" />
      <xs:field xpath="mstns:TM_KEIYAKUJIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_KIKIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_KIKIMEISAI" />
      <xs:field xpath="mstns:KIKIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_KIKIKOTAIKANRI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_KIKIKOTAIKANRI" />
      <xs:field xpath="mstns:KOTAIKANRI_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_TENKEN_MCM_TK_TENKENMEISAI" ns0:parent="MCM_TK_TENKEN" ns0:child="MCM_TK_TENKENMEISAI" ns0:parentkey="TK_TENKEN_ID" ns0:childkey="TK_TENKEN_ID" ns1:Generator_UserRelationName="MCM_TK_TENKEN_MCM_TK_TENKENMEISAI" ns1:Generator_RelationVarName="relationMCM_TK_TENKEN_MCM_TK_TENKENMEISAI" ns1:Generator_UserChildTable="MCM_TK_TENKENMEISAI" ns1:Generator_UserParentTable="MCM_TK_TENKEN" ns1:Generator_ParentPropName="MCM_TK_TENKENRow" ns1:Generator_ChildPropName="GetMCM_TK_TENKENMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_SIHARAI_MCM_TK_SIHARAIMEISAI" ns0:parent="MCM_TK_SIHARAI" ns0:child="MCM_TK_SIHARAIMEISAI" ns0:parentkey="TK_SIHARAI_ID" ns0:childkey="TK_SIHARAI_ID" ns1:Generator_UserRelationName="MCM_TK_SIHARAI_MCM_TK_SIHARAIMEISAI" ns1:Generator_RelationVarName="relationMCM_TK_SIHARAI_MCM_TK_SIHARAIMEISAI" ns1:Generator_UserChildTable="MCM_TK_SIHARAIMEISAI" ns1:Generator_UserParentTable="MCM_TK_SIHARAI" ns1:Generator_ParentPropName="MCM_TK_SIHARAIRow" ns1:Generator_ChildPropName="GetMCM_TK_SIHARAIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KEIYAKU_MCM_TK_KIKAN" ns0:parent="MCM_TK_KEIYAKU" ns0:child="MCM_TK_KIKAN" ns0:parentkey="TK_KEIYAKU_ID" ns0:childkey="TK_KEIYAKU_ID" ns1:Generator_UserRelationName="MCM_TK_KEIYAKU_MCM_TK_KIKAN" ns1:Generator_RelationVarName="relationMCM_TK_KEIYAKU_MCM_TK_KIKAN" ns1:Generator_UserChildTable="MCM_TK_KIKAN" ns1:Generator_UserParentTable="MCM_TK_KEIYAKU" ns1:Generator_ParentPropName="MCM_TK_KEIYAKURow" ns1:Generator_ChildPropName="GetMCM_TK_KIKANRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKAN_MCM_TK_KIKIJOHO" ns0:parent="MCM_TK_KIKAN" ns0:child="MCM_TK_KIKIJOHO" ns0:parentkey="TK_KIKAN_ID" ns0:childkey="TK_KIKAN_ID" ns1:Generator_UserRelationName="MCM_TK_KIKAN_MCM_TK_KIKIJOHO" ns1:Generator_RelationVarName="relationMCM_TK_KIKAN_MCM_TK_KIKIJOHO" ns1:Generator_UserChildTable="MCM_TK_KIKIJOHO" ns1:Generator_UserParentTable="MCM_TK_KIKAN" ns1:Generator_ParentPropName="MCM_TK_KIKANRow" ns1:Generator_ChildPropName="GetMCM_TK_KIKIJOHORows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKAN_MCM_TK_SIHARAI" ns0:parent="MCM_TK_KIKAN" ns0:child="MCM_TK_SIHARAI" ns0:parentkey="TK_KIKAN_ID" ns0:childkey="TK_KIKAN_ID" ns1:Generator_UserRelationName="MCM_TK_KIKAN_MCM_TK_SIHARAI" ns1:Generator_RelationVarName="relationMCM_TK_KIKAN_MCM_TK_SIHARAI" ns1:Generator_UserChildTable="MCM_TK_SIHARAI" ns1:Generator_UserParentTable="MCM_TK_KIKAN" ns1:Generator_ParentPropName="MCM_TK_KIKANRow" ns1:Generator_ChildPropName="GetMCM_TK_SIHARAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKAN_MCM_TK_KIKIKOSEI" ns0:parent="MCM_TK_KIKAN" ns0:child="MCM_TK_KIKIKOSEI" ns0:parentkey="TK_KIKAN_ID" ns0:childkey="TK_KIKAN_ID" ns1:Generator_UserRelationName="MCM_TK_KIKAN_MCM_TK_KIKIKOSEI" ns1:Generator_RelationVarName="relationMCM_TK_KIKAN_MCM_TK_KIKIKOSEI" ns1:Generator_UserChildTable="MCM_TK_KIKIKOSEI" ns1:Generator_UserParentTable="MCM_TK_KIKAN" ns1:Generator_ParentPropName="MCM_TK_KIKANRow" ns1:Generator_ChildPropName="GetMCM_TK_KIKIKOSEIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKIKOSEI_MCM_TK_KIKIMEISAI" ns0:parent="MCM_TK_KIKIKOSEI" ns0:child="MCM_TK_KIKIMEISAI" ns0:parentkey="TK_KIKIKOSEI_ID" ns0:childkey="TK_KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_TK_KIKIKOSEI_MCM_TK_KIKIMEISAI" ns1:Generator_RelationVarName="relationMCM_TK_KIKIKOSEI_MCM_TK_KIKIMEISAI" ns1:Generator_UserChildTable="MCM_TK_KIKIMEISAI" ns1:Generator_UserParentTable="MCM_TK_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_TK_KIKIKOSEIRow" ns1:Generator_ChildPropName="GetMCM_TK_KIKIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKIMEISAI_MCM_TK_KOTAIMEISAI" ns0:parent="MCM_TK_KIKIMEISAI" ns0:child="MCM_TK_KOTAIMEISAI" ns0:parentkey="TK_KIKIMEISAI_ID" ns0:childkey="TK_KIKIMEISAI_ID" ns1:Generator_UserRelationName="MCM_TK_KIKIMEISAI_MCM_TK_KOTAIMEISAI" ns1:Generator_RelationVarName="relationMCM_TK_KIKIMEISAI_MCM_TK_KOTAIMEISAI" ns1:Generator_UserChildTable="MCM_TK_KOTAIMEISAI" ns1:Generator_UserParentTable="MCM_TK_KIKIMEISAI" ns1:Generator_ParentPropName="MCM_TK_KIKIMEISAIRow" ns1:Generator_ChildPropName="GetMCM_TK_KOTAIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKIMEISAI_MCM_TK_TANKA" ns0:parent="MCM_TK_KIKIMEISAI" ns0:child="MCM_TK_TANKA" ns0:parentkey="TK_KIKIMEISAI_ID" ns0:childkey="TK_KIKIMEISAI_ID" ns1:Generator_UserRelationName="MCM_TK_KIKIMEISAI_MCM_TK_TANKA" ns1:Generator_RelationVarName="relationMCM_TK_KIKIMEISAI_MCM_TK_TANKA" ns1:Generator_UserChildTable="MCM_TK_TANKA" ns1:Generator_UserParentTable="MCM_TK_KIKIMEISAI" ns1:Generator_ParentPropName="MCM_TK_KIKIMEISAIRow" ns1:Generator_ChildPropName="GetMCM_TK_TANKARows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TK_KIKIKOSEI_MCM_TK_TENKEN" ns0:parent="MCM_TK_KIKIKOSEI" ns0:child="MCM_TK_TENKEN" ns0:parentkey="TK_KIKIKOSEI_ID" ns0:childkey="TK_KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_TK_KIKIKOSEI_MCM_TK_TENKEN" ns1:Generator_RelationVarName="relationMCM_TK_KIKIKOSEI_MCM_TK_TENKEN" ns1:Generator_UserChildTable="MCM_TK_TENKEN" ns1:Generator_UserParentTable="MCM_TK_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_TK_KIKIKOSEIRow" ns1:Generator_ChildPropName="GetMCM_TK_TENKENRows" />
    
```

</details>

### 1006　Mcm1006u1DataSet.xsd

[Mcm1006u1DataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1DataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_TM\_KIKAN | TM\_KIKAN\_ID | {"Generator\_UserColumnName": "TM\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_KIKAN | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | HYOJUNGOKEI\_KIN | {"Generator\_UserColumnName": "HYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUNGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUNGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | SIKIRISYOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRISYOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRISYOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | SYUSSEINEBIKI\_KIN | {"Generator\_UserColumnName": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInRow": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInTable": "SYUSSEINEBIKI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | SIKIRIGOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRIGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRIGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_KIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KIKAN | TM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "TM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "TM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "TM\_MITSUMORI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TM\_MITSUMORI | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_MITSUMORI | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_TM\_MITSUMORI | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_MITSUMORI | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_MITSUMORI | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_MITSUMORI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_TM\_MITSUMORI | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_MITSUMORI | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TM\_MITSUMORI | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TM\_MITSUMORI | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_MITSUMORI | TORITEL\_NO | {"Generator\_UserColumnName": "TORITEL\_NO", "Generator\_ColumnPropNameInRow": "TORITEL\_NO", "Generator\_ColumnVarNameInTable": "columnTORITEL\_NO", "Generator\_ColumnPropNameInTable": "TORITEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | TORIFAX\_NO | {"Generator\_UserColumnName": "TORIFAX\_NO", "Generator\_ColumnPropNameInRow": "TORIFAX\_NO", "Generator\_ColumnVarNameInTable": "columnTORIFAX\_NO", "Generator\_ColumnPropNameInTable": "TORIFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | TORIJIGYOSYO\_NK | {"Generator\_UserColumnName": "TORIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "TORIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnTORIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "TORIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TM\_MITSUMORI | TORISYUTANTOSYA\_NK | {"Generator\_UserColumnName": "TORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "TORISYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnTORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "TORISYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TM\_MITSUMORI | TORIASSISTANT\_NK | {"Generator\_UserColumnName": "TORIASSISTANT\_NK", "Generator\_ColumnPropNameInRow": "TORIASSISTANT\_NK", "Generator\_ColumnVarNameInTable": "columnTORIASSISTANT\_NK", "Generator\_ColumnPropNameInTable": "TORIASSISTANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TM\_MITSUMORI | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | IRAIJIGYOSYO\_NK | {"Generator\_UserColumnName": "IRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "IRAIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "IRAIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "200"}]] |
| MCM\_TM\_MITSUMORI | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | KAITOKIZITSU\_DT | {"Generator\_UserColumnName": "KAITOKIZITSU\_DT", "Generator\_ColumnPropNameInRow": "KAITOKIZITSU\_DT", "Generator\_ColumnVarNameInTable": "columnKAITOKIZITSU\_DT", "Generator\_ColumnPropNameInTable": "KAITOKIZITSU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_MITSUMORI | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_MITSUMORI | YAKANTAIOUMU | {"Generator\_UserColumnName": "YAKANTAIOUMU", "Generator\_ColumnPropNameInRow": "YAKANTAIOUMU", "Generator\_ColumnVarNameInTable": "columnYAKANTAIOUMU", "Generator\_ColumnPropNameInTable": "YAKANTAIOUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_MITSUMORI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_MITSUMORI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_MITSUMORI | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TM\_MITSUMORI | TENKENKANOYOBI | {"Generator\_UserColumnName": "TENKENKANOYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENKANOYOBI", "Generator\_ColumnPropNameInRow": "TENKENKANOYOBI", "Generator\_ColumnPropNameInTable": "TENKENKANOYOBIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_MITSUMORI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_MITSUMORI | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_MITSUMORI | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_MITSUMORI | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_TANKA | TM\_TANKA\_ID | {"Generator\_UserColumnName": "TM\_TANKA\_ID", "Generator\_ColumnPropNameInRow": "TM\_TANKA\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_TANKA\_ID", "Generator\_ColumnPropNameInTable": "TM\_TANKA\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_TANKA | TM\_KIKAN\_ID | {"Generator\_UserColumnName": "TM\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | HYOJUN\_KIN | {"Generator\_UserColumnName": "HYOJUN\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUN\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUN\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUN\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | SIKIRI\_KIN | {"Generator\_UserColumnName": "SIKIRI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_TANKA | KEIYAKUNAIYO | {"Generator\_UserColumnName": "KEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "KEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "KEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TM\_TANKA | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_TANKA | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_TANKA | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_TANKA | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_TANKA | TMF\_BIKO | {"Generator\_UserColumnName": "TMF\_BIKO", "Generator\_ColumnPropNameInRow": "TMF\_BIKO", "Generator\_ColumnVarNameInTable": "columnTMF\_BIKO", "Generator\_ColumnPropNameInTable": "TMF\_BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_TANKA | TM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_TANKA | SEIZOMAKER\_ID | {"Generator\_UserColumnName": "SEIZOMAKER\_ID", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_ID", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_ID", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_TANKA | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_TANKA | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TM\_TANKA | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_TANKA | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_TANKA | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1">
      <xs:selector xpath=".//mstns:MCM_TM_KIKAN" />
      <xs:field xpath="mstns:TM_KIKAN_ID" />
      <xs:field xpath="mstns:TM_KEIYAKUJIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint2" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_KIKAN" />
      <xs:field xpath="mstns:TM_KIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_MITSUMORI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_MITSUMORI" />
      <xs:field xpath="mstns:TM_IRAI_ID" />
      <xs:field xpath="mstns:TM_KEIYAKUJIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_TANKA_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_TANKA" />
      <xs:field xpath="mstns:TM_TANKA_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TM_KIKAN_MCM_TM_TANKA" ns0:parent="MCM_TM_KIKAN" ns0:child="MCM_TM_TANKA" ns0:parentkey="TM_KIKAN_ID" ns0:childkey="TM_KIKAN_ID" ns1:Generator_UserRelationName="MCM_TM_KIKAN_MCM_TM_TANKA" ns1:Generator_RelationVarName="relationMCM_TM_KIKAN_MCM_TM_TANKA" ns1:Generator_UserChildTable="MCM_TM_TANKA" ns1:Generator_UserParentTable="MCM_TM_KIKAN" ns1:Generator_ParentPropName="MCM_TM_KIKANRow" ns1:Generator_ChildPropName="GetMCM_TM_TANKARows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_TM_MITSUMORI_MCM_TM_KIKAN" ns0:parent="MCM_TM_MITSUMORI" ns0:child="MCM_TM_KIKAN" ns0:parentkey="TM_KEIYAKUJIKAN_ID" ns0:childkey="TM_KEIYAKUJIKAN_ID" ns1:Generator_UserRelationName="MCM_TM_MITSUMORI_MCM_TM_KIKAN" ns1:Generator_RelationVarName="relationMCM_TM_MITSUMORI_MCM_TM_KIKAN" ns1:Generator_UserChildTable="MCM_TM_KIKAN" ns1:Generator_UserParentTable="MCM_TM_MITSUMORI" ns1:Generator_ParentPropName="MCM_TM_MITSUMORIRow" ns1:Generator_ChildPropName="GetMCM_TM_KIKANRows" />
    
```

</details>

### 1006　Mcm1006u2DataSet.xsd

[Mcm1006u2DataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2DataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_MA\_KIKIKOSEI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOSEI | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn"} | [["maxLength", {"value": "100"}]] |
| MCM\_MA\_KIKIKOSEI | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOSEI | TANI | {"Generator\_UserColumnName": "TANI", "Generator\_ColumnVarNameInTable": "columnTANI", "Generator\_ColumnPropNameInRow": "TANI", "Generator\_ColumnPropNameInTable": "TANIColumn"} | [["maxLength", {"value": "10"}]] |
| MCM\_MA\_KIKIKOSEI | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_KIKIKOSEI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_KIKIKOSEI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | TM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_KIKIMEISAI | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | SEIZOMAKER\_ID | {"Generator\_UserColumnName": "SEIZOMAKER\_ID", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_ID", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_ID", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | MAF\_HYOJIJUN | {"Generator\_UserColumnName": "MAF\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAF\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAF\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAF\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn"} | [["maxLength", {"value": "100"}]] |
| MCM\_MA\_KIKIMEISAI | KOTAIKANRI\_FLG | {"Generator\_UserColumnName": "KOTAIKANRI\_FLG", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_FLG", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_FLG", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | DUMMY\_FLG | {"Generator\_UserColumnName": "DUMMY\_FLG", "Generator\_ColumnVarNameInTable": "columnDUMMY\_FLG", "Generator\_ColumnPropNameInRow": "DUMMY\_FLG", "Generator\_ColumnPropNameInTable": "DUMMY\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | OYAKIKIMEISAI\_ID | {"Generator\_UserColumnName": "OYAKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "OYAKIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnOYAKIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "OYAKIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | OYAKIKIHINMEI\_NK | {"Generator\_UserColumnName": "OYAKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "OYAKIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnOYAKIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "OYAKIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | OYAKIKIKATASHIKI | {"Generator\_UserColumnName": "OYAKIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "OYAKIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnOYAKIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "OYAKIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | OYAKIKIBUNRUI\_CD | {"Generator\_UserColumnName": "OYAKIKIBUNRUI\_CD", "Generator\_ColumnPropNameInRow": "OYAKIKIBUNRUI\_CD", "Generator\_ColumnVarNameInTable": "columnOYAKIKIBUNRUI\_CD", "Generator\_ColumnPropNameInTable": "OYAKIKIBUNRUI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | KOTAIKANRI\_ID | {"Generator\_UserColumnName": "KOTAIKANRI\_ID", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_ID", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_ID", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | KOTAI\_NK | {"Generator\_UserColumnName": "KOTAI\_NK", "Generator\_ColumnPropNameInRow": "KOTAI\_NK", "Generator\_ColumnVarNameInTable": "columnKOTAI\_NK", "Generator\_ColumnPropNameInTable": "KOTAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | SERIAL\_NO | {"Generator\_UserColumnName": "SERIAL\_NO", "Generator\_ColumnPropNameInRow": "SERIAL\_NO", "Generator\_ColumnVarNameInTable": "columnSERIAL\_NO", "Generator\_ColumnPropNameInTable": "SERIAL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | ITIJINONYU\_DT | {"Generator\_UserColumnName": "ITIJINONYU\_DT", "Generator\_ColumnPropNameInRow": "ITIJINONYU\_DT", "Generator\_ColumnVarNameInTable": "columnITIJINONYU\_DT", "Generator\_ColumnPropNameInTable": "ITIJINONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | SETCHIBASYO | {"Generator\_UserColumnName": "SETCHIBASYO", "Generator\_ColumnPropNameInRow": "SETCHIBASYO", "Generator\_ColumnVarNameInTable": "columnSETCHIBASYO", "Generator\_ColumnPropNameInTable": "SETCHIBASYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KOTAIMEISAI\_V | TEKKYO\_DT | {"Generator\_UserColumnName": "TEKKYO\_DT", "Generator\_ColumnPropNameInRow": "TEKKYO\_DT", "Generator\_ColumnVarNameInTable": "columnTEKKYO\_DT", "Generator\_ColumnPropNameInTable": "TEKKYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | KEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "KEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | ENCHOKEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "ENCHOKEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TM\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "TM\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KOTAIMEISAI\_V | TK\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TM\_TANKA\_ID | {"Generator\_UserColumnName": "TM\_TANKA\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_TANKA\_ID", "Generator\_ColumnPropNameInRow": "TM\_TANKA\_ID", "Generator\_ColumnPropNameInTable": "TM\_TANKA\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_TANKA | TM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TM\_KIKAN\_ID | {"Generator\_UserColumnName": "TM\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | HYOJUN\_KIN | {"Generator\_UserColumnName": "HYOJUN\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUN\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUN\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUN\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | SIKIRI\_KIN | {"Generator\_UserColumnName": "SIKIRI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | KEIYAKUNAIYO | {"Generator\_UserColumnName": "KEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "KEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "KEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TM\_TANKA | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_TANKA | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_TANKA | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_TANKA | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_TANKA | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_TANKA | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_TANKA | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_TANKA | EXPR1 | {"Generator\_UserColumnName": "EXPR1", "Generator\_ColumnVarNameInTable": "columnEXPR1", "Generator\_ColumnPropNameInRow": "EXPR1", "Generator\_ColumnPropNameInTable": "EXPR1Column", "type": "xs:decimal"} | [] |
| MCM\_TM\_TANKA | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | SYUSSEINEBIKI\_KIN | {"Generator\_UserColumnName": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInRow": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInTable": "SYUSSEINEBIKI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_TANKA | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TK\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TK\_TANKA\_ID | {"Generator\_UserColumnName": "TK\_TANKA\_ID", "Generator\_ColumnPropNameInRow": "TK\_TANKA\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TANKA\_ID", "Generator\_ColumnPropNameInTable": "TK\_TANKA\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | MAF\_HYOJIJUN | {"Generator\_UserColumnName": "MAF\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAF\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAF\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAF\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_TANKA | TK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_KIKIKOSEI" />
      <xs:field xpath="mstns:KIKIKOSEI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_KIKIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_KIKIMEISAI" />
      <xs:field xpath="mstns:KIKIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_TANKA_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_TANKA" />
      <xs:field xpath="mstns:TM_TANKA_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI" ns0:parent="MCM_MA_KIKIKOSEI" ns0:child="MCM_MA_KIKIMEISAI" ns0:parentkey="KIKIKOSEI_ID" ns0:childkey="KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI" ns1:Generator_RelationVarName="relationMCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI" ns1:Generator_UserChildTable="MCM_MA_KIKIMEISAI" ns1:Generator_UserParentTable="MCM_MA_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_MA_KIKIKOSEIRow" ns1:Generator_ChildPropName="GetMCM_MA_KIKIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_MA_KIKIKOSEI_MCM_MA_KOTAIMEISAI_V" ns0:parent="MCM_MA_KIKIKOSEI" ns0:child="MCM_MA_KOTAIMEISAI_V" ns0:parentkey="KIKIKOSEI_ID" ns0:childkey="KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_MA_KIKIKOSEI_MCM_MA_KOTAIMEISAI_V" ns1:Generator_RelationVarName="relationMCM_MA_KIKIKOSEI_MCM_MA_KOTAIMEISAI_V" ns1:Generator_UserChildTable="MCM_MA_KOTAIMEISAI_V" ns1:Generator_UserParentTable="MCM_MA_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_MA_KIKIKOSEIRow" ns1:Generator_ChildPropName="GetMCM_MA_KOTAIMEISAI_VRows" />
    
```

</details>

### 1009　Mcm1009uDataSet.xsd

[Mcm1009uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_TK\_KEIYAKU | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | IRAIJIGYOSYO\_NK | {"Generator\_UserColumnName": "IRAIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "IRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "IRAIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KEIYAKU | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUKAISI\_DT | {"Generator\_UserColumnName": "KEIYAKUKAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUSYURYO\_DT | {"Generator\_UserColumnName": "KEIYAKUSYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUSYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIDOKOSIN\_FLG | {"Generator\_UserColumnName": "JIDOKOSIN\_FLG", "Generator\_ColumnVarNameInTable": "columnJIDOKOSIN\_FLG", "Generator\_ColumnPropNameInRow": "JIDOKOSIN\_FLG", "Generator\_ColumnPropNameInTable": "JIDOKOSIN\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnJIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKU\_NO | {"Generator\_UserColumnName": "PACKKEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | SHOKAI\_KEIYAKU\_DT | {"Generator\_UserColumnName": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnSHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "SHOKAI\_KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TSUKI | {"Generator\_UserColumnName": "TSUKI", "Generator\_ColumnVarNameInTable": "columnTSUKI", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_TK\_SIHARAIMEISAI | ON\_FLG | {"Generator\_UserColumnName": "ON\_FLG", "Generator\_ColumnVarNameInTable": "columnON\_FLG", "Generator\_ColumnPropNameInRow": "ON\_FLG", "Generator\_ColumnPropNameInTable": "ON\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_DT | {"Generator\_UserColumnName": "ON\_DT", "Generator\_ColumnVarNameInTable": "columnON\_DT", "Generator\_ColumnPropNameInRow": "ON\_DT", "Generator\_ColumnPropNameInTable": "ON\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_BY | {"Generator\_UserColumnName": "ON\_BY", "Generator\_ColumnVarNameInTable": "columnON\_BY", "Generator\_ColumnPropNameInRow": "ON\_BY", "Generator\_ColumnPropNameInTable": "ON\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_FLG | {"Generator\_UserColumnName": "OFF\_FLG", "Generator\_ColumnVarNameInTable": "columnOFF\_FLG", "Generator\_ColumnPropNameInRow": "OFF\_FLG", "Generator\_ColumnPropNameInTable": "OFF\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_DT | {"Generator\_UserColumnName": "OFF\_DT", "Generator\_ColumnVarNameInTable": "columnOFF\_DT", "Generator\_ColumnPropNameInRow": "OFF\_DT", "Generator\_ColumnPropNameInTable": "OFF\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_BY | {"Generator\_UserColumnName": "OFF\_BY", "Generator\_ColumnVarNameInTable": "columnOFF\_BY", "Generator\_ColumnPropNameInRow": "OFF\_BY", "Generator\_ColumnPropNameInTable": "OFF\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_KEIYAKUJIKAN | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_KEIYAKUJIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_1003\_V | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_1003\_V | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | KAITOKIZITSU\_DT | {"Generator\_UserColumnName": "KAITOKIZITSU\_DT", "Generator\_ColumnPropNameInRow": "KAITOKIZITSU\_DT", "Generator\_ColumnVarNameInTable": "columnKAITOKIZITSU\_DT", "Generator\_ColumnPropNameInTable": "KAITOKIZITSU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_1003\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | FUSEI | {"Generator\_UserColumnName": "FUSEI", "Generator\_ColumnPropNameInRow": "FUSEI", "Generator\_ColumnVarNameInTable": "columnFUSEI", "Generator\_ColumnPropNameInTable": "FUSEIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_1003\_V | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_1003\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_1003\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | KYUNONYUSAKI\_NK | {"Generator\_UserColumnName": "KYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "KYUNONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnKYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "KYUNONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | NONYUSAKIKANA\_KN | {"Generator\_UserColumnName": "NONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInRow": "NONYUSAKIKANA\_KN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInTable": "NONYUSAKIKANA\_KNColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_1003\_V | NONYUSAKIEIMEI\_EN | {"Generator\_UserColumnName": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInRow": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInTable": "NONYUSAKIEIMEI\_ENColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_1003\_V | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_1003\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_1003\_V | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_1003\_V | TENKENKANOYOBI | {"Generator\_UserColumnName": "TENKENKANOYOBI", "Generator\_ColumnPropNameInRow": "TENKENKANOYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENKANOYOBI", "Generator\_ColumnPropNameInTable": "TENKENKANOYOBIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_1003\_V | YAKANTAIOUMU | {"Generator\_UserColumnName": "YAKANTAIOUMU", "Generator\_ColumnPropNameInRow": "YAKANTAIOUMU", "Generator\_ColumnVarNameInTable": "columnYAKANTAIOUMU", "Generator\_ColumnPropNameInTable": "YAKANTAIOUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_1003\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | MITSUMORI\_IRAI | {"Generator\_UserColumnName": "MITSUMORI\_IRAI", "Generator\_ColumnPropNameInRow": "MITSUMORI\_IRAI", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_IRAI", "Generator\_ColumnPropNameInTable": "MITSUMORI\_IRAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | MITSUMORI\_COPY | {"Generator\_UserColumnName": "MITSUMORI\_COPY", "Generator\_ColumnPropNameInRow": "MITSUMORI\_COPY", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_COPY", "Generator\_ColumnPropNameInTable": "MITSUMORI\_COPYColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | MITSUMORI\_DEL | {"Generator\_UserColumnName": "MITSUMORI\_DEL", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DEL", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | KEIYAKU\_KEIYAKU | {"Generator\_UserColumnName": "KEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInRow": "KEIYAKU\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInTable": "KEIYAKU\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | KEIYAKU\_DEL | {"Generator\_UserColumnName": "KEIYAKU\_DEL", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DEL", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DEL", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_1003\_V | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KEIYAKU" />
      <xs:field xpath="mstns:TK_KEIYAKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_SIHARAIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_SIHARAIMEISAI" />
      <xs:field xpath="mstns:TK_SIHARAIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_KEIYAKUJIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_KEIYAKUJIKAN" />
      <xs:field xpath="mstns:TM_KEIYAKUJIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_1003_V_Constraint1" ns1:ConstraintName="Constraint1">
      <xs:selector xpath=".//mstns:MCM_1003_V" />
      <xs:field xpath="mstns:NONYUSAKI_ID" />
      <xs:field xpath="mstns:SUPPORT_ID" />
    </xs:unique>
  
```

</details>

### 1010　Mcm1010uDataSet.xsd

[Mcm1010uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_1003\_V | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_1003\_V | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_1003\_V | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | KAITOKIZITSU\_DT | {"Generator\_UserColumnName": "KAITOKIZITSU\_DT", "Generator\_ColumnPropNameInRow": "KAITOKIZITSU\_DT", "Generator\_ColumnVarNameInTable": "columnKAITOKIZITSU\_DT", "Generator\_ColumnPropNameInTable": "KAITOKIZITSU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_1003\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_1003\_V | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | FUSEI | {"Generator\_UserColumnName": "FUSEI", "Generator\_ColumnPropNameInRow": "FUSEI", "Generator\_ColumnVarNameInTable": "columnFUSEI", "Generator\_ColumnPropNameInTable": "FUSEIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_1003\_V | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_1003\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_1003\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | KYUNONYUSAKI\_NK | {"Generator\_UserColumnName": "KYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "KYUNONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnKYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "KYUNONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | NONYUSAKIKANA\_KN | {"Generator\_UserColumnName": "NONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInRow": "NONYUSAKIKANA\_KN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInTable": "NONYUSAKIKANA\_KNColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_1003\_V | NONYUSAKIEIMEI\_EN | {"Generator\_UserColumnName": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInRow": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInTable": "NONYUSAKIEIMEI\_ENColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_1003\_V | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_1003\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_1003\_V | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_1003\_V | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_1003\_V | TENKENKANOYOBI | {"Generator\_UserColumnName": "TENKENKANOYOBI", "Generator\_ColumnPropNameInRow": "TENKENKANOYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENKANOYOBI", "Generator\_ColumnPropNameInTable": "TENKENKANOYOBIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_1003\_V | YAKANTAIOUMU | {"Generator\_UserColumnName": "YAKANTAIOUMU", "Generator\_ColumnPropNameInRow": "YAKANTAIOUMU", "Generator\_ColumnVarNameInTable": "columnYAKANTAIOUMU", "Generator\_ColumnPropNameInTable": "YAKANTAIOUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_1003\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_1003\_V | MITSUMORI\_IRAI | {"Generator\_UserColumnName": "MITSUMORI\_IRAI", "Generator\_ColumnPropNameInRow": "MITSUMORI\_IRAI", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_IRAI", "Generator\_ColumnPropNameInTable": "MITSUMORI\_IRAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | MITSUMORI\_COPY | {"Generator\_UserColumnName": "MITSUMORI\_COPY", "Generator\_ColumnPropNameInRow": "MITSUMORI\_COPY", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_COPY", "Generator\_ColumnPropNameInTable": "MITSUMORI\_COPYColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | MITSUMORI\_DEL | {"Generator\_UserColumnName": "MITSUMORI\_DEL", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DEL", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | KEIYAKU\_KEIYAKU | {"Generator\_UserColumnName": "KEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInRow": "KEIYAKU\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInTable": "KEIYAKU\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | KEIYAKU\_DEL | {"Generator\_UserColumnName": "KEIYAKU\_DEL", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DEL", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DEL", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_1003\_V | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_1003\_V | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_1003\_V | MEISHO4\_NK | {"Generator\_UserColumnName": "MEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "MEISHO4\_NK", "Generator\_ColumnPropNameInTable": "MEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_1003\_V | JIGYOSHO\_CD | {"Generator\_UserColumnName": "JIGYOSHO\_CD", "Generator\_ColumnPropNameInRow": "JIGYOSHO\_CD", "Generator\_ColumnVarNameInTable": "columnJIGYOSHO\_CD", "Generator\_ColumnPropNameInTable": "JIGYOSHO\_CDColumn"} | [["maxLength", {"value": "10"}]] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_KEIYAKUJIKAN | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_KEIYAKUJIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TSUKI | {"Generator\_UserColumnName": "TSUKI", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnVarNameInTable": "columnTSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_TK\_SIHARAIMEISAI | ON\_FLG | {"Generator\_UserColumnName": "ON\_FLG", "Generator\_ColumnPropNameInRow": "ON\_FLG", "Generator\_ColumnVarNameInTable": "columnON\_FLG", "Generator\_ColumnPropNameInTable": "ON\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_DT | {"Generator\_UserColumnName": "ON\_DT", "Generator\_ColumnPropNameInRow": "ON\_DT", "Generator\_ColumnVarNameInTable": "columnON\_DT", "Generator\_ColumnPropNameInTable": "ON\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_BY | {"Generator\_UserColumnName": "ON\_BY", "Generator\_ColumnPropNameInRow": "ON\_BY", "Generator\_ColumnVarNameInTable": "columnON\_BY", "Generator\_ColumnPropNameInTable": "ON\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_FLG | {"Generator\_UserColumnName": "OFF\_FLG", "Generator\_ColumnPropNameInRow": "OFF\_FLG", "Generator\_ColumnVarNameInTable": "columnOFF\_FLG", "Generator\_ColumnPropNameInTable": "OFF\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_DT | {"Generator\_UserColumnName": "OFF\_DT", "Generator\_ColumnPropNameInRow": "OFF\_DT", "Generator\_ColumnVarNameInTable": "columnOFF\_DT", "Generator\_ColumnPropNameInTable": "OFF\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_BY | {"Generator\_UserColumnName": "OFF\_BY", "Generator\_ColumnPropNameInRow": "OFF\_BY", "Generator\_ColumnVarNameInTable": "columnOFF\_BY", "Generator\_ColumnPropNameInTable": "OFF\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | IRAIJIGYOSYO\_NK | {"Generator\_UserColumnName": "IRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "IRAIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "IRAIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KEIYAKU | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUKAISI\_DT | {"Generator\_UserColumnName": "KEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUSYURYO\_DT | {"Generator\_UserColumnName": "KEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUSYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUSYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIDOKOSIN\_FLG | {"Generator\_UserColumnName": "JIDOKOSIN\_FLG", "Generator\_ColumnPropNameInRow": "JIDOKOSIN\_FLG", "Generator\_ColumnVarNameInTable": "columnJIDOKOSIN\_FLG", "Generator\_ColumnPropNameInTable": "JIDOKOSIN\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnJIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKU\_NO | {"Generator\_UserColumnName": "PACKKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | SHOKAI\_KEIYAKU\_DT | {"Generator\_UserColumnName": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnSHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "SHOKAI\_KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1">
      <xs:selector xpath=".//mstns:MCM_1003_V" />
      <xs:field xpath="mstns:NONYUSAKI_ID" />
      <xs:field xpath="mstns:SUPPORT_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_KEIYAKUJIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_KEIYAKUJIKAN" />
      <xs:field xpath="mstns:TM_KEIYAKUJIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_SIHARAIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_SIHARAIMEISAI" />
      <xs:field xpath="mstns:TK_SIHARAIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KEIYAKU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KEIYAKU" />
      <xs:field xpath="mstns:TK_KEIYAKU_ID" />
    </xs:unique>
  
```

</details>

### 1011　Mcm1011uDataSet.xsd

[Mcm1011uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_UM\_MITSUMORI | SYOUNIN\_JOTAI | {"Generator\_UserColumnName": "SYOUNIN\_JOTAI", "Generator\_ColumnPropNameInRow": "SYOUNIN\_JOTAI", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_JOTAI", "Generator\_ColumnPropNameInTable": "SYOUNIN\_JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_MITSUMORI | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_MITSUMORI | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_MITSUMORI | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_UM\_MITSUMORI | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_MITSUMORI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | RELATION\_ID | {"Generator\_UserColumnName": "RELATION\_ID", "Generator\_ColumnVarNameInTable": "columnRELATION\_ID", "Generator\_ColumnPropNameInRow": "RELATION\_ID", "Generator\_ColumnPropNameInTable": "RELATION\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_MITSUMORI | SHORUI\_NO | {"Generator\_UserColumnName": "SHORUI\_NO", "Generator\_ColumnVarNameInTable": "columnSHORUI\_NO", "Generator\_ColumnPropNameInRow": "SHORUI\_NO", "Generator\_ColumnPropNameInTable": "SHORUI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UM\_MITSUMORI | KINGAKU | {"Generator\_UserColumnName": "KINGAKU", "Generator\_ColumnVarNameInTable": "columnKINGAKU", "Generator\_ColumnPropNameInRow": "KINGAKU", "Generator\_ColumnPropNameInTable": "KINGAKUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | IRAITANTO\_NK | {"Generator\_UserColumnName": "IRAITANTO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITANTO\_NK", "Generator\_ColumnPropNameInRow": "IRAITANTO\_NK", "Generator\_ColumnPropNameInTable": "IRAITANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_MITSUMORI | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | SHONIN\_NO | {"Generator\_UserColumnName": "SHONIN\_NO", "Generator\_ColumnVarNameInTable": "columnSHONIN\_NO", "Generator\_ColumnPropNameInRow": "SHONIN\_NO", "Generator\_ColumnPropNameInTable": "SHONIN\_NOColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_MITSUMORI | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | SHINSA\_DT | {"Generator\_UserColumnName": "SHINSA\_DT", "Generator\_ColumnVarNameInTable": "columnSHINSA\_DT", "Generator\_ColumnPropNameInRow": "SHINSA\_DT", "Generator\_ColumnPropNameInTable": "SHINSA\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | SHINSA\_BY | {"Generator\_UserColumnName": "SHINSA\_BY", "Generator\_ColumnVarNameInTable": "columnSHINSA\_BY", "Generator\_ColumnPropNameInRow": "SHINSA\_BY", "Generator\_ColumnPropNameInTable": "SHINSA\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_MITSUMORI | SYOUNIN\_DT | {"Generator\_UserColumnName": "SYOUNIN\_DT", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_DT", "Generator\_ColumnPropNameInRow": "SYOUNIN\_DT", "Generator\_ColumnPropNameInTable": "SYOUNIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | SYOUNIN\_BY | {"Generator\_UserColumnName": "SYOUNIN\_BY", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_BY", "Generator\_ColumnPropNameInRow": "SYOUNIN\_BY", "Generator\_ColumnPropNameInTable": "SYOUNIN\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_MITSUMORI | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn"} | [["maxLength", {"value": "12"}]] |
| MCM\_UM\_MITSUMORI | EXPR1 | {"Generator\_UserColumnName": "EXPR1", "Generator\_ColumnPropNameInRow": "EXPR1", "Generator\_ColumnVarNameInTable": "columnEXPR1", "Generator\_ColumnPropNameInTable": "EXPR1Column"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_MITSUMORI | KYUNONYUSAKI\_NK | {"Generator\_UserColumnName": "KYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "KYUNONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnKYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "KYUNONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_MITSUMORI | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_MITSUMORI | NONYUSAKIKANA\_KN | {"Generator\_UserColumnName": "NONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInRow": "NONYUSAKIKANA\_KN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInTable": "NONYUSAKIKANA\_KNColumn"} | [["maxLength", {"value": "60"}]] |
| MCM\_UM\_MITSUMORI | NONYUSAKIEIMEI\_EN | {"Generator\_UserColumnName": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInRow": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInTable": "NONYUSAKIEIMEI\_ENColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_UM\_TENPU | TENPUFILE\_NK | {"Generator\_UserColumnName": "TENPUFILE\_NK", "Generator\_ColumnPropNameInRow": "TENPUFILE\_NK", "Generator\_ColumnVarNameInTable": "columnTENPUFILE\_NK", "Generator\_ColumnPropNameInTable": "TENPUFILE\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_UM\_TENPU | DIRECTORY | {"Generator\_UserColumnName": "DIRECTORY", "Generator\_ColumnPropNameInRow": "DIRECTORY", "Generator\_ColumnVarNameInTable": "columnDIRECTORY", "Generator\_ColumnPropNameInTable": "DIRECTORYColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_UM\_TENPU | TENPU\_ID | {"Generator\_UserColumnName": "TENPU\_ID", "Generator\_ColumnPropNameInRow": "TENPU\_ID", "Generator\_ColumnVarNameInTable": "columnTENPU\_ID", "Generator\_ColumnPropNameInTable": "TENPU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_TENPU | RELATION\_ID | {"Generator\_UserColumnName": "RELATION\_ID", "Generator\_ColumnVarNameInTable": "columnRELATION\_ID", "Generator\_ColumnPropNameInRow": "RELATION\_ID", "Generator\_ColumnPropNameInTable": "RELATION\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_TENPU | SHONIN\_NO | {"Generator\_UserColumnName": "SHONIN\_NO", "Generator\_ColumnVarNameInTable": "columnSHONIN\_NO", "Generator\_ColumnPropNameInRow": "SHONIN\_NO", "Generator\_ColumnPropNameInTable": "SHONIN\_NOColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | IRAIJIGYOSYO\_NK | {"Generator\_UserColumnName": "IRAIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "IRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "IRAIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "200"}]] |
| MCM\_TK\_KEIYAKU | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUKAISI\_DT | {"Generator\_UserColumnName": "KEIYAKUKAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUSYURYO\_DT | {"Generator\_UserColumnName": "KEIYAKUSYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUSYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIDOKOSIN\_FLG | {"Generator\_UserColumnName": "JIDOKOSIN\_FLG", "Generator\_ColumnVarNameInTable": "columnJIDOKOSIN\_FLG", "Generator\_ColumnPropNameInRow": "JIDOKOSIN\_FLG", "Generator\_ColumnPropNameInTable": "JIDOKOSIN\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnJIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKU\_NO | {"Generator\_UserColumnName": "PACKKEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | SHINSA\_DT | {"Generator\_UserColumnName": "SHINSA\_DT", "Generator\_ColumnPropNameInRow": "SHINSA\_DT", "Generator\_ColumnVarNameInTable": "columnSHINSA\_DT", "Generator\_ColumnPropNameInTable": "SHINSA\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | SHINSA\_BY | {"Generator\_UserColumnName": "SHINSA\_BY", "Generator\_ColumnPropNameInRow": "SHINSA\_BY", "Generator\_ColumnVarNameInTable": "columnSHINSA\_BY", "Generator\_ColumnPropNameInTable": "SHINSA\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | SHONIN\_DT | {"Generator\_UserColumnName": "SHONIN\_DT", "Generator\_ColumnPropNameInRow": "SHONIN\_DT", "Generator\_ColumnVarNameInTable": "columnSHONIN\_DT", "Generator\_ColumnPropNameInTable": "SHONIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | SHONIN\_BY | {"Generator\_UserColumnName": "SHONIN\_BY", "Generator\_ColumnPropNameInRow": "SHONIN\_BY", "Generator\_ColumnVarNameInTable": "columnSHONIN\_BY", "Generator\_ColumnPropNameInTable": "SHONIN\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | UM\_KIHON\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_MITSUMORI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SYOUNIN\_JOTAI | {"Generator\_UserColumnName": "SYOUNIN\_JOTAI", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_JOTAI", "Generator\_ColumnPropNameInRow": "SYOUNIN\_JOTAI", "Generator\_ColumnPropNameInTable": "SYOUNIN\_JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORISAKUSEISYA\_NK | {"Generator\_UserColumnName": "MITSUMORISAKUSEISYA\_NK", "Generator\_ColumnVarNameInTable": "columnMITSUMORISAKUSEISYA\_NK", "Generator\_ColumnPropNameInRow": "MITSUMORISAKUSEISYA\_NK", "Generator\_ColumnPropNameInTable": "MITSUMORISAKUSEISYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORIKIGEN | {"Generator\_UserColumnName": "MITSUMORIKIGEN", "Generator\_ColumnVarNameInTable": "columnMITSUMORIKIGEN", "Generator\_ColumnPropNameInRow": "MITSUMORIKIGEN", "Generator\_ColumnPropNameInTable": "MITSUMORIKIGENColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAITENPO\_ID | {"Generator\_UserColumnName": "IRAITENPO\_ID", "Generator\_ColumnVarNameInTable": "columnIRAITENPO\_ID", "Generator\_ColumnPropNameInRow": "IRAITENPO\_ID", "Generator\_ColumnPropNameInTable": "IRAITENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO1\_NK | {"Generator\_UserColumnName": "IRAIMEISHO1\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO1\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO2\_NK | {"Generator\_UserColumnName": "IRAIMEISHO2\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO2\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO3\_NK | {"Generator\_UserColumnName": "IRAIMEISHO3\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO3\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO3\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO4\_NK | {"Generator\_UserColumnName": "IRAIMEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAITENPORYAKU\_NK | {"Generator\_UserColumnName": "IRAITENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "IRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "IRAITENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAITANTO\_NK | {"Generator\_UserColumnName": "IRAITANTO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITANTO\_NK", "Generator\_ColumnPropNameInRow": "IRAITANTO\_NK", "Generator\_ColumnPropNameInTable": "IRAITANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUTENPO\_ID | {"Generator\_UserColumnName": "SOFUTENPO\_ID", "Generator\_ColumnVarNameInTable": "columnSOFUTENPO\_ID", "Generator\_ColumnPropNameInRow": "SOFUTENPO\_ID", "Generator\_ColumnPropNameInTable": "SOFUTENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO1\_NK | {"Generator\_UserColumnName": "SOFUMEISHO1\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO1\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO2\_NK | {"Generator\_UserColumnName": "SOFUMEISHO2\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO2\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO3\_NK | {"Generator\_UserColumnName": "SOFUMEISHO3\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO3\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO3\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO4\_NK | {"Generator\_UserColumnName": "SOFUMEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUTENPORYAKU\_NK | {"Generator\_UserColumnName": "SOFUTENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUTENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "SOFUTENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "SOFUTENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUTANTO\_NK | {"Generator\_UserColumnName": "SOFUTANTO\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUTANTO\_NK", "Generator\_ColumnPropNameInRow": "SOFUTANTO\_NK", "Generator\_ColumnPropNameInTable": "SOFUTANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORILEVEL | {"Generator\_UserColumnName": "MITSUMORILEVEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORILEVEL", "Generator\_ColumnPropNameInRow": "MITSUMORILEVEL", "Generator\_ColumnPropNameInTable": "MITSUMORILEVELColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORI\_JOUKEN | {"Generator\_UserColumnName": "MITSUMORI\_JOUKEN", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_JOUKEN", "Generator\_ColumnPropNameInRow": "MITSUMORI\_JOUKEN", "Generator\_ColumnPropNameInTable": "MITSUMORI\_JOUKENColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SHINSA\_DT | {"Generator\_UserColumnName": "SHINSA\_DT", "Generator\_ColumnPropNameInRow": "SHINSA\_DT", "Generator\_ColumnVarNameInTable": "columnSHINSA\_DT", "Generator\_ColumnPropNameInTable": "SHINSA\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | SHINSA\_BY | {"Generator\_UserColumnName": "SHINSA\_BY", "Generator\_ColumnPropNameInRow": "SHINSA\_BY", "Generator\_ColumnVarNameInTable": "columnSHINSA\_BY", "Generator\_ColumnPropNameInTable": "SHINSA\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SYOUNIN\_DT | {"Generator\_UserColumnName": "SYOUNIN\_DT", "Generator\_ColumnPropNameInRow": "SYOUNIN\_DT", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_DT", "Generator\_ColumnPropNameInTable": "SYOUNIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | SYOUNIN\_BY | {"Generator\_UserColumnName": "SYOUNIN\_BY", "Generator\_ColumnPropNameInRow": "SYOUNIN\_BY", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_BY", "Generator\_ColumnPropNameInTable": "SYOUNIN\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKAN | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TK\_KIKAN | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | TORITEL\_NO | {"Generator\_UserColumnName": "TORITEL\_NO", "Generator\_ColumnPropNameInRow": "TORITEL\_NO", "Generator\_ColumnVarNameInTable": "columnTORITEL\_NO", "Generator\_ColumnPropNameInTable": "TORITEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | TORIFAX\_NO | {"Generator\_UserColumnName": "TORIFAX\_NO", "Generator\_ColumnPropNameInRow": "TORIFAX\_NO", "Generator\_ColumnVarNameInTable": "columnTORIFAX\_NO", "Generator\_ColumnPropNameInTable": "TORIFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | TORIJIGYOSYO\_NK | {"Generator\_UserColumnName": "TORIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "TORIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnTORIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "TORIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KIKAN | TORISYUTANTOSYA\_NK | {"Generator\_UserColumnName": "TORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "TORISYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnTORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "TORISYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KIKAN | TORIASSISTANT\_NK | {"Generator\_UserColumnName": "TORIASSISTANT\_NK", "Generator\_ColumnPropNameInRow": "TORIASSISTANT\_NK", "Generator\_ColumnVarNameInTable": "columnTORIASSISTANT\_NK", "Generator\_ColumnPropNameInTable": "TORIASSISTANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TK\_KIKAN | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TK\_KIKAN | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | HYOJUNGOKEI\_KIN | {"Generator\_UserColumnName": "HYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUNGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUNGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SIKIRISYOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRISYOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRISYOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SYUSSEINEBIKI\_KIN | {"Generator\_UserColumnName": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInRow": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInTable": "SYUSSEINEBIKI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SIKIRIGOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRIGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRIGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_TK\_KIKAN | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_TK\_KIKAN | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_KIKAN | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KIKAN | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KIKAN | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KIKAN | YUKO\_FLG | {"Generator\_UserColumnName": "YUKO\_FLG", "Generator\_ColumnPropNameInRow": "YUKO\_FLG", "Generator\_ColumnVarNameInTable": "columnYUKO\_FLG", "Generator\_ColumnPropNameInTable": "YUKO\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | TK\_TENPU\_ID | {"Generator\_UserColumnName": "TK\_TENPU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_TENPU\_ID", "Generator\_ColumnPropNameInRow": "TK\_TENPU\_ID", "Generator\_ColumnPropNameInTable": "TK\_TENPU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_TENPU | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | TENPUFILE\_NK | {"Generator\_UserColumnName": "TENPUFILE\_NK", "Generator\_ColumnVarNameInTable": "columnTENPUFILE\_NK", "Generator\_ColumnPropNameInRow": "TENPUFILE\_NK", "Generator\_ColumnPropNameInTable": "TENPUFILE\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_TK\_TENPU | DIRECTORY | {"Generator\_UserColumnName": "DIRECTORY", "Generator\_ColumnVarNameInTable": "columnDIRECTORY", "Generator\_ColumnPropNameInRow": "DIRECTORY", "Generator\_ColumnPropNameInTable": "DIRECTORYColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_TK\_TENPU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_TENPU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_TENPU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | SHONIN\_BY | {"Generator\_UserColumnName": "SHONIN\_BY", "Generator\_ColumnPropNameInRow": "SHONIN\_BY", "Generator\_ColumnVarNameInTable": "columnSHONIN\_BY", "Generator\_ColumnPropNameInTable": "SHONIN\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | SHONIN\_DT | {"Generator\_UserColumnName": "SHONIN\_DT", "Generator\_ColumnPropNameInRow": "SHONIN\_DT", "Generator\_ColumnVarNameInTable": "columnSHONIN\_DT", "Generator\_ColumnPropNameInTable": "SHONIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENPU | SHINSA\_BY | {"Generator\_UserColumnName": "SHINSA\_BY", "Generator\_ColumnPropNameInRow": "SHINSA\_BY", "Generator\_ColumnVarNameInTable": "columnSHINSA\_BY", "Generator\_ColumnPropNameInTable": "SHINSA\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENPU | SHINSA\_DT | {"Generator\_UserColumnName": "SHINSA\_DT", "Generator\_ColumnPropNameInRow": "SHINSA\_DT", "Generator\_ColumnVarNameInTable": "columnSHINSA\_DT", "Generator\_ColumnPropNameInTable": "SHINSA\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UM_MITSUMORI" />
      <xs:field xpath="mstns:RELATION_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UM_TENPU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UM_TENPU" />
      <xs:field xpath="mstns:TENPU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KEIYAKU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KEIYAKU" />
      <xs:field xpath="mstns:TK_KEIYAKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UM_KIHON_MITSUMORI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UM_KIHON_MITSUMORI" />
      <xs:field xpath="mstns:UM_KIHON_MITSUMORI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KIKAN" />
      <xs:field xpath="mstns:TK_KIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_TENPU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_TENPU" />
      <xs:field xpath="mstns:TK_TENPU_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UM_KIHON_MITSUMORI1_MCM_UM_TENPU" ns0:parent="MCM_UM_MITSUMORI" ns0:child="MCM_UM_TENPU" ns0:parentkey="RELATION_ID SHONIN_NO" ns0:childkey="RELATION_ID SHONIN_NO" ns1:Generator_UserRelationName="MCM_UM_KIHON_MITSUMORI1_MCM_UM_TENPU" ns1:Generator_RelationVarName="relationMCM_UM_KIHON_MITSUMORI1_MCM_UM_TENPU" ns1:Generator_UserChildTable="MCM_UM_TENPU" ns1:Generator_UserParentTable="MCM_UM_MITSUMORI" ns1:Generator_ParentPropName="MCM_UM_MITSUMORIRowParent" ns1:Generator_ChildPropName="GetMCM_UM_TENPURows" />
    
```

</details>

### 2006　Mcm2006uDataSet.xsd

[Mcm2006uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_UK\_BRAND | UK\_BRAND\_ID | {"Generator\_UserColumnName": "UK\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UK\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UK\_BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_BRAND | UK\_KIKAN\_ID | {"Generator\_UserColumnName": "UK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_BRAND | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_BRAND | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_BRAND | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_BRAND | DREMOS\_FLG | {"Generator\_UserColumnName": "DREMOS\_FLG", "Generator\_ColumnVarNameInTable": "columnDREMOS\_FLG", "Generator\_ColumnPropNameInRow": "DREMOS\_FLG", "Generator\_ColumnPropNameInTable": "DREMOS\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | REMOTE\_FLG | {"Generator\_UserColumnName": "REMOTE\_FLG", "Generator\_ColumnVarNameInTable": "columnREMOTE\_FLG", "Generator\_ColumnPropNameInRow": "REMOTE\_FLG", "Generator\_ColumnPropNameInTable": "REMOTE\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | REMOTERENRAKUSAKI | {"Generator\_UserColumnName": "REMOTERENRAKUSAKI", "Generator\_ColumnVarNameInTable": "columnREMOTERENRAKUSAKI", "Generator\_ColumnPropNameInRow": "REMOTERENRAKUSAKI", "Generator\_ColumnPropNameInTable": "REMOTERENRAKUSAKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "200"}]] |
| MCM\_UK\_BRAND | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_BRAND | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_BRAND | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_BRAND | SOFT\_FLG | {"Generator\_UserColumnName": "SOFT\_FLG", "Generator\_ColumnVarNameInTable": "columnSOFT\_FLG", "Generator\_ColumnPropNameInRow": "SOFT\_FLG", "Generator\_ColumnPropNameInTable": "SOFT\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_BRAND | DREMOS\_NM | {"Generator\_UserColumnName": "DREMOS\_NM", "Generator\_ColumnVarNameInTable": "columnDREMOS\_NM", "Generator\_ColumnPropNameInRow": "DREMOS\_NM", "Generator\_ColumnPropNameInTable": "DREMOS\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | UK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KIKIKOSEI | UK\_BRAND\_ID | {"Generator\_UserColumnName": "UK\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UK\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UK\_BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKIKOSEI | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | TANI | {"Generator\_UserColumnName": "TANI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnTANI", "Generator\_ColumnPropNameInTable": "TANIColumn", "Generator\_ColumnPropNameInRow": "TANI", "minOccurs": "0"} | [["maxLength", {"value": "10"}]] |
| MCM\_UK\_KIKIKOSEI | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKIKOSEI | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnPropNameInRow": "BIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_KIKIKOSEI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKIKOSEI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKIKOSEI | CONTROLLER\_FLG | {"Generator\_UserColumnName": "CONTROLLER\_FLG", "Generator\_ColumnPropNameInRow": "CONTROLLER\_FLG", "Generator\_ColumnVarNameInTable": "columnCONTROLLER\_FLG", "Generator\_ColumnPropNameInTable": "CONTROLLER\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KIKIKOSEI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_KIKIKOSEI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIKOSEI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | UK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KIKIMEISAI | UK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | SEIZOMAKER\_ID | {"Generator\_UserColumnName": "SEIZOMAKER\_ID", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_ID", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_ID", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKIMEISAI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKIMEISAI | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | KEIYAKUNAIYO | {"Generator\_UserColumnName": "KEIYAKUNAIYO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "KEIYAKUNAIYOColumn", "Generator\_ColumnPropNameInRow": "KEIYAKUNAIYO", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_UK\_KIKIMEISAI | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKIMEISAI | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_KIKIMEISAI | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | DAIFUKUHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "DAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnDAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "DAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "DAIFUKUHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | TENKENKAISU | {"Generator\_UserColumnName": "TENKENKAISU", "Generator\_ColumnVarNameInTable": "columnTENKENKAISU", "Generator\_ColumnPropNameInRow": "TENKENKAISU", "Generator\_ColumnPropNameInTable": "TENKENKAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | TENKENYOBI | {"Generator\_UserColumnName": "TENKENYOBI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnTENKENYOBI", "Generator\_ColumnPropNameInTable": "TENKENYOBIColumn", "Generator\_ColumnPropNameInRow": "TENKENYOBI", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_KIKIMEISAI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_KIKIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnPropNameInRow": "BIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_KIKIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKIMEISAI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | UK\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "UK\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KOTAIMEISAI | UK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | KOTAIKANRI\_ID | {"Generator\_UserColumnName": "KOTAIKANRI\_ID", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_ID", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_ID", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | KOTAI\_NK | {"Generator\_UserColumnName": "KOTAI\_NK", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKOTAI\_NK", "Generator\_ColumnPropNameInTable": "KOTAI\_NKColumn", "Generator\_ColumnPropNameInRow": "KOTAI\_NK", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_UK\_KOTAIMEISAI | SERIAL\_NO | {"Generator\_UserColumnName": "SERIAL\_NO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnSERIAL\_NO", "Generator\_ColumnPropNameInTable": "SERIAL\_NOColumn", "Generator\_ColumnPropNameInRow": "SERIAL\_NO", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KOTAIMEISAI | ITIJINONYU\_DT | {"Generator\_UserColumnName": "ITIJINONYU\_DT", "Generator\_ColumnVarNameInTable": "columnITIJINONYU\_DT", "Generator\_ColumnPropNameInRow": "ITIJINONYU\_DT", "Generator\_ColumnPropNameInTable": "ITIJINONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | SETCHIBASYO | {"Generator\_UserColumnName": "SETCHIBASYO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnSETCHIBASYO", "Generator\_ColumnPropNameInTable": "SETCHIBASYOColumn", "Generator\_ColumnPropNameInRow": "SETCHIBASYO", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_UK\_KOTAIMEISAI | KEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "KEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | ENCHOKEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "ENCHOKEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnPropNameInRow": "BIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_KOTAIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KOTAIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KOTAIMEISAI | UK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KOTAIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KOTAIMEISAI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | UK\_SEIBAN\_ID | {"Generator\_UserColumnName": "UK\_SEIBAN\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_SEIBAN\_ID", "Generator\_ColumnPropNameInRow": "UK\_SEIBAN\_ID", "Generator\_ColumnPropNameInTable": "UK\_SEIBAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_SEIBAN | UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | HARD\_SEIBAN | {"Generator\_UserColumnName": "HARD\_SEIBAN", "Generator\_ColumnVarNameInTable": "columnHARD\_SEIBAN", "Generator\_ColumnPropNameInRow": "HARD\_SEIBAN", "Generator\_ColumnPropNameInTable": "HARD\_SEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UK\_SEIBAN | SOFT\_SEIBAN | {"Generator\_UserColumnName": "SOFT\_SEIBAN", "Generator\_ColumnVarNameInTable": "columnSOFT\_SEIBAN", "Generator\_ColumnPropNameInRow": "SOFT\_SEIBAN", "Generator\_ColumnPropNameInTable": "SOFT\_SEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UK\_SEIBAN | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_SEIBAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | KAKUNIN\_DT | {"Generator\_UserColumnName": "KAKUNIN\_DT", "Generator\_ColumnVarNameInTable": "columnKAKUNIN\_DT", "Generator\_ColumnPropNameInRow": "KAKUNIN\_DT", "Generator\_ColumnPropNameInTable": "KAKUNIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | KAKUNIN\_BY | {"Generator\_UserColumnName": "KAKUNIN\_BY", "Generator\_ColumnVarNameInTable": "columnKAKUNIN\_BY", "Generator\_ColumnPropNameInRow": "KAKUNIN\_BY", "Generator\_ColumnPropNameInTable": "KAKUNIN\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | KAKUNIN\_KBN | {"Generator\_UserColumnName": "KAKUNIN\_KBN", "Generator\_ColumnPropNameInRow": "KAKUNIN\_KBN", "Generator\_ColumnVarNameInTable": "columnKAKUNIN\_KBN", "Generator\_ColumnPropNameInTable": "KAKUNIN\_KBNColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_SEIBAN | KAKUNINIRAI\_BY | {"Generator\_UserColumnName": "KAKUNINIRAI\_BY", "Generator\_ColumnPropNameInRow": "KAKUNINIRAI\_BY", "Generator\_ColumnVarNameInTable": "columnKAKUNINIRAI\_BY", "Generator\_ColumnPropNameInTable": "KAKUNINIRAI\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | KAKUNINIRAI\_DT | {"Generator\_UserColumnName": "KAKUNINIRAI\_DT", "Generator\_ColumnPropNameInRow": "KAKUNINIRAI\_DT", "Generator\_ColumnVarNameInTable": "columnKAKUNINIRAI\_DT", "Generator\_ColumnPropNameInTable": "KAKUNINIRAI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_TENKEN | UK\_TENKEN\_ID | {"Generator\_UserColumnName": "UK\_TENKEN\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_TENKEN\_ID", "Generator\_ColumnPropNameInRow": "UK\_TENKEN\_ID", "Generator\_ColumnPropNameInTable": "UK\_TENKEN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_TENKEN | UK\_KIKAN\_ID | {"Generator\_UserColumnName": "UK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_TENKEN | NAIYO | {"Generator\_UserColumnName": "NAIYO", "Generator\_ColumnVarNameInTable": "columnNAIYO", "Generator\_ColumnPropNameInRow": "NAIYO", "Generator\_ColumnPropNameInTable": "NAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_TENKEN | M01 | {"Generator\_UserColumnName": "M01", "Generator\_ColumnVarNameInTable": "columnM01", "Generator\_ColumnPropNameInRow": "M01", "Generator\_ColumnPropNameInTable": "M01Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M02 | {"Generator\_UserColumnName": "M02", "Generator\_ColumnVarNameInTable": "columnM02", "Generator\_ColumnPropNameInRow": "M02", "Generator\_ColumnPropNameInTable": "M02Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M03 | {"Generator\_UserColumnName": "M03", "Generator\_ColumnVarNameInTable": "columnM03", "Generator\_ColumnPropNameInRow": "M03", "Generator\_ColumnPropNameInTable": "M03Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M04 | {"Generator\_UserColumnName": "M04", "Generator\_ColumnVarNameInTable": "columnM04", "Generator\_ColumnPropNameInRow": "M04", "Generator\_ColumnPropNameInTable": "M04Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M05 | {"Generator\_UserColumnName": "M05", "Generator\_ColumnVarNameInTable": "columnM05", "Generator\_ColumnPropNameInRow": "M05", "Generator\_ColumnPropNameInTable": "M05Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M06 | {"Generator\_UserColumnName": "M06", "Generator\_ColumnVarNameInTable": "columnM06", "Generator\_ColumnPropNameInRow": "M06", "Generator\_ColumnPropNameInTable": "M06Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M07 | {"Generator\_UserColumnName": "M07", "Generator\_ColumnVarNameInTable": "columnM07", "Generator\_ColumnPropNameInRow": "M07", "Generator\_ColumnPropNameInTable": "M07Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M08 | {"Generator\_UserColumnName": "M08", "Generator\_ColumnVarNameInTable": "columnM08", "Generator\_ColumnPropNameInRow": "M08", "Generator\_ColumnPropNameInTable": "M08Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M09 | {"Generator\_UserColumnName": "M09", "Generator\_ColumnVarNameInTable": "columnM09", "Generator\_ColumnPropNameInRow": "M09", "Generator\_ColumnPropNameInTable": "M09Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M10 | {"Generator\_UserColumnName": "M10", "Generator\_ColumnVarNameInTable": "columnM10", "Generator\_ColumnPropNameInRow": "M10", "Generator\_ColumnPropNameInTable": "M10Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M11 | {"Generator\_UserColumnName": "M11", "Generator\_ColumnVarNameInTable": "columnM11", "Generator\_ColumnPropNameInRow": "M11", "Generator\_ColumnPropNameInTable": "M11Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | M12 | {"Generator\_UserColumnName": "M12", "Generator\_ColumnVarNameInTable": "columnM12", "Generator\_ColumnPropNameInRow": "M12", "Generator\_ColumnPropNameInTable": "M12Column", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_TENKEN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_TENKEN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_TENKEN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_TENKEN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_TENKEN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | AUTO\_FLG | {"Generator\_UserColumnName": "AUTO\_FLG", "Generator\_ColumnPropNameInRow": "AUTO\_FLG", "Generator\_ColumnVarNameInTable": "columnAUTO\_FLG", "Generator\_ColumnPropNameInTable": "AUTO\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | SHOKAI\_KEIYAKU\_DT | {"Generator\_UserColumnName": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnSHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "SHOKAI\_KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnJIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | UK\_KIKAN\_ID | {"Generator\_UserColumnName": "UK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KIKAN | UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | HOSYU\_GKIN | {"Generator\_UserColumnName": "HOSYU\_GKIN", "Generator\_ColumnPropNameInRow": "HOSYU\_GKIN", "Generator\_ColumnVarNameInTable": "columnHOSYU\_GKIN", "Generator\_ColumnPropNameInTable": "HOSYU\_GKINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_UK\_KIKAN | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKAN | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKAN | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKAN | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_UK\_KIKAN | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_KIKAN | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_UK\_KIKAN | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_UK\_KIKAN | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKAN | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_KIKAN | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_KIKAN | IRAITENPO\_ID | {"Generator\_UserColumnName": "IRAITENPO\_ID", "Generator\_ColumnPropNameInRow": "IRAITENPO\_ID", "Generator\_ColumnVarNameInTable": "columnIRAITENPO\_ID", "Generator\_ColumnPropNameInTable": "IRAITENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | IRAIMEISHO1\_NK | {"Generator\_UserColumnName": "IRAIMEISHO1\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "IRAIMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO1\_NKColumn", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO1\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKAN | IRAIMEISHO2\_NK | {"Generator\_UserColumnName": "IRAIMEISHO2\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "IRAIMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO2\_NKColumn", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO2\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKAN | IRAIMEISHO3\_NK | {"Generator\_UserColumnName": "IRAIMEISHO3\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "IRAIMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO3\_NKColumn", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO3\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKAN | IRAIMEISHO4\_NK | {"Generator\_UserColumnName": "IRAIMEISHO4\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "IRAIMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO4\_NKColumn", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO4\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKAN | IRAITENPORYAKU\_NK | {"Generator\_UserColumnName": "IRAITENPORYAKU\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "IRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "IRAITENPORYAKU\_NKColumn", "Generator\_ColumnVarNameInTable": "columnIRAITENPORYAKU\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKAN | IRAITANTO\_NK | {"Generator\_UserColumnName": "IRAITANTO\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "IRAITANTO\_NK", "Generator\_ColumnPropNameInTable": "IRAITANTO\_NKColumn", "Generator\_ColumnVarNameInTable": "columnIRAITANTO\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_KIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnVarNameInTable": "columnBIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_KIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KIKAN | YUKO\_FLG | {"Generator\_UserColumnName": "YUKO\_FLG", "Generator\_ColumnVarNameInTable": "columnYUKO\_FLG", "Generator\_ColumnPropNameInRow": "YUKO\_FLG", "Generator\_ColumnPropNameInTable": "YUKO\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_MITSUMORI\_BRAND | UK\_BRAND\_ID | {"Generator\_UserColumnName": "UK\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UK\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UK\_BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_MITSUMORI\_BRAND | UM\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_MITSUMORI\_BRAND | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_MITSUMORI\_BRAND | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_MITSUMORI\_BRAND | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_MITSUMORI\_BRAND | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_MITSUMORI\_BRAND | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_MITSUMORI\_BRAND | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_BRAND\_KOSEI | SYSTEMSEKKEI\_KIN | {"Generator\_UserColumnName": "SYSTEMSEKKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYSTEMSEKKEI\_KIN", "Generator\_ColumnPropNameInRow": "SYSTEMSEKKEI\_KIN", "Generator\_ColumnPropNameInTable": "SYSTEMSEKKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | KIHONSEKKEI\_KIN | {"Generator\_UserColumnName": "KIHONSEKKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnKIHONSEKKEI\_KIN", "Generator\_ColumnPropNameInRow": "KIHONSEKKEI\_KIN", "Generator\_ColumnPropNameInTable": "KIHONSEKKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | PROGRAMSAKUSEI\_KIN | {"Generator\_UserColumnName": "PROGRAMSAKUSEI\_KIN", "Generator\_ColumnVarNameInTable": "columnPROGRAMSAKUSEI\_KIN", "Generator\_ColumnPropNameInRow": "PROGRAMSAKUSEI\_KIN", "Generator\_ColumnPropNameInTable": "PROGRAMSAKUSEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | NONYU\_DT | {"Generator\_UserColumnName": "NONYU\_DT", "Generator\_ColumnVarNameInTable": "columnNONYU\_DT", "Generator\_ColumnPropNameInRow": "NONYU\_DT", "Generator\_ColumnPropNameInTable": "NONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | TEKKYO\_DT | {"Generator\_UserColumnName": "TEKKYO\_DT", "Generator\_ColumnVarNameInTable": "columnTEKKYO\_DT", "Generator\_ColumnPropNameInRow": "TEKKYO\_DT", "Generator\_ColumnPropNameInTable": "TEKKYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | HOSYUSYUSOKU\_DT | {"Generator\_UserColumnName": "HOSYUSYUSOKU\_DT", "Generator\_ColumnVarNameInTable": "columnHOSYUSYUSOKU\_DT", "Generator\_ColumnPropNameInRow": "HOSYUSYUSOKU\_DT", "Generator\_ColumnPropNameInTable": "HOSYUSYUSOKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | REMOTE\_FLG | {"Generator\_UserColumnName": "REMOTE\_FLG", "Generator\_ColumnVarNameInTable": "columnREMOTE\_FLG", "Generator\_ColumnPropNameInRow": "REMOTE\_FLG", "Generator\_ColumnPropNameInTable": "REMOTE\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | REMOTERENRAKUSAKI | {"Generator\_UserColumnName": "REMOTERENRAKUSAKI", "Generator\_ColumnVarNameInTable": "columnREMOTERENRAKUSAKI", "Generator\_ColumnPropNameInRow": "REMOTERENRAKUSAKI", "Generator\_ColumnPropNameInTable": "REMOTERENRAKUSAKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "200"}]] |
| MCM\_MA\_BRAND\_KOSEI | DREMOS\_FLG | {"Generator\_UserColumnName": "DREMOS\_FLG", "Generator\_ColumnVarNameInTable": "columnDREMOS\_FLG", "Generator\_ColumnPropNameInRow": "DREMOS\_FLG", "Generator\_ColumnPropNameInTable": "DREMOS\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_BRAND\_KOSEI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_BRAND\_KOSEI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_BRAND\_KOSEI | DREMOS\_NM | {"Generator\_UserColumnName": "DREMOS\_NM", "Generator\_ColumnVarNameInTable": "columnDREMOS\_NM", "Generator\_ColumnPropNameInRow": "DREMOS\_NM", "Generator\_ColumnPropNameInTable": "DREMOS\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_BRAND" />
      <xs:field xpath="mstns:UK_BRAND_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_KIKIKOSEI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_KIKIKOSEI" />
      <xs:field xpath="mstns:UK_KIKIKOSEI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_KIKIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_KIKIMEISAI" />
      <xs:field xpath="mstns:UK_KIKIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_KOTAIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_KOTAIMEISAI" />
      <xs:field xpath="mstns:UK_KOTAIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_SEIBAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_SEIBAN" />
      <xs:field xpath="mstns:UK_SEIBAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_TENKEN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_TENKEN" />
      <xs:field xpath="mstns:UK_TENKEN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_KEIYAKU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_KEIYAKU" />
      <xs:field xpath="mstns:UK_KEIYAKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_KIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_KIKAN" />
      <xs:field xpath="mstns:UK_KIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_MITSUMORI_BRAND_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_MITSUMORI_BRAND" />
      <xs:field xpath="mstns:UK_BRAND_ID" />
      <xs:field xpath="mstns:UM_MITSUMORI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_BRAND_KOSEI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_BRAND_KOSEI" />
      <xs:field xpath="mstns:BRANDKOSEI_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_KEIYAKU_MCM_UK_KIKAN" ns0:parent="MCM_UK_KEIYAKU" ns0:child="MCM_UK_KIKAN" ns0:parentkey="UK_KEIYAKU_ID" ns0:childkey="UK_KEIYAKU_ID" ns1:Generator_UserRelationName="MCM_UK_KEIYAKU_MCM_UK_KIKAN" ns1:Generator_RelationVarName="relationMCM_UK_KEIYAKU_MCM_UK_KIKAN" ns1:Generator_UserChildTable="MCM_UK_KIKAN" ns1:Generator_UserParentTable="MCM_UK_KEIYAKU" ns1:Generator_ParentPropName="MCM_UK_KEIYAKURow" ns1:Generator_ChildPropName="GetMCM_UK_KIKANRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_KIKAN_MCM_UK_BRAND" ns0:parent="MCM_UK_KIKAN" ns0:child="MCM_UK_BRAND" ns0:parentkey="UK_KIKAN_ID" ns0:childkey="UK_KIKAN_ID" ns1:Generator_UserRelationName="MCM_UK_KIKAN_MCM_UK_BRAND" ns1:Generator_RelationVarName="relationMCM_UK_KIKAN_MCM_UK_BRAND" ns1:Generator_UserChildTable="MCM_UK_BRAND" ns1:Generator_UserParentTable="MCM_UK_KIKAN" ns1:Generator_ParentPropName="MCM_UK_KIKANRow" ns1:Generator_ChildPropName="GetMCM_UK_BRANDRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_KIKAN_MCM_UK_TENKEN" ns0:parent="MCM_UK_KIKAN" ns0:child="MCM_UK_TENKEN" ns0:parentkey="UK_KIKAN_ID" ns0:childkey="UK_KIKAN_ID" ns1:Generator_UserRelationName="MCM_UK_KIKAN_MCM_UK_TENKEN" ns1:Generator_RelationVarName="relationMCM_UK_KIKAN_MCM_UK_TENKEN" ns1:Generator_UserChildTable="MCM_UK_TENKEN" ns1:Generator_UserParentTable="MCM_UK_KIKAN" ns1:Generator_ParentPropName="MCM_UK_KIKANRow" ns1:Generator_ChildPropName="GetMCM_UK_TENKENRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_KEIYAKU_MCM_UK_SEIBAN" ns0:parent="MCM_UK_KEIYAKU" ns0:child="MCM_UK_SEIBAN" ns0:parentkey="UK_KEIYAKU_ID" ns0:childkey="UK_KEIYAKU_ID" ns1:Generator_UserRelationName="MCM_UK_KEIYAKU_MCM_UK_SEIBAN" ns1:Generator_RelationVarName="relationMCM_UK_KEIYAKU_MCM_UK_SEIBAN" ns1:Generator_UserChildTable="MCM_UK_SEIBAN" ns1:Generator_UserParentTable="MCM_UK_KEIYAKU" ns1:Generator_ParentPropName="MCM_UK_KEIYAKURow" ns1:Generator_ChildPropName="GetMCM_UK_SEIBANRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_BRAND_MCM_UK_KIKIKOSEI" ns0:parent="MCM_UK_BRAND" ns0:child="MCM_UK_KIKIKOSEI" ns0:parentkey="UK_BRAND_ID" ns0:childkey="UK_BRAND_ID" ns1:Generator_UserRelationName="MCM_UK_BRAND_MCM_UK_KIKIKOSEI" ns1:Generator_RelationVarName="relationMCM_UK_BRAND_MCM_UK_KIKIKOSEI" ns1:Generator_UserChildTable="MCM_UK_KIKIKOSEI" ns1:Generator_UserParentTable="MCM_UK_BRAND" ns1:Generator_ParentPropName="MCM_UK_BRANDRow" ns1:Generator_ChildPropName="GetMCM_UK_KIKIKOSEIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_KIKIMEISAI_MCM_UK_KOTAIMEISAI" ns0:parent="MCM_UK_KIKIMEISAI" ns0:child="MCM_UK_KOTAIMEISAI" ns0:parentkey="UK_KIKIMEISAI_ID" ns0:childkey="UK_KIKIMEISAI_ID" ns1:Generator_UserRelationName="MCM_UK_KIKIMEISAI_MCM_UK_KOTAIMEISAI" ns1:Generator_RelationVarName="relationMCM_UK_KIKIMEISAI_MCM_UK_KOTAIMEISAI" ns1:Generator_UserChildTable="MCM_UK_KOTAIMEISAI" ns1:Generator_UserParentTable="MCM_UK_KIKIMEISAI" ns1:Generator_ParentPropName="MCM_UK_KIKIMEISAIRow" ns1:Generator_ChildPropName="GetMCM_UK_KOTAIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UK_KIKIKOSEI_MCM_UK_KIKIMEISAI" ns0:parent="MCM_UK_KIKIKOSEI" ns0:child="MCM_UK_KIKIMEISAI" ns0:parentkey="UK_KIKIKOSEI_ID" ns0:childkey="UK_KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_UK_KIKIKOSEI_MCM_UK_KIKIMEISAI" ns1:Generator_RelationVarName="relationMCM_UK_KIKIKOSEI_MCM_UK_KIKIMEISAI" ns1:Generator_UserChildTable="MCM_UK_KIKIMEISAI" ns1:Generator_UserParentTable="MCM_UK_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_UK_KIKIKOSEIRow" ns1:Generator_ChildPropName="GetMCM_UK_KIKIMEISAIRows" />
    
```

</details>

### 2007　Mcm2007u2DataSet.xsd

[Mcm2007u2DataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2DataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_MA\_BRAND\_KOSEI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_BRAND\_KOSEI | REMOTE\_FLG | {"Generator\_UserColumnName": "REMOTE\_FLG", "Generator\_ColumnVarNameInTable": "columnREMOTE\_FLG", "Generator\_ColumnPropNameInRow": "REMOTE\_FLG", "Generator\_ColumnPropNameInTable": "REMOTE\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | REMOTERENRAKUSAKI | {"Generator\_UserColumnName": "REMOTERENRAKUSAKI", "Generator\_ColumnVarNameInTable": "columnREMOTERENRAKUSAKI", "Generator\_ColumnPropNameInRow": "REMOTERENRAKUSAKI", "Generator\_ColumnPropNameInTable": "REMOTERENRAKUSAKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "200"}]] |
| MCM\_MA\_BRAND\_KOSEI | DREMOS\_FLG | {"Generator\_UserColumnName": "DREMOS\_FLG", "Generator\_ColumnVarNameInTable": "columnDREMOS\_FLG", "Generator\_ColumnPropNameInRow": "DREMOS\_FLG", "Generator\_ColumnPropNameInTable": "DREMOS\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | SOFTHOSYUHOHO | {"Generator\_UserColumnName": "SOFTHOSYUHOHO", "Generator\_ColumnPropNameInRow": "SOFTHOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnSOFTHOSYUHOHO", "Generator\_ColumnPropNameInTable": "SOFTHOSYUHOHOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | SOFT\_FLG | {"Generator\_UserColumnName": "SOFT\_FLG", "Generator\_ColumnPropNameInRow": "SOFT\_FLG", "Generator\_ColumnVarNameInTable": "columnSOFT\_FLG", "Generator\_ColumnPropNameInTable": "SOFT\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | UM\_KIHON\_BRAND\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOSEI | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn"} | [["maxLength", {"value": "100"}]] |
| MCM\_MA\_KIKIKOSEI | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOSEI | TANI | {"Generator\_UserColumnName": "TANI", "Generator\_ColumnVarNameInTable": "columnTANI", "Generator\_ColumnPropNameInRow": "TANI", "Generator\_ColumnPropNameInTable": "TANIColumn"} | [["maxLength", {"value": "10"}]] |
| MCM\_MA\_KIKIKOSEI | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_KIKIKOSEI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_KIKIKOSEI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOSEI | UM\_KIHON\_BRAND\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | UM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | MAD\_HYOJIJUN | {"Generator\_UserColumnName": "MAD\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAD\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAD\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAD\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | UK\_KIKIKOSEI\_ID | {"Caption": "UK\_KIKIKOSEI\_ID1", "Generator\_UserColumnName": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | CONTROLLER\_FLG | {"Generator\_UserColumnName": "CONTROLLER\_FLG", "Generator\_ColumnVarNameInTable": "columnCONTROLLER\_FLG", "Generator\_ColumnPropNameInRow": "CONTROLLER\_FLG", "Generator\_ColumnPropNameInTable": "CONTROLLER\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOSEI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOSEI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KOTAIKANRI\_ID | {"Generator\_UserColumnName": "KOTAIKANRI\_ID", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_ID", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_ID", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KOTAI\_NK | {"Generator\_UserColumnName": "KOTAI\_NK", "Generator\_ColumnVarNameInTable": "columnKOTAI\_NK", "Generator\_ColumnPropNameInRow": "KOTAI\_NK", "Generator\_ColumnPropNameInTable": "KOTAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | SERIAL\_NO | {"Generator\_UserColumnName": "SERIAL\_NO", "Generator\_ColumnVarNameInTable": "columnSERIAL\_NO", "Generator\_ColumnPropNameInRow": "SERIAL\_NO", "Generator\_ColumnPropNameInTable": "SERIAL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | ITIJINONYU\_DT | {"Generator\_UserColumnName": "ITIJINONYU\_DT", "Generator\_ColumnVarNameInTable": "columnITIJINONYU\_DT", "Generator\_ColumnPropNameInRow": "ITIJINONYU\_DT", "Generator\_ColumnPropNameInTable": "ITIJINONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | SETCHIBASYO | {"Generator\_UserColumnName": "SETCHIBASYO", "Generator\_ColumnVarNameInTable": "columnSETCHIBASYO", "Generator\_ColumnPropNameInRow": "SETCHIBASYO", "Generator\_ColumnPropNameInTable": "SETCHIBASYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | TEKKYO\_DT | {"Generator\_UserColumnName": "TEKKYO\_DT", "Generator\_ColumnVarNameInTable": "columnTEKKYO\_DT", "Generator\_ColumnPropNameInRow": "TEKKYO\_DT", "Generator\_ColumnPropNameInTable": "TEKKYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | SET\_NM | {"Generator\_UserColumnName": "SET\_NM", "Generator\_ColumnPropNameInRow": "SET\_NM", "Generator\_ColumnVarNameInTable": "columnSET\_NM", "Generator\_ColumnPropNameInTable": "SET\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIKOTAIKANRI | MAG\_ENCHOKEIYAKUKIGEN\_DT | {"Generator\_UserColumnName": "MAG\_ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInRow": "MAG\_ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnVarNameInTable": "columnMAG\_ENCHOKEIYAKUKIGEN\_DT", "Generator\_ColumnPropNameInTable": "MAG\_ENCHOKEIYAKUKIGEN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UM\_KIHON\_BRAND\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UM\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "UM\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UK\_KOTAIMEISAI\_ID | {"Generator\_UserColumnName": "UK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KOTAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KOTAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KOTAIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | MAD\_HYOJIJUN | {"Generator\_UserColumnName": "MAD\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAD\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAD\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAD\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | MAF\_HYOJIJUN | {"Generator\_UserColumnName": "MAF\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAF\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAF\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAF\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIKOTAIKANRI | MAG\_HYOJIJUN | {"Generator\_UserColumnName": "MAG\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAG\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAG\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAG\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIKOSEI\_ID | {"Generator\_UserColumnName": "KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | HYOJIJUN | {"Generator\_UserColumnName": "HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnHYOJIJUN", "Generator\_ColumnPropNameInRow": "HYOJIJUN", "Generator\_ColumnPropNameInTable": "HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_KIKIMEISAI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | SEIZOMAKER\_ID | {"Generator\_UserColumnName": "SEIZOMAKER\_ID", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_ID", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_ID", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_KIKIMEISAI | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | UM\_KIHON\_BRAND\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | UM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | UM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | UK\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | UK\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UK\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UK\_KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | MITSUMORICHECK\_FLG | {"Generator\_UserColumnName": "MITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInRow": "MITSUMORICHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnMITSUMORICHECK\_FLG", "Generator\_ColumnPropNameInTable": "MITSUMORICHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | CHECK\_FLG\_OLD | {"Generator\_UserColumnName": "CHECK\_FLG\_OLD", "Generator\_ColumnPropNameInRow": "CHECK\_FLG\_OLD", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG\_OLD", "Generator\_ColumnPropNameInTable": "CHECK\_FLG\_OLDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | MAE\_HYOJIJUN | {"Generator\_UserColumnName": "MAE\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAE\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAE\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAE\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | MAF\_HYOJIJUN | {"Generator\_UserColumnName": "MAF\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAF\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAF\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAF\_HYOJIJUNColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_KIKIMEISAI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | KEIYAKUNAIYO | {"Generator\_UserColumnName": "KEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "KEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "KEIYAKUNAIYOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | DAIFUKUHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "DAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnDAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "DAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "DAIFUKUHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TENKENKAISU | {"Generator\_UserColumnName": "TENKENKAISU", "Generator\_ColumnVarNameInTable": "columnTENKENKAISU", "Generator\_ColumnPropNameInRow": "TENKENKAISU", "Generator\_ColumnPropNameInTable": "TENKENKAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | TENKENYOBI | {"Generator\_UserColumnName": "TENKENYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENYOBI", "Generator\_ColumnPropNameInRow": "TENKENYOBI", "Generator\_ColumnPropNameInTable": "TENKENYOBIColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | MAD\_HYOJIJUN | {"Generator\_UserColumnName": "MAD\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAD\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAD\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAD\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_KIKIMEISAI | KOTAIKANRI\_FLG | {"Generator\_UserColumnName": "KOTAIKANRI\_FLG", "Generator\_ColumnPropNameInRow": "KOTAIKANRI\_FLG", "Generator\_ColumnVarNameInTable": "columnKOTAIKANRI\_FLG", "Generator\_ColumnPropNameInTable": "KOTAIKANRI\_FLGColumn", "type": "xs:decimal"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_BRAND_KOSEI" />
      <xs:field xpath="mstns:BRANDKOSEI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_KIKIKOSEI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_KIKIKOSEI" />
      <xs:field xpath="mstns:KIKIKOSEI_ID" />
      <xs:field xpath="mstns:BRANDKOSEI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_KIKIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_KIKIMEISAI" />
      <xs:field xpath="mstns:KIKIMEISAI_ID" />
      <xs:field xpath="mstns:SEIZOMAKER_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_MA_BRAND_KOSEI_MCM_MA_KIKIKOSEI" ns0:parent="MCM_MA_BRAND_KOSEI" ns0:child="MCM_MA_KIKIKOSEI" ns0:parentkey="BRANDKOSEI_ID" ns0:childkey="BRANDKOSEI_ID" ns1:Generator_UserRelationName="MCM_MA_BRAND_KOSEI_MCM_MA_KIKIKOSEI" ns1:Generator_RelationVarName="relationMCM_MA_BRAND_KOSEI_MCM_MA_KIKIKOSEI" ns1:Generator_UserChildTable="MCM_MA_KIKIKOSEI" ns1:Generator_UserParentTable="MCM_MA_BRAND_KOSEI" ns1:Generator_ParentPropName="MCM_MA_BRAND_KOSEIRow" ns1:Generator_ChildPropName="GetMCM_MA_KIKIKOSEIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI" ns0:parent="MCM_MA_KIKIKOSEI" ns0:child="MCM_MA_KIKIMEISAI" ns0:parentkey="BRANDKOSEI_ID KIKIKOSEI_ID" ns0:childkey="BRANDKOSEI_ID KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI" ns1:Generator_RelationVarName="relationMCM_MA_KIKIKOSEI_MCM_MA_KIKIMEISAI" ns1:Generator_UserChildTable="MCM_MA_KIKIMEISAI" ns1:Generator_UserParentTable="MCM_MA_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_MA_KIKIKOSEIRowParent" ns1:Generator_ChildPropName="GetMCM_MA_KIKIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_MA_KIKIKOSEI_MCM_MA_KIKIKOTAIKANRI" ns0:parent="MCM_MA_KIKIKOSEI" ns0:child="MCM_MA_KIKIKOTAIKANRI" ns0:parentkey="BRANDKOSEI_ID KIKIKOSEI_ID" ns0:childkey="BRANDKOSEI_ID KIKIKOSEI_ID" ns1:Generator_UserRelationName="MCM_MA_KIKIKOSEI_MCM_MA_KIKIKOTAIKANRI" ns1:Generator_RelationVarName="relationMCM_MA_KIKIKOSEI_MCM_MA_KIKIKOTAIKANRI" ns1:Generator_UserChildTable="MCM_MA_KIKIKOTAIKANRI" ns1:Generator_UserParentTable="MCM_MA_KIKIKOSEI" ns1:Generator_ParentPropName="MCM_MA_KIKIKOSEIRowParent" ns1:Generator_ChildPropName="GetMCM_MA_KIKIKOTAIKANRIRows" />
    
```

</details>

### 2007　Mcm2007uDataSet.xsd

[Mcm2007uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_UM\_KIHON\_BRAND | UM\_KIHON\_BRAND\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_KIHON\_BRAND | UM\_KIHON\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_MITSUMORI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_BRAND | BRANDKOSEI\_ID | {"Generator\_UserColumnName": "BRANDKOSEI\_ID", "Generator\_ColumnPropNameInRow": "BRANDKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnBRANDKOSEI\_ID", "Generator\_ColumnPropNameInTable": "BRANDKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_BRAND | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIHON\_BRAND | SOFTHOSYUHOHO | {"Generator\_UserColumnName": "SOFTHOSYUHOHO", "Generator\_ColumnPropNameInRow": "SOFTHOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnSOFTHOSYUHOHO", "Generator\_ColumnPropNameInTable": "SOFTHOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIHON\_BRAND | DREMOS\_FLG | {"Generator\_UserColumnName": "DREMOS\_FLG", "Generator\_ColumnPropNameInRow": "DREMOS\_FLG", "Generator\_ColumnVarNameInTable": "columnDREMOS\_FLG", "Generator\_ColumnPropNameInTable": "DREMOS\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | REMOTE\_FLG | {"Generator\_UserColumnName": "REMOTE\_FLG", "Generator\_ColumnPropNameInRow": "REMOTE\_FLG", "Generator\_ColumnVarNameInTable": "columnREMOTE\_FLG", "Generator\_ColumnPropNameInTable": "REMOTE\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | SYSTEMSEKKEI\_KIN | {"Generator\_UserColumnName": "SYSTEMSEKKEI\_KIN", "Generator\_ColumnPropNameInRow": "SYSTEMSEKKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYSTEMSEKKEI\_KIN", "Generator\_ColumnPropNameInTable": "SYSTEMSEKKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | KIHONSEKKEI\_KIN | {"Generator\_UserColumnName": "KIHONSEKKEI\_KIN", "Generator\_ColumnPropNameInRow": "KIHONSEKKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnKIHONSEKKEI\_KIN", "Generator\_ColumnPropNameInTable": "KIHONSEKKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | PROGRAMSAKUSEI\_KIN | {"Generator\_UserColumnName": "PROGRAMSAKUSEI\_KIN", "Generator\_ColumnPropNameInRow": "PROGRAMSAKUSEI\_KIN", "Generator\_ColumnVarNameInTable": "columnPROGRAMSAKUSEI\_KIN", "Generator\_ColumnPropNameInTable": "PROGRAMSAKUSEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | SYSTEMSUPPORT\_KIN | {"Generator\_UserColumnName": "SYSTEMSUPPORT\_KIN", "Generator\_ColumnPropNameInRow": "SYSTEMSUPPORT\_KIN", "Generator\_ColumnVarNameInTable": "columnSYSTEMSUPPORT\_KIN", "Generator\_ColumnPropNameInTable": "SYSTEMSUPPORT\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | DTSSUPPORT\_KIN | {"Generator\_UserColumnName": "DTSSUPPORT\_KIN", "Generator\_ColumnPropNameInRow": "DTSSUPPORT\_KIN", "Generator\_ColumnVarNameInTable": "columnDTSSUPPORT\_KIN", "Generator\_ColumnPropNameInTable": "DTSSUPPORT\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | DAIFUKUGIJUTSU\_KIN | {"Generator\_UserColumnName": "DAIFUKUGIJUTSU\_KIN", "Generator\_ColumnPropNameInRow": "DAIFUKUGIJUTSU\_KIN", "Generator\_ColumnVarNameInTable": "columnDAIFUKUGIJUTSU\_KIN", "Generator\_ColumnPropNameInTable": "DAIFUKUGIJUTSU\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | SOFTHOSHU\_KIN | {"Generator\_UserColumnName": "SOFTHOSHU\_KIN", "Generator\_ColumnPropNameInRow": "SOFTHOSHU\_KIN", "Generator\_ColumnVarNameInTable": "columnSOFTHOSHU\_KIN", "Generator\_ColumnPropNameInTable": "SOFTHOSHU\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HOSEISYSTEMSUPPORT\_KIN | {"Generator\_UserColumnName": "HOSEISYSTEMSUPPORT\_KIN", "Generator\_ColumnPropNameInRow": "HOSEISYSTEMSUPPORT\_KIN", "Generator\_ColumnVarNameInTable": "columnHOSEISYSTEMSUPPORT\_KIN", "Generator\_ColumnPropNameInTable": "HOSEISYSTEMSUPPORT\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HOSEIDTSSUPPORT\_KIN | {"Generator\_UserColumnName": "HOSEIDTSSUPPORT\_KIN", "Generator\_ColumnPropNameInRow": "HOSEIDTSSUPPORT\_KIN", "Generator\_ColumnVarNameInTable": "columnHOSEIDTSSUPPORT\_KIN", "Generator\_ColumnPropNameInTable": "HOSEIDTSSUPPORT\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HOSEIDAIFUKUGIJUTSU\_KIN | {"Generator\_UserColumnName": "HOSEIDAIFUKUGIJUTSU\_KIN", "Generator\_ColumnPropNameInRow": "HOSEIDAIFUKUGIJUTSU\_KIN", "Generator\_ColumnVarNameInTable": "columnHOSEIDAIFUKUGIJUTSU\_KIN", "Generator\_ColumnPropNameInTable": "HOSEIDAIFUKUGIJUTSU\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HOSEISOFTHOSHU\_KIN | {"Generator\_UserColumnName": "HOSEISOFTHOSHU\_KIN", "Generator\_ColumnPropNameInRow": "HOSEISOFTHOSHU\_KIN", "Generator\_ColumnVarNameInTable": "columnHOSEISOFTHOSHU\_KIN", "Generator\_ColumnPropNameInTable": "HOSEISOFTHOSHU\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIHON\_BRAND | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_BRAND | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_BRAND | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HARDHOSYU\_KIN | {"Generator\_UserColumnName": "HARDHOSYU\_KIN", "Generator\_ColumnPropNameInRow": "HARDHOSYU\_KIN", "Generator\_ColumnVarNameInTable": "columnHARDHOSYU\_KIN", "Generator\_ColumnPropNameInTable": "HARDHOSYU\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | CHOSEI\_KIN | {"Generator\_UserColumnName": "CHOSEI\_KIN", "Generator\_ColumnPropNameInRow": "CHOSEI\_KIN", "Generator\_ColumnVarNameInTable": "columnCHOSEI\_KIN", "Generator\_ColumnPropNameInTable": "CHOSEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HOSHU\_KIN | {"Generator\_UserColumnName": "HOSHU\_KIN", "Generator\_ColumnPropNameInRow": "HOSHU\_KIN", "Generator\_ColumnVarNameInTable": "columnHOSHU\_KIN", "Generator\_ColumnPropNameInTable": "HOSHU\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_BRAND | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIHON\_BRAND | UM\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_KIKIMEISAI | UM\_KIKIMEISAI\_ID | {"Generator\_UserColumnName": "UM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_KIKIMEISAI | UM\_KIKIKOSEI\_ID | {"Generator\_UserColumnName": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIKIKOSEI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIKIKOSEI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIKIKOSEI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | SEIZOMAKER\_ID | {"Generator\_UserColumnName": "SEIZOMAKER\_ID", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_ID", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_ID", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIKIMEISAI | KIKIMEISAI\_ID | {"Generator\_UserColumnName": "KIKIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "KIKIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnKIKIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "KIKIMEISAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIKIMEISAI | KIKIKATASHIKI | {"Generator\_UserColumnName": "KIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnKIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIKIMEISAI | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | ATSUKAIKIKI\_ID | {"Generator\_UserColumnName": "ATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInRow": "ATSUKAIKIKI\_ID", "Generator\_ColumnVarNameInTable": "columnATSUKAIKIKI\_ID", "Generator\_ColumnPropNameInTable": "ATSUKAIKIKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIKIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIKIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIKIMEISAI | UM\_KIHON\_BRAND\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_BRAND\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_BRAND\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | KIKIKOSEI\_NK | {"Generator\_UserColumnName": "KIKIKOSEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIKOSEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIKOSEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIKIMEISAI | TEHAISEIBAN | {"Generator\_UserColumnName": "TEHAISEIBAN", "Generator\_ColumnPropNameInRow": "TEHAISEIBAN", "Generator\_ColumnVarNameInTable": "columnTEHAISEIBAN", "Generator\_ColumnPropNameInTable": "TEHAISEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIKIMEISAI | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIKIMEISAI | BRANDSYOSAI\_NK | {"Generator\_UserColumnName": "BRANDSYOSAI\_NK", "Generator\_ColumnPropNameInRow": "BRANDSYOSAI\_NK", "Generator\_ColumnVarNameInTable": "columnBRANDSYOSAI\_NK", "Generator\_ColumnPropNameInTable": "BRANDSYOSAI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIKIMEISAI | HYOJUN\_KIN | {"Generator\_UserColumnName": "HYOJUN\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUN\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUN\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUN\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | SIKIRI\_KIN | {"Generator\_UserColumnName": "SIKIRI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | KEIYAKUNAIYO | {"Generator\_UserColumnName": "KEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "KEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "KEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_UM\_KIKIMEISAI | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIKIMEISAI | TORIHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TORIHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TORIHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | DAIFUKUHOSYUJIKAN\_ID | {"Generator\_UserColumnName": "DAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "DAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnDAIFUKUHOSYUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "DAIFUKUHOSYUJIKAN\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | TENKENKAISU | {"Generator\_UserColumnName": "TENKENKAISU", "Generator\_ColumnPropNameInRow": "TENKENKAISU", "Generator\_ColumnVarNameInTable": "columnTENKENKAISU", "Generator\_ColumnPropNameInTable": "TENKENKAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIKIMEISAI | TENKENYOBI | {"Generator\_UserColumnName": "TENKENYOBI", "Generator\_ColumnPropNameInRow": "TENKENYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENYOBI", "Generator\_ColumnPropNameInTable": "TENKENYOBIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIKIMEISAI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIKIMEISAI | SERVICEKEITAI | {"Generator\_UserColumnName": "SERVICEKEITAI", "Generator\_ColumnPropNameInRow": "SERVICEKEITAI", "Generator\_ColumnVarNameInTable": "columnSERVICEKEITAI", "Generator\_ColumnPropNameInTable": "SERVICEKEITAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIKIMEISAI | UM\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | UM\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_MITSUMORI | UM\_KIHON\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_MITSUMORI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UM\_MITSUMORI | SYOUNIN\_JOTAI | {"Generator\_UserColumnName": "SYOUNIN\_JOTAI", "Generator\_ColumnPropNameInRow": "SYOUNIN\_JOTAI", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_JOTAI", "Generator\_ColumnPropNameInTable": "SYOUNIN\_JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_MITSUMORI | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | MITSUMORI\_JOUKEN | {"Generator\_UserColumnName": "MITSUMORI\_JOUKEN", "Generator\_ColumnPropNameInRow": "MITSUMORI\_JOUKEN", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_JOUKEN", "Generator\_ColumnPropNameInTable": "MITSUMORI\_JOUKENColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_MITSUMORI | MITSUMORI\_GKIN | {"Generator\_UserColumnName": "MITSUMORI\_GKIN", "Generator\_ColumnPropNameInRow": "MITSUMORI\_GKIN", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_GKIN", "Generator\_ColumnPropNameInTable": "MITSUMORI\_GKINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | SYUSSEINEBIKI\_KIN | {"Generator\_UserColumnName": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInRow": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInTable": "SYUSSEINEBIKI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_MITSUMORI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_MITSUMORI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_MITSUMORI | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_MITSUMORI | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UM_MITSUMORI" />
      <xs:field xpath="mstns:UM_MITSUMORI_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UM_MITSUMORI_MCM_UM_KIKIMEISAI" ns0:parent="MCM_UM_MITSUMORI" ns0:child="MCM_UM_KIKIMEISAI" ns0:parentkey="UM_MITSUMORI_ID" ns0:childkey="UM_MITSUMORI_ID" ns1:Generator_UserRelationName="MCM_UM_MITSUMORI_MCM_UM_KIKIMEISAI" ns1:Generator_RelationVarName="relationMCM_UM_MITSUMORI_MCM_UM_KIKIMEISAI" ns1:Generator_UserChildTable="MCM_UM_KIKIMEISAI" ns1:Generator_UserParentTable="MCM_UM_MITSUMORI" ns1:Generator_ParentPropName="MCM_UM_MITSUMORIRow" ns1:Generator_ChildPropName="GetMCM_UM_KIKIMEISAIRows" />
      
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_UM_MITSUMORI_MCM_UM_KIHON_BRAND" ns0:parent="MCM_UM_MITSUMORI" ns0:child="MCM_UM_KIHON_BRAND" ns0:parentkey="UM_MITSUMORI_ID" ns0:childkey="UM_MITSUMORI_ID" ns1:Generator_UserRelationName="MCM_UM_MITSUMORI_MCM_UM_KIHON_BRAND" ns1:Generator_RelationVarName="relationMCM_UM_MITSUMORI_MCM_UM_KIHON_BRAND" ns1:Generator_UserChildTable="MCM_UM_KIHON_BRAND" ns1:Generator_UserParentTable="MCM_UM_MITSUMORI" ns1:Generator_ParentPropName="MCM_UM_MITSUMORIRow" ns1:Generator_ChildPropName="GetMCM_UM_KIHON_BRANDRows" />
    
```

</details>

### 3002　Mcm3002uDataSet.xsd

[Mcm3002uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3002U/Mcm3002uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_TK\_SIHARAI\_V | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAI\_V | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAI\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD"} | [["maxLength", {"value": "12"}]] |
| MCM\_TK\_SIHARAI\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn"} | [["maxLength", {"value": "7"}]] |
| MCM\_TK\_SIHARAI\_V | JOTAI | {"Generator\_UserColumnName": "JOTAI", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "Generator\_ColumnVarNameInTable": "columnJOTAI", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_TK\_SIHARAI\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_SIHARAI\_V | TENPO\_NK | {"Generator\_UserColumnName": "TENPO\_NK", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnTENPO\_NK", "Generator\_ColumnPropNameInTable": "TENPO\_NKColumn", "Generator\_ColumnPropNameInRow": "TENPO\_NK", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_TK\_SIHARAI\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAI\_V | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_TK\_SIHARAI\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_SIHARAI\_V | SIHARAI | {"Generator\_UserColumnName": "SIHARAI", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "SIHARAI", "Generator\_ColumnPropNameInTable": "SIHARAIColumn", "Generator\_ColumnVarNameInTable": "columnSIHARAI", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_TK\_SIHARAI\_V | KAISU | {"Generator\_UserColumnName": "KAISU", "Generator\_ColumnVarNameInTable": "columnKAISU", "Generator\_ColumnPropNameInRow": "KAISU", "Generator\_ColumnPropNameInTable": "KAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAI\_V | BIKO | {"Generator\_UserColumnName": "BIKO", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "Generator\_ColumnVarNameInTable": "columnBIKO", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_SIHARAI\_V | TSUKI | {"Generator\_UserColumnName": "TSUKI", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "Generator\_ColumnVarNameInTable": "columnTSUKI", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_TK\_SIHARAI\_V | HARD\_SEIBAN | {"Generator\_UserColumnName": "HARD\_SEIBAN", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnHARD\_SEIBAN", "Generator\_ColumnPropNameInTable": "HARD\_SEIBANColumn", "Generator\_ColumnPropNameInRow": "HARD\_SEIBAN", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_SIHARAI\_V | TORISYUTANTOSYA\_NK | {"Generator\_UserColumnName": "TORISYUTANTOSYA\_NK", "nullValue": "\_null", "Generator\_ColumnPropNameInRow": "TORISYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "TORISYUTANTOSYA\_NKColumn", "Generator\_ColumnVarNameInTable": "columnTORISYUTANTOSYA\_NK", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_SIHARAI\_V | MCM\_FN\_SIHARAI | {"Generator\_UserColumnName": "MCM\_FN\_SIHARAI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnMCM\_FN\_SIHARAI", "Generator\_ColumnPropNameInTable": "MCM\_FN\_SIHARAIColumn", "Generator\_ColumnPropNameInRow": "MCM\_FN\_SIHARAI", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |

<details>
<summary>キー・関連の定義原文</summary>

</details>

### 3003　Mcm3003uDataSet.xsd

[Mcm3003uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3003U/Mcm3003uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_TK\_TENKEN\_V | TSUKI | {"Generator\_UserColumnName": "TSUKI", "Generator\_ColumnVarNameInTable": "columnTSUKI", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | TENKENKAISU | {"Generator\_UserColumnName": "TENKENKAISU", "Generator\_ColumnVarNameInTable": "columnTENKENKAISU", "Generator\_ColumnPropNameInRow": "TENKENKAISU", "Generator\_ColumnPropNameInTable": "TENKENKAISUColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | NONYUSAKIKANA\_KN | {"Generator\_UserColumnName": "NONYUSAKIKANA\_KN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInRow": "NONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInTable": "NONYUSAKIKANA\_KNColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_TENKEN\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_TK\_TENKEN\_V | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_TENKEN\_V | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_TENKEN\_V | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKEN\_V | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | KIKIHINMEI\_NK | {"Generator\_UserColumnName": "KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnKIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "KIKIHINMEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | TENKENKANOYOBI | {"Generator\_UserColumnName": "TENKENKANOYOBI", "Generator\_ColumnPropNameInRow": "TENKENKANOYOBI", "Generator\_ColumnVarNameInTable": "columnTENKENKANOYOBI", "Generator\_ColumnPropNameInTable": "TENKENKANOYOBIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_TENKEN\_V | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | HOSHUGAISHA\_NK | {"Generator\_UserColumnName": "HOSHUGAISHA\_NK", "Generator\_ColumnPropNameInRow": "HOSHUGAISHA\_NK", "Generator\_ColumnVarNameInTable": "columnHOSHUGAISHA\_NK", "Generator\_ColumnPropNameInTable": "HOSHUGAISHA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_TK\_TENKEN\_V | HOSHUGAISHAJIGYOSYO\_NK | {"Generator\_UserColumnName": "HOSHUGAISHAJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "HOSHUGAISHAJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnHOSHUGAISHAJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "HOSHUGAISHAJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_TENKEN\_V | HOSHUGAISHATANTOSYA\_NK | {"Generator\_UserColumnName": "HOSHUGAISHATANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "HOSHUGAISHATANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnHOSHUGAISHATANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "HOSHUGAISHATANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_TENKEN\_V | HOSHUGAISHATEL\_NO | {"Generator\_UserColumnName": "HOSHUGAISHATEL\_NO", "Generator\_ColumnPropNameInRow": "HOSHUGAISHATEL\_NO", "Generator\_ColumnVarNameInTable": "columnHOSHUGAISHATEL\_NO", "Generator\_ColumnPropNameInTable": "HOSHUGAISHATEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKEN\_V | HOSHUGAISHAFAX\_NO | {"Generator\_UserColumnName": "HOSHUGAISHAFAX\_NO", "Generator\_ColumnPropNameInRow": "HOSHUGAISHAFAX\_NO", "Generator\_ColumnVarNameInTable": "columnHOSHUGAISHAFAX\_NO", "Generator\_ColumnPropNameInTable": "HOSHUGAISHAFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKEN\_V | ON\_DT | {"Generator\_UserColumnName": "ON\_DT", "Generator\_ColumnPropNameInRow": "ON\_DT", "Generator\_ColumnVarNameInTable": "columnON\_DT", "Generator\_ColumnPropNameInTable": "ON\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | MEISHO4\_NK | {"Generator\_UserColumnName": "MEISHO4\_NK", "Generator\_ColumnPropNameInRow": "MEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "MEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_TK\_TENKEN\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_TENKEN\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_TENKEN\_V | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |

<details>
<summary>キー・関連の定義原文</summary>

</details>

### 3004　Mcm3004uDataSet.xsd

[Mcm3004uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_ALERT1\_V | MAR\_HYOJIJUN | {"Generator\_UserColumnName": "MAR\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAR\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAR\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAR\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT1\_V | MAR\_TENPORYAKU\_NK | {"Generator\_UserColumnName": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT1\_V | MAR\_TANTO\_NK | {"Generator\_UserColumnName": "MAR\_TANTO\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TANTO\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TANTO\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT1\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT1\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT1\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT1\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT1\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT1\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT1\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT1\_V | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT1\_V | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT1\_V | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT1\_V | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT1\_V | SEARCH\_DT | {"Generator\_UserColumnName": "SEARCH\_DT", "Generator\_ColumnPropNameInRow": "SEARCH\_DT", "Generator\_ColumnVarNameInTable": "columnSEARCH\_DT", "Generator\_ColumnPropNameInTable": "SEARCH\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_SAGYOYOTEI | TOROKU\_ID | {"Generator\_UserColumnName": "TOROKU\_ID", "Generator\_ColumnVarNameInTable": "columnTOROKU\_ID", "Generator\_ColumnPropNameInRow": "TOROKU\_ID", "Generator\_ColumnPropNameInTable": "TOROKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MO\_SAGYOYOTEI | YOTEI\_DT | {"Generator\_UserColumnName": "YOTEI\_DT", "Generator\_ColumnVarNameInTable": "columnYOTEI\_DT", "Generator\_ColumnPropNameInRow": "YOTEI\_DT", "Generator\_ColumnPropNameInTable": "YOTEI\_DTColumn", "type": "xs:dateTime"} | [] |
| MCM\_MO\_SAGYOYOTEI | KANRYO\_FLG | {"Generator\_UserColumnName": "KANRYO\_FLG", "Generator\_ColumnVarNameInTable": "columnKANRYO\_FLG", "Generator\_ColumnPropNameInRow": "KANRYO\_FLG", "Generator\_ColumnPropNameInTable": "KANRYO\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MO\_SAGYOYOTEI | NAIYO | {"Generator\_UserColumnName": "NAIYO", "Generator\_ColumnVarNameInTable": "columnNAIYO", "Generator\_ColumnPropNameInRow": "NAIYO", "Generator\_ColumnPropNameInTable": "NAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MO\_SAGYOYOTEI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MO\_SAGYOYOTEI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_SAGYOYOTEI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MO\_SAGYOYOTEI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_SAGYOYOTEI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT2\_V | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_ALERT2\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT2\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT2\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT2\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT2\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT2\_V | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_ALERT2\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT2\_V | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_ALERT2\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT2\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT2\_V | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT2\_V | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT2\_V | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT2\_V | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT3\_V | MAR\_HYOJIJUN | {"Generator\_UserColumnName": "MAR\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAR\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAR\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAR\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | MAR\_TENPORYAKU\_NK | {"Generator\_UserColumnName": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT3\_V | TKA\_JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "TKA\_JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "TKA\_JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnTKA\_JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "TKA\_JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | TKA\_KEIYAKU\_NO | {"Generator\_UserColumnName": "TKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "TKA\_KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnTKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "TKA\_KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT3\_V | TKB\_TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | TKB\_TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_ALERT3\_V | TKB\_TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT3\_V | TKB\_NONYUSAKI\_NK | {"Generator\_UserColumnName": "TKB\_NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "TKB\_NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "TKB\_NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT3\_V | TKB\_NONYUSAKI\_CD | {"Generator\_UserColumnName": "TKB\_NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "TKB\_NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTKB\_NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "TKB\_NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT3\_V | TKB\_SUPPORT\_ID | {"Generator\_UserColumnName": "TKB\_SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "TKB\_SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnTKB\_SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "TKB\_SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT3\_V | TKB\_NONYUBUSYO\_NK | {"Generator\_UserColumnName": "TKB\_NONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "TKB\_NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_NONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "TKB\_NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_ALERT3\_V | TKE\_BRAND\_NK | {"Generator\_UserColumnName": "TKE\_BRAND\_NK", "Generator\_ColumnPropNameInRow": "TKE\_BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnTKE\_BRAND\_NK", "Generator\_ColumnPropNameInTable": "TKE\_BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT3\_V | TKD\_SEIZOMAKER\_NK | {"Generator\_UserColumnName": "TKD\_SEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "TKD\_SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnTKD\_SEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "TKD\_SEIZOMAKER\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT3\_V | TKD\_KIKIHINMEI\_NK | {"Generator\_UserColumnName": "TKD\_KIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "TKD\_KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnTKD\_KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "TKD\_KIKIHINMEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT3\_V | TKD\_KIKIKATASHIKI | {"Generator\_UserColumnName": "TKD\_KIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "TKD\_KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnTKD\_KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "TKD\_KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT3\_V | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | TANKA\_KIN | {"Generator\_UserColumnName": "TANKA\_KIN", "Generator\_ColumnPropNameInRow": "TANKA\_KIN", "Generator\_ColumnVarNameInTable": "columnTANKA\_KIN", "Generator\_ColumnPropNameInTable": "TANKA\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | TOTAL\_KIN | {"Generator\_UserColumnName": "TOTAL\_KIN", "Generator\_ColumnPropNameInRow": "TOTAL\_KIN", "Generator\_ColumnVarNameInTable": "columnTOTAL\_KIN", "Generator\_ColumnPropNameInTable": "TOTAL\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT3\_V | KEIYAKU\_KIKAN | {"Generator\_UserColumnName": "KEIYAKU\_KIKAN", "Generator\_ColumnPropNameInRow": "KEIYAKU\_KIKAN", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_KIKAN", "Generator\_ColumnPropNameInTable": "KEIYAKU\_KIKANColumn", "minOccurs": "0"} | [["maxLength", {"value": "24"}]] |
| MCM\_ALERT3\_V | MITSUMORI\_WORK | {"Generator\_UserColumnName": "MITSUMORI\_WORK", "Generator\_ColumnPropNameInRow": "MITSUMORI\_WORK", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_WORK", "Generator\_ColumnPropNameInTable": "MITSUMORI\_WORKColumn", "minOccurs": "0"} | [["maxLength", {"value": "10"}]] |
| MCM\_ALERT3\_V | MITSUMORI\_END\_DT | {"Generator\_UserColumnName": "MITSUMORI\_END\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_END\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_END\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_END\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | MAR\_HYOJIJUN | {"Generator\_UserColumnName": "MAR\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAR\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAR\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAR\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | MAR\_TENPORYAKU\_NK | {"Generator\_UserColumnName": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT4\_V | TKA\_JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "TKA\_JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "TKA\_JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnTKA\_JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "TKA\_JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | TKA\_KEIYAKU\_NO | {"Generator\_UserColumnName": "TKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "TKA\_KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnTKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "TKA\_KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT4\_V | TKB\_TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | TKB\_TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_ALERT4\_V | TKB\_TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT4\_V | TKB\_NONYUSAKI\_NK | {"Generator\_UserColumnName": "TKB\_NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "TKB\_NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "TKB\_NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT4\_V | TKB\_NONYUSAKI\_CD | {"Generator\_UserColumnName": "TKB\_NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "TKB\_NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTKB\_NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "TKB\_NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT4\_V | TKB\_SUPPORT\_ID | {"Generator\_UserColumnName": "TKB\_SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "TKB\_SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnTKB\_SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "TKB\_SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT4\_V | TKB\_NONYUBUSYO\_NK | {"Generator\_UserColumnName": "TKB\_NONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "TKB\_NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_NONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "TKB\_NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_ALERT4\_V | TKE\_BRAND\_NK | {"Generator\_UserColumnName": "TKE\_BRAND\_NK", "Generator\_ColumnPropNameInRow": "TKE\_BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnTKE\_BRAND\_NK", "Generator\_ColumnPropNameInTable": "TKE\_BRAND\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT4\_V | TKD\_SEIZOMAKER\_NK | {"Generator\_UserColumnName": "TKD\_SEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "TKD\_SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnTKD\_SEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "TKD\_SEIZOMAKER\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT4\_V | TKD\_KIKIHINMEI\_NK | {"Generator\_UserColumnName": "TKD\_KIKIHINMEI\_NK", "Generator\_ColumnPropNameInRow": "TKD\_KIKIHINMEI\_NK", "Generator\_ColumnVarNameInTable": "columnTKD\_KIKIHINMEI\_NK", "Generator\_ColumnPropNameInTable": "TKD\_KIKIHINMEI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT4\_V | TKD\_KIKIKATASHIKI | {"Generator\_UserColumnName": "TKD\_KIKIKATASHIKI", "Generator\_ColumnPropNameInRow": "TKD\_KIKIKATASHIKI", "Generator\_ColumnVarNameInTable": "columnTKD\_KIKIKATASHIKI", "Generator\_ColumnPropNameInTable": "TKD\_KIKIKATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT4\_V | SURYO\_NM | {"Generator\_UserColumnName": "SURYO\_NM", "Generator\_ColumnPropNameInRow": "SURYO\_NM", "Generator\_ColumnVarNameInTable": "columnSURYO\_NM", "Generator\_ColumnPropNameInTable": "SURYO\_NMColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | TANKA\_KIN | {"Generator\_UserColumnName": "TANKA\_KIN", "Generator\_ColumnPropNameInRow": "TANKA\_KIN", "Generator\_ColumnVarNameInTable": "columnTANKA\_KIN", "Generator\_ColumnPropNameInTable": "TANKA\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | TOTAL\_KIN | {"Generator\_UserColumnName": "TOTAL\_KIN", "Generator\_ColumnPropNameInRow": "TOTAL\_KIN", "Generator\_ColumnVarNameInTable": "columnTOTAL\_KIN", "Generator\_ColumnPropNameInTable": "TOTAL\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT4\_V | KEIYAKU\_KIKAN | {"Generator\_UserColumnName": "KEIYAKU\_KIKAN", "Generator\_ColumnPropNameInRow": "KEIYAKU\_KIKAN", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_KIKAN", "Generator\_ColumnPropNameInTable": "KEIYAKU\_KIKANColumn", "minOccurs": "0"} | [["maxLength", {"value": "24"}]] |
| MCM\_ALERT5\_V | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT5\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT5\_V | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT5\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT5\_V | TKA\_TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TKA\_TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TKA\_TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTKA\_TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TKA\_TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | TKA\_KEIYAKU\_NO | {"Generator\_UserColumnName": "TKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "TKA\_KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnTKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "TKA\_KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT5\_V | TKA\_KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "TKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "TKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnTKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "TKA\_KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | TKA\_ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "TKA\_ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "TKA\_ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnTKA\_ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "TKA\_ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | TKA\_SEARCH\_DT | {"Generator\_UserColumnName": "TKA\_SEARCH\_DT", "Generator\_ColumnPropNameInRow": "TKA\_SEARCH\_DT", "Generator\_ColumnVarNameInTable": "columnTKA\_SEARCH\_DT", "Generator\_ColumnPropNameInTable": "TKA\_SEARCH\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | TKB\_TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_ALERT5\_V | TKB\_TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TKB\_TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TKB\_TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTKB\_TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TKB\_TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT5\_V | UKA\_UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UKA\_UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UKA\_UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUKA\_UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UKA\_UK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | UKA\_KEIYAKU\_NO | {"Generator\_UserColumnName": "UKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "UKA\_KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnUKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "UKA\_KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT5\_V | UKA\_KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "UKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "UKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnUKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "UKA\_KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | UKA\_ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "UKA\_ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "UKA\_ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnUKA\_ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "UKA\_ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | UKA\_SEARCH\_DT | {"Generator\_UserColumnName": "UKA\_SEARCH\_DT", "Generator\_ColumnPropNameInRow": "UKA\_SEARCH\_DT", "Generator\_ColumnVarNameInTable": "columnUKA\_SEARCH\_DT", "Generator\_ColumnPropNameInTable": "UKA\_SEARCH\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | MAR\_TENPO\_ID | {"Generator\_UserColumnName": "MAR\_TENPO\_ID", "Generator\_ColumnPropNameInRow": "MAR\_TENPO\_ID", "Generator\_ColumnVarNameInTable": "columnMAR\_TENPO\_ID", "Generator\_ColumnPropNameInTable": "MAR\_TENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT5\_V | MAR\_TENPORYAKU\_NK | {"Generator\_UserColumnName": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT5\_V | MAR\_TANTO\_NK | {"Generator\_UserColumnName": "MAR\_TANTO\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TANTO\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TANTO\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT5\_V | MAR\_HYOJIJUN | {"Generator\_UserColumnName": "MAR\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAR\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAR\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAR\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT6\_V | MAR\_HYOJIJUN | {"Generator\_UserColumnName": "MAR\_HYOJIJUN", "Generator\_ColumnPropNameInRow": "MAR\_HYOJIJUN", "Generator\_ColumnVarNameInTable": "columnMAR\_HYOJIJUN", "Generator\_ColumnPropNameInTable": "MAR\_HYOJIJUNColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT6\_V | MAR\_TENPORYAKU\_NK | {"Generator\_UserColumnName": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT6\_V | MAR\_TANTO\_NK | {"Generator\_UserColumnName": "MAR\_TANTO\_NK", "Generator\_ColumnPropNameInRow": "MAR\_TANTO\_NK", "Generator\_ColumnVarNameInTable": "columnMAR\_TANTO\_NK", "Generator\_ColumnPropNameInTable": "MAR\_TANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT6\_V | UKB\_NONYUSAKI\_CD | {"Generator\_UserColumnName": "UKB\_NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "UKB\_NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnUKB\_NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "UKB\_NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT6\_V | UKB\_NONYUSAKI\_NK | {"Generator\_UserColumnName": "UKB\_NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "UKB\_NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnUKB\_NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "UKB\_NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT6\_V | UKB\_SUPPORT\_ID | {"Generator\_UserColumnName": "UKB\_SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "UKB\_SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnUKB\_SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "UKB\_SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT6\_V | UKB\_PLANT\_NK | {"Generator\_UserColumnName": "UKB\_PLANT\_NK", "Generator\_ColumnPropNameInRow": "UKB\_PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnUKB\_PLANT\_NK", "Generator\_ColumnPropNameInTable": "UKB\_PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT6\_V | UKA\_KEIYAKU\_NO | {"Generator\_UserColumnName": "UKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "UKA\_KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnUKA\_KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "UKA\_KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT6\_V | UKA\_KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "UKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "UKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnUKA\_KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "UKA\_KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT6\_V | START\_DT | {"Generator\_UserColumnName": "START\_DT", "Generator\_ColumnPropNameInRow": "START\_DT", "Generator\_ColumnVarNameInTable": "columnSTART\_DT", "Generator\_ColumnPropNameInTable": "START\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT6\_V | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT6\_V | HARD\_SEIBAN | {"Generator\_UserColumnName": "HARD\_SEIBAN", "Generator\_ColumnPropNameInRow": "HARD\_SEIBAN", "Generator\_ColumnVarNameInTable": "columnHARD\_SEIBAN", "Generator\_ColumnPropNameInTable": "HARD\_SEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_ALERT6\_V | UKB\_HOSYU\_GKIN | {"Generator\_UserColumnName": "UKB\_HOSYU\_GKIN", "Generator\_ColumnVarNameInTable": "columnUKB\_HOSYU\_GKIN", "Generator\_ColumnPropNameInRow": "UKB\_HOSYU\_GKIN", "Generator\_ColumnPropNameInTable": "UKB\_HOSYU\_GKINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT8\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT8\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT8\_V | TENPORYAKU\_NK | {"Generator\_UserColumnName": "TENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "TENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnTENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "TENPORYAKU\_NKColumn"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT8\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT8\_V | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT8\_V | PC\_TENKEN | {"Generator\_UserColumnName": "PC\_TENKEN", "Generator\_ColumnPropNameInRow": "PC\_TENKEN", "Generator\_ColumnVarNameInTable": "columnPC\_TENKEN", "Generator\_ColumnPropNameInTable": "PC\_TENKENColumn", "type": "xs:string", "minOccurs": "0"} | [] |
| MCM\_ALERT8\_V | DENTI\_SYURUI | {"Generator\_UserColumnName": "DENTI\_SYURUI", "Generator\_ColumnPropNameInRow": "DENTI\_SYURUI", "Generator\_ColumnVarNameInTable": "columnDENTI\_SYURUI", "Generator\_ColumnPropNameInTable": "DENTI\_SYURUIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_ALERT8\_V | SYURUI | {"Generator\_UserColumnName": "SYURUI", "Generator\_ColumnPropNameInRow": "SYURUI", "Generator\_ColumnVarNameInTable": "columnSYURUI", "Generator\_ColumnPropNameInTable": "SYURUIColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_ALERT8\_V | GOUKI | {"Generator\_UserColumnName": "GOUKI", "Generator\_ColumnPropNameInRow": "GOUKI", "Generator\_ColumnVarNameInTable": "columnGOUKI", "Generator\_ColumnPropNameInTable": "GOUKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_ALERT8\_V | SEIZOMAKER\_NK | {"Generator\_UserColumnName": "SEIZOMAKER\_NK", "Generator\_ColumnPropNameInRow": "SEIZOMAKER\_NK", "Generator\_ColumnVarNameInTable": "columnSEIZOMAKER\_NK", "Generator\_ColumnPropNameInTable": "SEIZOMAKER\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT8\_V | KATASHIKI | {"Generator\_UserColumnName": "KATASHIKI", "Generator\_ColumnPropNameInRow": "KATASHIKI", "Generator\_ColumnVarNameInTable": "columnKATASHIKI", "Generator\_ColumnPropNameInTable": "KATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_ALERT8\_V | BATTERY\_KATASHIKI | {"Generator\_UserColumnName": "BATTERY\_KATASHIKI", "Generator\_ColumnPropNameInRow": "BATTERY\_KATASHIKI", "Generator\_ColumnVarNameInTable": "columnBATTERY\_KATASHIKI", "Generator\_ColumnPropNameInTable": "BATTERY\_KATASHIKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_ALERT8\_V | ITIJINONYU\_DT | {"Generator\_UserColumnName": "ITIJINONYU\_DT", "Generator\_ColumnPropNameInRow": "ITIJINONYU\_DT", "Generator\_ColumnVarNameInTable": "columnITIJINONYU\_DT", "Generator\_ColumnPropNameInTable": "ITIJINONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT8\_V | NEXT\_DT | {"Generator\_UserColumnName": "NEXT\_DT", "Generator\_ColumnPropNameInRow": "NEXT\_DT", "Generator\_ColumnVarNameInTable": "columnNEXT\_DT", "Generator\_ColumnPropNameInTable": "NEXT\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT8\_V | UK\_KEIYAKU | {"Generator\_UserColumnName": "UK\_KEIYAKU", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT8\_V | TK\_KEIYAKU | {"Generator\_UserColumnName": "TK\_KEIYAKU", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT8\_V | CT\_KEIYAKU | {"Generator\_UserColumnName": "CT\_KEIYAKU", "Generator\_ColumnPropNameInRow": "CT\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnCT\_KEIYAKU", "Generator\_ColumnPropNameInTable": "CT\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT9\_V | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_ALERT9\_V | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT9\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_ALERT9\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT9\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_ALERT9\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_ALERT9\_V | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT9\_V | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_ALERT9\_V | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_ALERT9\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_ALERT9\_V | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_ALERT9\_V | TK\_KIKAN\_ID | {"Generator\_UserColumnName": "TK\_KIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KIKAN\_ID", "Generator\_ColumnPropNameInRow": "TK\_KIKAN\_ID", "Generator\_ColumnPropNameInTable": "TK\_KIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_ALERT9\_V | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_ALERT9\_V | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_ALERT9\_V | HYOJUNGOKEI\_KIN | {"Generator\_UserColumnName": "HYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "HYOJUNGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnHYOJUNGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "HYOJUNGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | SIKIRISYOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRISYOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRISYOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRISYOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | SYUSSEINEBIKI\_KIN | {"Generator\_UserColumnName": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInRow": "SYUSSEINEBIKI\_KIN", "Generator\_ColumnVarNameInTable": "columnSYUSSEINEBIKI\_KIN", "Generator\_ColumnPropNameInTable": "SYUSSEINEBIKI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_ALERT9\_V | SIKIRIGOKEI\_KIN | {"Generator\_UserColumnName": "SIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInRow": "SIKIRIGOKEI\_KIN", "Generator\_ColumnVarNameInTable": "columnSIKIRIGOKEI\_KIN", "Generator\_ColumnPropNameInTable": "SIKIRIGOKEI\_KINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MO_SAGYOYOTEI" />
      <xs:field xpath="mstns:TOROKU_ID" />
    </xs:unique>
  
```

</details>

### 3005　Mcm3005uDataSet.xsd

[Mcm3005uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_MO\_SAGYOYOTEI | TOROKU\_ID | {"Generator\_UserColumnName": "TOROKU\_ID", "Generator\_ColumnVarNameInTable": "columnTOROKU\_ID", "Generator\_ColumnPropNameInRow": "TOROKU\_ID", "Generator\_ColumnPropNameInTable": "TOROKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MO\_SAGYOYOTEI | YOTEI\_DT | {"Generator\_UserColumnName": "YOTEI\_DT", "Generator\_ColumnVarNameInTable": "columnYOTEI\_DT", "Generator\_ColumnPropNameInRow": "YOTEI\_DT", "Generator\_ColumnPropNameInTable": "YOTEI\_DTColumn", "type": "xs:dateTime"} | [] |
| MCM\_MO\_SAGYOYOTEI | KANRYO\_FLG | {"Generator\_UserColumnName": "KANRYO\_FLG", "Generator\_ColumnVarNameInTable": "columnKANRYO\_FLG", "Generator\_ColumnPropNameInRow": "KANRYO\_FLG", "Generator\_ColumnPropNameInTable": "KANRYO\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_MO\_SAGYOYOTEI | NAIYO | {"Generator\_UserColumnName": "NAIYO", "Generator\_ColumnVarNameInTable": "columnNAIYO", "Generator\_ColumnPropNameInRow": "NAIYO", "Generator\_ColumnPropNameInTable": "NAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MO\_SAGYOYOTEI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MO\_SAGYOYOTEI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_SAGYOYOTEI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MO\_SAGYOYOTEI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_SAGYOYOTEI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MO\_TENPU | TENPU\_ID | {"Generator\_UserColumnName": "TENPU\_ID", "Generator\_ColumnVarNameInTable": "columnTENPU\_ID", "Generator\_ColumnPropNameInRow": "TENPU\_ID", "Generator\_ColumnPropNameInTable": "TENPU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MO\_TENPU | TOROKU\_ID | {"Generator\_UserColumnName": "TOROKU\_ID", "Generator\_ColumnVarNameInTable": "columnTOROKU\_ID", "Generator\_ColumnPropNameInRow": "TOROKU\_ID", "Generator\_ColumnPropNameInTable": "TOROKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MO\_TENPU | TENPUFILE\_NK | {"Generator\_UserColumnName": "TENPUFILE\_NK", "Generator\_ColumnVarNameInTable": "columnTENPUFILE\_NK", "Generator\_ColumnPropNameInRow": "TENPUFILE\_NK", "Generator\_ColumnPropNameInTable": "TENPUFILE\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_MO\_TENPU | DIRECTORY | {"Generator\_UserColumnName": "DIRECTORY", "Generator\_ColumnVarNameInTable": "columnDIRECTORY", "Generator\_ColumnPropNameInRow": "DIRECTORY", "Generator\_ColumnPropNameInTable": "DIRECTORYColumn", "minOccurs": "0"} | [["maxLength", {"value": "2000"}]] |
| MCM\_MO\_TENPU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MO\_TENPU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_TENPU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MO\_TENPU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MO\_TENPU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MO_SAGYOYOTEI" />
      <xs:field xpath="mstns:TOROKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MO_TENPU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MO_TENPU" />
      <xs:field xpath="mstns:TENPU_ID" />
    </xs:unique>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="MCM_MO_SAGYOYOTEI_MCM_MO_TENPU" ns0:parent="MCM_MO_SAGYOYOTEI" ns0:child="MCM_MO_TENPU" ns0:parentkey="TOROKU_ID" ns0:childkey="TOROKU_ID" ns1:Generator_UserRelationName="MCM_MO_SAGYOYOTEI_MCM_MO_TENPU" ns1:Generator_RelationVarName="relationMCM_MO_SAGYOYOTEI_MCM_MO_TENPU" ns1:Generator_UserChildTable="MCM_MO_TENPU" ns1:Generator_UserParentTable="MCM_MO_SAGYOYOTEI" ns1:Generator_ParentPropName="MCM_MO_SAGYOYOTEIRow" ns1:Generator_ChildPropName="GetMCM_MO_TENPURows" />
    
```

</details>

### 3007　Mcm3007uDataSet.xsd

[Mcm3007uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_2004\_V | UM\_KIHON\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_MITSUMORI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_2004\_V | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_2004\_V | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_2004\_V | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_2004\_V | UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_2004\_V | FUSEI | {"Generator\_UserColumnName": "FUSEI", "Generator\_ColumnPropNameInRow": "FUSEI", "Generator\_ColumnVarNameInTable": "columnFUSEI", "Generator\_ColumnPropNameInTable": "FUSEIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_2004\_V | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_2004\_V | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_2004\_V | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_2004\_V | KYUNONYUSAKI\_NK | {"Generator\_UserColumnName": "KYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "KYUNONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnKYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "KYUNONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_2004\_V | NONYUSAKIKANA\_KN | {"Generator\_UserColumnName": "NONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInRow": "NONYUSAKIKANA\_KN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInTable": "NONYUSAKIKANA\_KNColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_2004\_V | NONYUSAKIEIMEI\_EN | {"Generator\_UserColumnName": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInRow": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInTable": "NONYUSAKIEIMEI\_ENColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_2004\_V | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_2004\_V | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_2004\_V | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_2004\_V | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_2004\_V | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_2004\_V | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_2004\_V | MITSUMORISAKUSEISYA\_NK | {"Generator\_UserColumnName": "MITSUMORISAKUSEISYA\_NK", "Generator\_ColumnPropNameInRow": "MITSUMORISAKUSEISYA\_NK", "Generator\_ColumnVarNameInTable": "columnMITSUMORISAKUSEISYA\_NK", "Generator\_ColumnPropNameInTable": "MITSUMORISAKUSEISYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_2004\_V | SOFUTENPO\_ID | {"Generator\_UserColumnName": "SOFUTENPO\_ID", "Generator\_ColumnPropNameInRow": "SOFUTENPO\_ID", "Generator\_ColumnVarNameInTable": "columnSOFUTENPO\_ID", "Generator\_ColumnPropNameInTable": "SOFUTENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_2004\_V | SOFUMEISHO1\_NK | {"Generator\_UserColumnName": "SOFUMEISHO1\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO1\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_2004\_V | SOFUMEISHO2\_NK | {"Generator\_UserColumnName": "SOFUMEISHO2\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO2\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_2004\_V | SOFUMEISHO3\_NK | {"Generator\_UserColumnName": "SOFUMEISHO3\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO3\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO3\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_2004\_V | SOFUMEISHO4\_NK | {"Generator\_UserColumnName": "SOFUMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_2004\_V | SOFUTENPORYAKU\_NK | {"Generator\_UserColumnName": "SOFUTENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "SOFUTENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUTENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "SOFUTENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_2004\_V | SOFUTANTO\_NK | {"Generator\_UserColumnName": "SOFUTANTO\_NK", "Generator\_ColumnPropNameInRow": "SOFUTANTO\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUTANTO\_NK", "Generator\_ColumnPropNameInTable": "SOFUTANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_2004\_V | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_2004\_V | SYOUNIN\_JOTAI | {"Generator\_UserColumnName": "SYOUNIN\_JOTAI", "Generator\_ColumnPropNameInRow": "SYOUNIN\_JOTAI", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_JOTAI", "Generator\_ColumnPropNameInTable": "SYOUNIN\_JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "10"}]] |
| MCM\_2004\_V | MITSUMORI\_COPY | {"Generator\_UserColumnName": "MITSUMORI\_COPY", "Generator\_ColumnPropNameInRow": "MITSUMORI\_COPY", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_COPY", "Generator\_ColumnPropNameInTable": "MITSUMORI\_COPYColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_2004\_V | MITSUMORI\_DEL | {"Generator\_UserColumnName": "MITSUMORI\_DEL", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DEL", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_2004\_V | KEIYAKU\_KEIYAKU | {"Generator\_UserColumnName": "KEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInRow": "KEIYAKU\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInTable": "KEIYAKU\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_2004\_V | KEIYAKU\_DEL | {"Generator\_UserColumnName": "KEIYAKU\_DEL", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DEL", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DEL", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| MCM\_2004\_V | MITSUMORI\_UPD | {"Generator\_UserColumnName": "MITSUMORI\_UPD", "Generator\_ColumnPropNameInRow": "MITSUMORI\_UPD", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_UPD", "Generator\_ColumnPropNameInTable": "MITSUMORI\_UPDColumn", "minOccurs": "0"} | [["maxLength", {"value": "8"}]] |
| MCM\_2004\_V | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_2004\_V | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_2004\_V | MITSUMORILEVEL | {"Generator\_UserColumnName": "MITSUMORILEVEL", "Generator\_ColumnPropNameInRow": "MITSUMORILEVEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORILEVEL", "Generator\_ColumnPropNameInTable": "MITSUMORILEVELColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | AUTO\_FLG | {"Generator\_UserColumnName": "AUTO\_FLG", "Generator\_ColumnVarNameInTable": "columnAUTO\_FLG", "Generator\_ColumnPropNameInRow": "AUTO\_FLG", "Generator\_ColumnPropNameInTable": "AUTO\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_KEIYAKU | UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_KEIYAKU | SHOKAI\_KEIYAKU\_DT | {"Generator\_UserColumnName": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnSHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "SHOKAI\_KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | UM\_KIHON\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_MITSUMORI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | UM\_MITSUMORI\_NO | {"Generator\_UserColumnName": "UM\_MITSUMORI\_NO", "Generator\_ColumnVarNameInTable": "columnUM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInRow": "UM\_MITSUMORI\_NO", "Generator\_ColumnPropNameInTable": "UM\_MITSUMORI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SYOUNIN\_JOTAI | {"Generator\_UserColumnName": "SYOUNIN\_JOTAI", "Generator\_ColumnVarNameInTable": "columnSYOUNIN\_JOTAI", "Generator\_ColumnPropNameInRow": "SYOUNIN\_JOTAI", "Generator\_ColumnPropNameInTable": "SYOUNIN\_JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "nullValue": "\_throw", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORISAKUSEISYA\_NK | {"Generator\_UserColumnName": "MITSUMORISAKUSEISYA\_NK", "Generator\_ColumnVarNameInTable": "columnMITSUMORISAKUSEISYA\_NK", "Generator\_ColumnPropNameInRow": "MITSUMORISAKUSEISYA\_NK", "Generator\_ColumnPropNameInTable": "MITSUMORISAKUSEISYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORIKIGEN | {"Generator\_UserColumnName": "MITSUMORIKIGEN", "Generator\_ColumnVarNameInTable": "columnMITSUMORIKIGEN", "Generator\_ColumnPropNameInRow": "MITSUMORIKIGEN", "Generator\_ColumnPropNameInTable": "MITSUMORIKIGENColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAITENPO\_ID | {"Generator\_UserColumnName": "IRAITENPO\_ID", "Generator\_ColumnVarNameInTable": "columnIRAITENPO\_ID", "Generator\_ColumnPropNameInRow": "IRAITENPO\_ID", "Generator\_ColumnPropNameInTable": "IRAITENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO1\_NK | {"Generator\_UserColumnName": "IRAIMEISHO1\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO1\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO2\_NK | {"Generator\_UserColumnName": "IRAIMEISHO2\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO2\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO3\_NK | {"Generator\_UserColumnName": "IRAIMEISHO3\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO3\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO3\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAIMEISHO4\_NK | {"Generator\_UserColumnName": "IRAIMEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAITENPORYAKU\_NK | {"Generator\_UserColumnName": "IRAITENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "IRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "IRAITENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | IRAITANTO\_NK | {"Generator\_UserColumnName": "IRAITANTO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITANTO\_NK", "Generator\_ColumnPropNameInRow": "IRAITANTO\_NK", "Generator\_ColumnPropNameInTable": "IRAITANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUTENPO\_ID | {"Generator\_UserColumnName": "SOFUTENPO\_ID", "Generator\_ColumnVarNameInTable": "columnSOFUTENPO\_ID", "Generator\_ColumnPropNameInRow": "SOFUTENPO\_ID", "Generator\_ColumnPropNameInTable": "SOFUTENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO1\_NK | {"Generator\_UserColumnName": "SOFUMEISHO1\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO1\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO2\_NK | {"Generator\_UserColumnName": "SOFUMEISHO2\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO2\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO3\_NK | {"Generator\_UserColumnName": "SOFUMEISHO3\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO3\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO3\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUMEISHO4\_NK | {"Generator\_UserColumnName": "SOFUMEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "SOFUMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "SOFUMEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUTENPORYAKU\_NK | {"Generator\_UserColumnName": "SOFUTENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUTENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "SOFUTENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "SOFUTENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | SOFUTANTO\_NK | {"Generator\_UserColumnName": "SOFUTANTO\_NK", "Generator\_ColumnVarNameInTable": "columnSOFUTANTO\_NK", "Generator\_ColumnPropNameInRow": "SOFUTANTO\_NK", "Generator\_ColumnPropNameInTable": "SOFUTANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORILEVEL | {"Generator\_UserColumnName": "MITSUMORILEVEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORILEVEL", "Generator\_ColumnPropNameInRow": "MITSUMORILEVEL", "Generator\_ColumnPropNameInTable": "MITSUMORILEVELColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | MITSUMORI\_JOUKEN | {"Generator\_UserColumnName": "MITSUMORI\_JOUKEN", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_JOUKEN", "Generator\_ColumnPropNameInRow": "MITSUMORI\_JOUKEN", "Generator\_ColumnPropNameInTable": "MITSUMORI\_JOUKENColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UM\_KIHON\_MITSUMORI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UM\_KIHON\_MITSUMORI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| UVA | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| UVA | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| UVA | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| UVA | TM\_IRAI\_NO | {"Generator\_UserColumnName": "TM\_IRAI\_NO", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_NO", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_NO", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| UVA | MITSUMORI\_DT | {"Generator\_UserColumnName": "MITSUMORI\_DT", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DT", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DT", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| UVA | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| UVA | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| UVA | FUSEI | {"Generator\_UserColumnName": "FUSEI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnFUSEI", "Generator\_ColumnPropNameInTable": "FUSEIColumn", "Generator\_ColumnPropNameInRow": "FUSEI", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| UVA | TORIHIKISAKI\_ID | {"Generator\_UserColumnName": "TORIHIKISAKI\_ID", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_ID", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_ID", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| UVA | TORIHIKISAKI\_CD | {"Generator\_UserColumnName": "TORIHIKISAKI\_CD", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_CD", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_CD", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| UVA | TORIHIKISAKI\_NK | {"Generator\_UserColumnName": "TORIHIKISAKI\_NK", "Generator\_ColumnPropNameInRow": "TORIHIKISAKI\_NK", "Generator\_ColumnVarNameInTable": "columnTORIHIKISAKI\_NK", "Generator\_ColumnPropNameInTable": "TORIHIKISAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| UVA | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| UVA | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| UVA | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| UVA | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| UVA | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| UVA | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| UVA | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| UVA | TENKENUMU | {"Generator\_UserColumnName": "TENKENUMU", "Generator\_ColumnPropNameInRow": "TENKENUMU", "Generator\_ColumnVarNameInTable": "columnTENKENUMU", "Generator\_ColumnPropNameInTable": "TENKENUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| UVA | TENKENKANOYOBI | {"Generator\_UserColumnName": "TENKENKANOYOBI", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnTENKENKANOYOBI", "Generator\_ColumnPropNameInTable": "TENKENKANOYOBIColumn", "Generator\_ColumnPropNameInRow": "TENKENKANOYOBI", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| UVA | YAKANTAIOUMU | {"Generator\_UserColumnName": "YAKANTAIOUMU", "Generator\_ColumnPropNameInRow": "YAKANTAIOUMU", "Generator\_ColumnVarNameInTable": "columnYAKANTAIOUMU", "Generator\_ColumnPropNameInTable": "YAKANTAIOUMUColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| UVA | MITSUMORI\_IRAI | {"Generator\_UserColumnName": "MITSUMORI\_IRAI", "Generator\_ColumnPropNameInRow": "MITSUMORI\_IRAI", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_IRAI", "Generator\_ColumnPropNameInTable": "MITSUMORI\_IRAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| UVA | MITSUMORI\_COPY | {"Generator\_UserColumnName": "MITSUMORI\_COPY", "Generator\_ColumnPropNameInRow": "MITSUMORI\_COPY", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_COPY", "Generator\_ColumnPropNameInTable": "MITSUMORI\_COPYColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| UVA | MITSUMORI\_DEL | {"Generator\_UserColumnName": "MITSUMORI\_DEL", "Generator\_ColumnPropNameInRow": "MITSUMORI\_DEL", "Generator\_ColumnVarNameInTable": "columnMITSUMORI\_DEL", "Generator\_ColumnPropNameInTable": "MITSUMORI\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| UVA | KEIYAKU\_KEIYAKU | {"Generator\_UserColumnName": "KEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInRow": "KEIYAKU\_KEIYAKU", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_KEIYAKU", "Generator\_ColumnPropNameInTable": "KEIYAKU\_KEIYAKUColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| UVA | KEIYAKU\_DEL | {"Generator\_UserColumnName": "KEIYAKU\_DEL", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DEL", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DEL", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DELColumn", "minOccurs": "0"} | [["maxLength", {"value": "4"}]] |
| UVA | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| UVA | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "8"}]] |
| UVA | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| UVA | KAITOKIZITSU\_DT | {"Generator\_UserColumnName": "KAITOKIZITSU\_DT", "nullValue": "\_null", "Generator\_ColumnVarNameInTable": "columnKAITOKIZITSU\_DT", "Generator\_ColumnPropNameInTable": "KAITOKIZITSU\_DTColumn", "Generator\_ColumnPropNameInRow": "KAITOKIZITSU\_DT", "type": "xs:string", "minOccurs": "0"} | [] |
| UVA | KAITOKIZITSU\_DT1 | {"Caption": "KAITOKIZITSU\_DT", "Generator\_UserColumnName": "KAITOKIZITSU\_DT1", "Generator\_ColumnPropNameInRow": "KAITOKIZITSU\_DT1", "Generator\_ColumnVarNameInTable": "columnKAITOKIZITSU\_DT1", "Generator\_ColumnPropNameInTable": "KAITOKIZITSU\_DT1Column", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| UVA | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| UVA | KYUNONYUSAKI\_NK | {"Generator\_UserColumnName": "KYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "KYUNONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnKYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "KYUNONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| UVA | UM\_KIHON\_MITSUMORI\_ID | {"Generator\_UserColumnName": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInRow": "UM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnVarNameInTable": "columnUM\_KIHON\_MITSUMORI\_ID", "Generator\_ColumnPropNameInTable": "UM\_KIHON\_MITSUMORI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| UVA | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| UVA | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| UVA | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | TK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "TK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "TK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "TK\_KEIYAKU\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | IRAIJIGYOSYO\_NK | {"Generator\_UserColumnName": "IRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInRow": "IRAIJIGYOSYO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIJIGYOSYO\_NK", "Generator\_ColumnPropNameInTable": "IRAIJIGYOSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_TK\_KEIYAKU | IRAITANTOSYA | {"Generator\_UserColumnName": "IRAITANTOSYA", "Generator\_ColumnPropNameInRow": "IRAITANTOSYA", "Generator\_ColumnVarNameInTable": "columnIRAITANTOSYA", "Generator\_ColumnPropNameInTable": "IRAITANTOSYAColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_NO | {"Generator\_UserColumnName": "KEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "KEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "KEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | KEIYAKU\_DT | {"Generator\_UserColumnName": "KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUKAISI\_DT | {"Generator\_UserColumnName": "KEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUKAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUKAISI\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUKAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUSYURYO\_DT | {"Generator\_UserColumnName": "KEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUSYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUSYURYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUSYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIDOKOSIN\_FLG | {"Generator\_UserColumnName": "JIDOKOSIN\_FLG", "Generator\_ColumnPropNameInRow": "JIDOKOSIN\_FLG", "Generator\_ColumnVarNameInTable": "columnJIDOKOSIN\_FLG", "Generator\_ColumnPropNameInTable": "JIDOKOSIN\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JIKAIKOSIN\_DT | {"Generator\_UserColumnName": "JIKAIKOSIN\_DT", "Generator\_ColumnPropNameInRow": "JIKAIKOSIN\_DT", "Generator\_ColumnVarNameInTable": "columnJIKAIKOSIN\_DT", "Generator\_ColumnPropNameInTable": "JIKAIKOSIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "KEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "KEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "KEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | ENTYOKEIYAKUMANRYO\_DT | {"Generator\_UserColumnName": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInRow": "ENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnVarNameInTable": "columnENTYOKEIYAKUMANRYO\_DT", "Generator\_ColumnPropNameInTable": "ENTYOKEIYAKUMANRYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | KAIYAKU\_DT | {"Generator\_UserColumnName": "KAIYAKU\_DT", "Generator\_ColumnPropNameInRow": "KAIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnKAIYAKU\_DT", "Generator\_ColumnPropNameInTable": "KAIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | SHONINJOTAI | {"Generator\_UserColumnName": "SHONINJOTAI", "Generator\_ColumnPropNameInRow": "SHONINJOTAI", "Generator\_ColumnVarNameInTable": "columnSHONINJOTAI", "Generator\_ColumnPropNameInTable": "SHONINJOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TK\_KEIYAKU | PACK\_FLG | {"Generator\_UserColumnName": "PACK\_FLG", "Generator\_ColumnPropNameInRow": "PACK\_FLG", "Generator\_ColumnVarNameInTable": "columnPACK\_FLG", "Generator\_ColumnPropNameInTable": "PACK\_FLGColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKUNAIYO | {"Generator\_UserColumnName": "PACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKUNAIYO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKUNAIYO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKUNAIYOColumn", "minOccurs": "0"} | [["maxLength", {"value": "400"}]] |
| MCM\_TK\_KEIYAKU | PACKKEIYAKU\_NO | {"Generator\_UserColumnName": "PACKKEIYAKU\_NO", "Generator\_ColumnPropNameInRow": "PACKKEIYAKU\_NO", "Generator\_ColumnVarNameInTable": "columnPACKKEIYAKU\_NO", "Generator\_ColumnPropNameInTable": "PACKKEIYAKU\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TK\_KEIYAKU | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_KEIYAKU | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_KEIYAKU | SHOKAI\_KEIYAKU\_DT | {"Generator\_UserColumnName": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInRow": "SHOKAI\_KEIYAKU\_DT", "Generator\_ColumnVarNameInTable": "columnSHOKAI\_KEIYAKU\_DT", "Generator\_ColumnPropNameInTable": "SHOKAI\_KEIYAKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAIMEISAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAIMEISAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAIMEISAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAIMEISAI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TK\_SIHARAI\_ID | {"Generator\_UserColumnName": "TK\_SIHARAI\_ID", "Generator\_ColumnPropNameInRow": "TK\_SIHARAI\_ID", "Generator\_ColumnVarNameInTable": "columnTK\_SIHARAI\_ID", "Generator\_ColumnPropNameInTable": "TK\_SIHARAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | TSUKI | {"Generator\_UserColumnName": "TSUKI", "Generator\_ColumnPropNameInRow": "TSUKI", "Generator\_ColumnVarNameInTable": "columnTSUKI", "Generator\_ColumnPropNameInTable": "TSUKIColumn", "minOccurs": "0"} | [["maxLength", {"value": "6"}]] |
| MCM\_TK\_SIHARAIMEISAI | ON\_FLG | {"Generator\_UserColumnName": "ON\_FLG", "Generator\_ColumnPropNameInRow": "ON\_FLG", "Generator\_ColumnVarNameInTable": "columnON\_FLG", "Generator\_ColumnPropNameInTable": "ON\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_DT | {"Generator\_UserColumnName": "ON\_DT", "Generator\_ColumnPropNameInRow": "ON\_DT", "Generator\_ColumnVarNameInTable": "columnON\_DT", "Generator\_ColumnPropNameInTable": "ON\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | ON\_BY | {"Generator\_UserColumnName": "ON\_BY", "Generator\_ColumnPropNameInRow": "ON\_BY", "Generator\_ColumnVarNameInTable": "columnON\_BY", "Generator\_ColumnPropNameInTable": "ON\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_FLG | {"Generator\_UserColumnName": "OFF\_FLG", "Generator\_ColumnPropNameInRow": "OFF\_FLG", "Generator\_ColumnVarNameInTable": "columnOFF\_FLG", "Generator\_ColumnPropNameInTable": "OFF\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_DT | {"Generator\_UserColumnName": "OFF\_DT", "Generator\_ColumnPropNameInRow": "OFF\_DT", "Generator\_ColumnVarNameInTable": "columnOFF\_DT", "Generator\_ColumnPropNameInTable": "OFF\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | OFF\_BY | {"Generator\_UserColumnName": "OFF\_BY", "Generator\_ColumnPropNameInRow": "OFF\_BY", "Generator\_ColumnVarNameInTable": "columnOFF\_BY", "Generator\_ColumnPropNameInTable": "OFF\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TK\_SIHARAIMEISAI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_KEIYAKUJIKAN\_ID | {"Generator\_UserColumnName": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInRow": "TM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_KEIYAKUJIKAN\_ID", "Generator\_ColumnPropNameInTable": "TM\_KEIYAKUJIKAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | TM\_IRAI\_ID | {"Generator\_UserColumnName": "TM\_IRAI\_ID", "Generator\_ColumnPropNameInRow": "TM\_IRAI\_ID", "Generator\_ColumnVarNameInTable": "columnTM\_IRAI\_ID", "Generator\_ColumnPropNameInTable": "TM\_IRAI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_TM\_KEIYAKUJIKAN | JOTAI | {"Generator\_UserColumnName": "JOTAI", "Generator\_ColumnPropNameInRow": "JOTAI", "Generator\_ColumnVarNameInTable": "columnJOTAI", "Generator\_ColumnPropNameInTable": "JOTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_TM\_KEIYAKUJIKAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_TM\_KEIYAKUJIKAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_NONYUSAKI | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_NONYUSAKI | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn"} | [["maxLength", {"value": "12"}]] |
| MCM\_MA\_NONYUSAKI | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_NONYUSAKI | KYUNONYUSAKI\_NK | {"Generator\_UserColumnName": "KYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "KYUNONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnKYUNONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "KYUNONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_NONYUSAKI | NONYUSAKIKOJO\_NK | {"Generator\_UserColumnName": "NONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIKOJO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKOJO\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIKOJO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_NONYUSAKI | NONYUSAKIKANA\_KN | {"Generator\_UserColumnName": "NONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInRow": "NONYUSAKIKANA\_KN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIKANA\_KN", "Generator\_ColumnPropNameInTable": "NONYUSAKIKANA\_KNColumn"} | [["maxLength", {"value": "60"}]] |
| MCM\_MA\_NONYUSAKI | NONYUSAKIEIMEI\_EN | {"Generator\_UserColumnName": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInRow": "NONYUSAKIEIMEI\_EN", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIEIMEI\_EN", "Generator\_ColumnPropNameInTable": "NONYUSAKIEIMEI\_ENColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_MA\_NONYUSAKI | YUBIN\_NO | {"Generator\_UserColumnName": "YUBIN\_NO", "Generator\_ColumnPropNameInRow": "YUBIN\_NO", "Generator\_ColumnVarNameInTable": "columnYUBIN\_NO", "Generator\_ColumnPropNameInTable": "YUBIN\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "8"}]] |
| MCM\_MA\_NONYUSAKI | JUSYO1\_NK | {"Generator\_UserColumnName": "JUSYO1\_NK", "Generator\_ColumnPropNameInRow": "JUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "JUSYO1\_NKColumn"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_NONYUSAKI | JUSYO2\_NK | {"Generator\_UserColumnName": "JUSYO2\_NK", "Generator\_ColumnPropNameInRow": "JUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "JUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_NONYUSAKI | KUNI\_NK | {"Generator\_UserColumnName": "KUNI\_NK", "Generator\_ColumnPropNameInRow": "KUNI\_NK", "Generator\_ColumnVarNameInTable": "columnKUNI\_NK", "Generator\_ColumnPropNameInTable": "KUNI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_MA\_NONYUSAKI | TEL\_NO | {"Generator\_UserColumnName": "TEL\_NO", "Generator\_ColumnPropNameInRow": "TEL\_NO", "Generator\_ColumnVarNameInTable": "columnTEL\_NO", "Generator\_ColumnPropNameInTable": "TEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_NONYUSAKI | FAX\_NO | {"Generator\_UserColumnName": "FAX\_NO", "Generator\_ColumnPropNameInRow": "FAX\_NO", "Generator\_ColumnVarNameInTable": "columnFAX\_NO", "Generator\_ColumnPropNameInTable": "FAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_NONYUSAKI | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_NONYUSAKI | DTSRENKEI\_FLG | {"Generator\_UserColumnName": "DTSRENKEI\_FLG", "Generator\_ColumnPropNameInRow": "DTSRENKEI\_FLG", "Generator\_ColumnVarNameInTable": "columnDTSRENKEI\_FLG", "Generator\_ColumnPropNameInTable": "DTSRENKEI\_FLGColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_MA\_NONYUSAKI | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_NONYUSAKI | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_NONYUSAKI | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_NONYUSAKI | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_MA\_NONYUSAKI | DTSRENKEI\_FLG1 | {"Caption": "DTSRENKEI\_FLG", "Generator\_UserColumnName": "DTSRENKEI\_FLG1", "Generator\_ColumnPropNameInRow": "DTSRENKEI\_FLG1", "Generator\_ColumnVarNameInTable": "columnDTSRENKEI\_FLG1", "Generator\_ColumnPropNameInTable": "DTSRENKEI\_FLG1Column", "type": "xs:decimal"} | [] |
| MCM\_MA\_NONYUSAKI | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn"} | [["maxLength", {"value": "7"}]] |
| MCM\_MA\_NONYUSAKI | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_MA\_NONYUSAKI | NONYU\_DT | {"Generator\_UserColumnName": "NONYU\_DT", "Generator\_ColumnPropNameInRow": "NONYU\_DT", "Generator\_ColumnVarNameInTable": "columnNONYU\_DT", "Generator\_ColumnPropNameInTable": "NONYU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_NONYUSAKI | HOSYUSYUSOKU\_DT | {"Generator\_UserColumnName": "HOSYUSYUSOKU\_DT", "Generator\_ColumnPropNameInRow": "HOSYUSYUSOKU\_DT", "Generator\_ColumnVarNameInTable": "columnHOSYUSYUSOKU\_DT", "Generator\_ColumnPropNameInTable": "HOSYUSYUSOKU\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_NONYUSAKI | TEKKYO\_DT | {"Generator\_UserColumnName": "TEKKYO\_DT", "Generator\_ColumnPropNameInRow": "TEKKYO\_DT", "Generator\_ColumnVarNameInTable": "columnTEKKYO\_DT", "Generator\_ColumnPropNameInTable": "TEKKYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_MA\_NONYUSAKI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_NONYUSAKI | KIKIKOSEILINK | {"Generator\_UserColumnName": "KIKIKOSEILINK", "Generator\_ColumnPropNameInRow": "KIKIKOSEILINK", "Generator\_ColumnVarNameInTable": "columnKIKIKOSEILINK", "Generator\_ColumnPropNameInTable": "KIKIKOSEILINKColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_MA\_BRAND\_KOSEI | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRAND\_ID | {"Generator\_UserColumnName": "BRAND\_ID", "Generator\_ColumnPropNameInRow": "BRAND\_ID", "Generator\_ColumnVarNameInTable": "columnBRAND\_ID", "Generator\_ColumnPropNameInTable": "BRAND\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_MA\_BRAND\_KOSEI | SURYO | {"Generator\_UserColumnName": "SURYO", "Generator\_ColumnVarNameInTable": "columnSURYO", "Generator\_ColumnPropNameInRow": "SURYO", "Generator\_ColumnPropNameInTable": "SURYOColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_MA\_BRAND\_KOSEI | BRAND\_NK | {"Generator\_UserColumnName": "BRAND\_NK", "Generator\_ColumnVarNameInTable": "columnBRAND\_NK", "Generator\_ColumnPropNameInRow": "BRAND\_NK", "Generator\_ColumnPropNameInTable": "BRAND\_NKColumn"} | [["maxLength", {"value": "80"}]] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1">
      <xs:selector xpath=".//mstns:MCM_2004_V" />
      <xs:field xpath="mstns:UM_KIHON_MITSUMORI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UK_KEIYAKU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_KEIYAKU" />
      <xs:field xpath="mstns:UK_KEIYAKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_UM_KIHON_MITSUMORI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UM_KIHON_MITSUMORI" />
      <xs:field xpath="mstns:UM_KIHON_MITSUMORI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_KEIYAKU_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_KEIYAKU" />
      <xs:field xpath="mstns:TK_KEIYAKU_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TK_SIHARAIMEISAI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TK_SIHARAIMEISAI" />
      <xs:field xpath="mstns:TK_SIHARAIMEISAI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_TM_KEIYAKUJIKAN_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_TM_KEIYAKUJIKAN" />
      <xs:field xpath="mstns:TM_KEIYAKUJIKAN_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="MCM_MA_NONYUSAKI_Constraint1" ns1:ConstraintName="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_MA_NONYUSAKI" />
      <xs:field xpath="mstns:NONYUSAKI_ID" />
    </xs:unique>
    
```

```xml
<xs:unique xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint2">
      <xs:selector xpath=".//mstns:MCM_MA_NONYUSAKI" />
      <xs:field xpath="mstns:NONYUSAKI_ID" />
      <xs:field xpath="mstns:SUPPORT_ID" />
    </xs:unique>
    
```

```xml
<xs:keyref xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="FK_MCM_2004_V_UVA" refer="Constraint1" ns1:rel_Generator_UserRelationName="FK_MCM_2004_V_UVA" ns1:rel_Generator_RelationVarName="relationFK_MCM_2004_V_UVA" ns1:rel_Generator_UserChildTable="UVA" ns1:rel_Generator_UserParentTable="MCM_2004_V" ns1:rel_Generator_ParentPropName="MCM_2004_VRow" ns1:rel_Generator_ChildPropName="GetUVARows">
      <xs:selector xpath=".//mstns:UVA" />
      <xs:field xpath="mstns:UM_KIHON_MITSUMORI_ID" />
    </xs:keyref>
    
```

```xml
<xs:keyref xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="FK_MCM_MA_NONYUSAKI_MCM_2004_V" refer="Constraint2" ns1:rel_Generator_UserRelationName="FK_MCM_MA_NONYUSAKI_MCM_2004_V" ns1:rel_Generator_RelationVarName="relationFK_MCM_MA_NONYUSAKI_MCM_2004_V" ns1:rel_Generator_UserChildTable="MCM_2004_V" ns1:rel_Generator_UserParentTable="MCM_MA_NONYUSAKI" ns1:rel_Generator_ParentPropName="MCM_MA_NONYUSAKIRowParent" ns1:rel_Generator_ChildPropName="GetMCM_2004_VRows">
      <xs:selector xpath=".//mstns:MCM_2004_V" />
      <xs:field xpath="mstns:NONYUSAKI_ID" />
      <xs:field xpath="mstns:SUPPORT_ID" />
    </xs:keyref>
  
```

```xml
<ns0:Relationship xmlns:ns0="urn:schemas-microsoft-com:xml-msdata" xmlns:ns1="urn:schemas-microsoft-com:xml-msprop" name="FK_MCM_MA_NONYUSAKI_MCM_MA_BRAND_KOSEI" ns0:parent="MCM_MA_NONYUSAKI" ns0:child="MCM_MA_BRAND_KOSEI" ns0:parentkey="PLANT_ID" ns0:childkey="PLANT_ID" ns1:Generator_UserRelationName="FK_MCM_MA_NONYUSAKI_MCM_MA_BRAND_KOSEI" ns1:Generator_RelationVarName="relationFK_MCM_MA_NONYUSAKI_MCM_MA_BRAND_KOSEI" ns1:Generator_UserChildTable="MCM_MA_BRAND_KOSEI" ns1:Generator_UserParentTable="MCM_MA_NONYUSAKI" ns1:Generator_ParentPropName="MCM_MA_NONYUSAKIRow" ns1:Generator_ChildPropName="GetMCM_MA_BRAND_KOSEIRows" />
    
```

</details>

### 3008　Mcm3008uDataSet.xsd

[Mcm3008uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uDataSet.xsd>)

| テーブル | 列 | 型・省略・その他属性 | 長さ・その他制約 |
| --- | --- | --- | --- |
| MCM\_UK\_SEIBAN | UK\_SEIBAN\_ID | {"Generator\_UserColumnName": "UK\_SEIBAN\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_SEIBAN\_ID", "Generator\_ColumnPropNameInRow": "UK\_SEIBAN\_ID", "Generator\_ColumnPropNameInTable": "UK\_SEIBAN\_IDColumn", "type": "xs:decimal"} | [] |
| MCM\_UK\_SEIBAN | UK\_KEIYAKU\_ID | {"Generator\_UserColumnName": "UK\_KEIYAKU\_ID", "Generator\_ColumnVarNameInTable": "columnUK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInRow": "UK\_KEIYAKU\_ID", "Generator\_ColumnPropNameInTable": "UK\_KEIYAKU\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | HARD\_SEIBAN | {"Generator\_UserColumnName": "HARD\_SEIBAN", "Generator\_ColumnVarNameInTable": "columnHARD\_SEIBAN", "Generator\_ColumnPropNameInRow": "HARD\_SEIBAN", "Generator\_ColumnPropNameInTable": "HARD\_SEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UK\_SEIBAN | SOFT\_SEIBAN | {"Generator\_UserColumnName": "SOFT\_SEIBAN", "Generator\_ColumnVarNameInTable": "columnSOFT\_SEIBAN", "Generator\_ColumnPropNameInRow": "SOFT\_SEIBAN", "Generator\_ColumnPropNameInTable": "SOFT\_SEIBANColumn", "minOccurs": "0"} | [["maxLength", {"value": "20"}]] |
| MCM\_UK\_SEIBAN | KAISI\_DT | {"Generator\_UserColumnName": "KAISI\_DT", "Generator\_ColumnVarNameInTable": "columnKAISI\_DT", "Generator\_ColumnPropNameInRow": "KAISI\_DT", "Generator\_ColumnPropNameInTable": "KAISI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | SYURYO\_DT | {"Generator\_UserColumnName": "SYURYO\_DT", "Generator\_ColumnVarNameInTable": "columnSYURYO\_DT", "Generator\_ColumnPropNameInRow": "SYURYO\_DT", "Generator\_ColumnPropNameInTable": "SYURYO\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | BIKO | {"Generator\_UserColumnName": "BIKO", "Generator\_ColumnVarNameInTable": "columnBIKO", "Generator\_ColumnPropNameInRow": "BIKO", "Generator\_ColumnPropNameInTable": "BIKOColumn", "minOccurs": "0"} | [["maxLength", {"value": "4000"}]] |
| MCM\_UK\_SEIBAN | CREATED\_DT | {"Generator\_UserColumnName": "CREATED\_DT", "Generator\_ColumnVarNameInTable": "columnCREATED\_DT", "Generator\_ColumnPropNameInRow": "CREATED\_DT", "Generator\_ColumnPropNameInTable": "CREATED\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | CREATED\_BY | {"Generator\_UserColumnName": "CREATED\_BY", "Generator\_ColumnVarNameInTable": "columnCREATED\_BY", "Generator\_ColumnPropNameInRow": "CREATED\_BY", "Generator\_ColumnPropNameInTable": "CREATED\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | LASTUPDATE\_DT | {"Generator\_UserColumnName": "LASTUPDATE\_DT", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_DT", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_DT", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | LASTUPDATE\_BY | {"Generator\_UserColumnName": "LASTUPDATE\_BY", "Generator\_ColumnVarNameInTable": "columnLASTUPDATE\_BY", "Generator\_ColumnPropNameInRow": "LASTUPDATE\_BY", "Generator\_ColumnPropNameInTable": "LASTUPDATE\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | KAKUNIN\_DT | {"Generator\_UserColumnName": "KAKUNIN\_DT", "Generator\_ColumnVarNameInTable": "columnKAKUNIN\_DT", "Generator\_ColumnPropNameInRow": "KAKUNIN\_DT", "Generator\_ColumnPropNameInTable": "KAKUNIN\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | KAKUNIN\_BY | {"Generator\_UserColumnName": "KAKUNIN\_BY", "Generator\_ColumnVarNameInTable": "columnKAKUNIN\_BY", "Generator\_ColumnPropNameInRow": "KAKUNIN\_BY", "Generator\_ColumnPropNameInTable": "KAKUNIN\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | KAKUNIN\_KBN | {"Generator\_UserColumnName": "KAKUNIN\_KBN", "Generator\_ColumnVarNameInTable": "columnKAKUNIN\_KBN", "Generator\_ColumnPropNameInRow": "KAKUNIN\_KBN", "Generator\_ColumnPropNameInTable": "KAKUNIN\_KBNColumn", "minOccurs": "0"} | [["maxLength", {"value": "1"}]] |
| MCM\_UK\_SEIBAN | NONYUBUSYO\_NK | {"Generator\_UserColumnName": "NONYUBUSYO\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUBUSYO\_NK", "Generator\_ColumnPropNameInRow": "NONYUBUSYO\_NK", "Generator\_ColumnPropNameInTable": "NONYUBUSYO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "60"}]] |
| MCM\_UK\_SEIBAN | NONYUSAKI\_ID | {"Generator\_UserColumnName": "NONYUSAKI\_ID", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_ID", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_ID", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | NONYUSAKI\_CD | {"Generator\_UserColumnName": "NONYUSAKI\_CD", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_CD", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_CD", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_CDColumn", "minOccurs": "0"} | [["maxLength", {"value": "12"}]] |
| MCM\_UK\_SEIBAN | NONYUSAKI\_NK | {"Generator\_UserColumnName": "NONYUSAKI\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKI\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKI\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKI\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_SEIBAN | NONYUSAKIJUSYO1\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO1\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_SEIBAN | NONYUSAKIJUSYO2\_NK | {"Generator\_UserColumnName": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInRow": "NONYUSAKIJUSYO2\_NK", "Generator\_ColumnPropNameInTable": "NONYUSAKIJUSYO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_SEIBAN | HOSYU\_GKIN | {"Generator\_UserColumnName": "HOSYU\_GKIN", "Generator\_ColumnVarNameInTable": "columnHOSYU\_GKIN", "Generator\_ColumnPropNameInRow": "HOSYU\_GKIN", "Generator\_ColumnPropNameInTable": "HOSYU\_GKINColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | PLANT\_ID | {"Generator\_UserColumnName": "PLANT\_ID", "Generator\_ColumnVarNameInTable": "columnPLANT\_ID", "Generator\_ColumnPropNameInRow": "PLANT\_ID", "Generator\_ColumnPropNameInTable": "PLANT\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | SUPPORT\_ID | {"Generator\_UserColumnName": "SUPPORT\_ID", "Generator\_ColumnVarNameInTable": "columnSUPPORT\_ID", "Generator\_ColumnPropNameInRow": "SUPPORT\_ID", "Generator\_ColumnPropNameInTable": "SUPPORT\_IDColumn", "minOccurs": "0"} | [["maxLength", {"value": "7"}]] |
| MCM\_UK\_SEIBAN | PLANT\_NK | {"Generator\_UserColumnName": "PLANT\_NK", "Generator\_ColumnVarNameInTable": "columnPLANT\_NK", "Generator\_ColumnPropNameInRow": "PLANT\_NK", "Generator\_ColumnPropNameInTable": "PLANT\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "80"}]] |
| MCM\_UK\_SEIBAN | NONYUTANTOSYA\_NK | {"Generator\_UserColumnName": "NONYUTANTOSYA\_NK", "Generator\_ColumnVarNameInTable": "columnNONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInRow": "NONYUTANTOSYA\_NK", "Generator\_ColumnPropNameInTable": "NONYUTANTOSYA\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "40"}]] |
| MCM\_UK\_SEIBAN | NONYUTEL\_NO | {"Generator\_UserColumnName": "NONYUTEL\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUTEL\_NO", "Generator\_ColumnPropNameInRow": "NONYUTEL\_NO", "Generator\_ColumnPropNameInTable": "NONYUTEL\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | NONYUFAX\_NO | {"Generator\_UserColumnName": "NONYUFAX\_NO", "Generator\_ColumnVarNameInTable": "columnNONYUFAX\_NO", "Generator\_ColumnPropNameInRow": "NONYUFAX\_NO", "Generator\_ColumnPropNameInTable": "NONYUFAX\_NOColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | KEIYAKUJIKANTAI | {"Generator\_UserColumnName": "KEIYAKUJIKANTAI", "Generator\_ColumnVarNameInTable": "columnKEIYAKUJIKANTAI", "Generator\_ColumnPropNameInRow": "KEIYAKUJIKANTAI", "Generator\_ColumnPropNameInTable": "KEIYAKUJIKANTAIColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_SEIBAN | HOSYUHOHO | {"Generator\_UserColumnName": "HOSYUHOHO", "Generator\_ColumnVarNameInTable": "columnHOSYUHOHO", "Generator\_ColumnPropNameInRow": "HOSYUHOHO", "Generator\_ColumnPropNameInTable": "HOSYUHOHOColumn", "minOccurs": "0"} | [["maxLength", {"value": "2"}]] |
| MCM\_UK\_SEIBAN | IRAITENPO\_ID | {"Generator\_UserColumnName": "IRAITENPO\_ID", "Generator\_ColumnVarNameInTable": "columnIRAITENPO\_ID", "Generator\_ColumnPropNameInRow": "IRAITENPO\_ID", "Generator\_ColumnPropNameInTable": "IRAITENPO\_IDColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | IRAIMEISHO1\_NK | {"Generator\_UserColumnName": "IRAIMEISHO1\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO1\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO1\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO1\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_SEIBAN | IRAIMEISHO2\_NK | {"Generator\_UserColumnName": "IRAIMEISHO2\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO2\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO2\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO2\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_SEIBAN | IRAIMEISHO3\_NK | {"Generator\_UserColumnName": "IRAIMEISHO3\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO3\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO3\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO3\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_SEIBAN | IRAIMEISHO4\_NK | {"Generator\_UserColumnName": "IRAIMEISHO4\_NK", "Generator\_ColumnVarNameInTable": "columnIRAIMEISHO4\_NK", "Generator\_ColumnPropNameInRow": "IRAIMEISHO4\_NK", "Generator\_ColumnPropNameInTable": "IRAIMEISHO4\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_SEIBAN | IRAITENPORYAKU\_NK | {"Generator\_UserColumnName": "IRAITENPORYAKU\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInRow": "IRAITENPORYAKU\_NK", "Generator\_ColumnPropNameInTable": "IRAITENPORYAKU\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_SEIBAN | IRAITANTO\_NK | {"Generator\_UserColumnName": "IRAITANTO\_NK", "Generator\_ColumnVarNameInTable": "columnIRAITANTO\_NK", "Generator\_ColumnPropNameInRow": "IRAITANTO\_NK", "Generator\_ColumnPropNameInTable": "IRAITANTO\_NKColumn", "minOccurs": "0"} | [["maxLength", {"value": "100"}]] |
| MCM\_UK\_SEIBAN | KAKUNINIRAI\_DT | {"Generator\_UserColumnName": "KAKUNINIRAI\_DT", "Generator\_ColumnVarNameInTable": "columnKAKUNINIRAI\_DT", "Generator\_ColumnPropNameInRow": "KAKUNINIRAI\_DT", "Generator\_ColumnPropNameInTable": "KAKUNINIRAI\_DTColumn", "type": "xs:dateTime", "minOccurs": "0"} | [] |
| MCM\_UK\_SEIBAN | KAKUNINIRAI\_BY | {"Generator\_UserColumnName": "KAKUNINIRAI\_BY", "Generator\_ColumnVarNameInTable": "columnKAKUNINIRAI\_BY", "Generator\_ColumnPropNameInRow": "KAKUNINIRAI\_BY", "Generator\_ColumnPropNameInTable": "KAKUNINIRAI\_BYColumn", "minOccurs": "0"} | [["maxLength", {"value": "50"}]] |
| MCM\_UK\_SEIBAN | CHECK\_FLG | {"Generator\_UserColumnName": "CHECK\_FLG", "Generator\_ColumnPropNameInRow": "CHECK\_FLG", "Generator\_ColumnVarNameInTable": "columnCHECK\_FLG", "Generator\_ColumnPropNameInTable": "CHECK\_FLGColumn", "type": "xs:decimal", "minOccurs": "0"} | [] |

<details>
<summary>キー・関連の定義原文</summary>

```xml
<xs:unique xmlns:ns1="urn:schemas-microsoft-com:xml-msdata" xmlns:xs="http://www.w3.org/2001/XMLSchema" name="Constraint1" ns1:PrimaryKey="true">
      <xs:selector xpath=".//mstns:MCM_UK_SEIBAN" />
      <xs:field xpath="mstns:UK_SEIBAN_ID" />
    </xs:unique>
  
```

</details>

## 判定を保留する範囲

- 業務137項目・共通10項目の確認先を、単体150件・結合207件へ対応づけました。追加単体・未共有7画面・条件の全組合せ・要合意事項はVB機能照合に記録。対応先の存在だけで条件網羅や実行合格とは判定しません。
- 実DBのビュー・関数・ストアド、共有保存先・SMTP・Office・利用者設定・複数利用者の競合は未確認。
- 共通全クラス・全イベントと、関連P帳票の明細・テンプレート内部まで全件展開した一覧ではありません。今回は13U画面の条件と、直接関係する共通処理を確認。
- 2006Uの保守方法集約など、コメントと実コードの相違は不整合候補。VBの不具合をWebへ再現する要件にはせず、期待を合意してから仕様書へ反映。

## 追加確認した原本

| 原本 | SHA-256 |
| --- | --- |
| [CPCoreUserControl.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPCoreUserControl.vb>) | 5596bb4c336f743deacd084abedd7888a0db42aac23b812188521920a8548c72 |
| [CPMessageUtility.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPMessageUtility.vb>) | e388173cc264248e4b9138bcdf2be66b313c15b47a18e70461ad93373a06674a |
| [CPMessageConstant.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONSTANT/CPMessageConstant.vb>) | 209d9e9cbbf1757726bbb4cffb639afd08539d357ae90a2be6db29b35175a91c |
| [CPBaseForm.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPBaseForm.vb>) | 39c41e599bafbefb122b7b0615e0178d9aceb1ad13f3947388b5937e29124267 |
| [CPUserInfo.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/INFO/CPUserInfo.vb>) | 64de352d4755fb26fdac8317140b611ab24576ab4eb3b5c208e39e0dd8c87bbb |
| [CPValidateUtilty.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPValidateUtilty.vb>) | 137163759736eeb5034798f9c4b0a64ff100baa0086e3b2faed9680b915e754a |
| [CPDataGridView.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPDataGridView.vb>) | 81562aec0bc8a54d853552bf0104d5f1c03be57ac1c35e6865187ee5ede376e8 |
| [CPTextBox.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPTextBox.vb>) | 221b4a364a4648dc8d6a888b80277ea81aba329065262e38452a0fd2b159f788 |
| [CPCommonUtility.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPCommonUtility.vb>) | ea2b0a40901b30b5fde647dcb4fe1a62963b0b7aa94f9d6dc4ae5069499afd3d |
| [CPBaseUserControl.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONTAINER/CPBaseUserControl.vb>) | 9c7178ed79fdbc35a6ba22782a4a53a60410cf0274421fc84f3986df0dbaeaae |
| [CPButton.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPButton.vb>) | 0b24bc3afe077f7f2f25b9e7ceb4e8c5bd6da8904a545797d763cd735d9943ad |
| [CPComponentUtility.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPComponentUtility.vb>) | 4cf9692d25c359cd9af8288f4a3f3b4c927ad3dda0485dc5dd3e8c7edd8be091 |
| [CPExcelManager.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/EXCEL/CPExcelManager.vb>) | 0a285f9107c6738feb13d5be41aee287e6a0a91173d51ca690e3dc2c5455ed4b |
| [CPConstant.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CONSTANT/CPConstant.vb>) | 91ac3247c2c33105413b4bae5fb0ca750f093043fa27ade1155ff34fc5729dfc |
| [CPDBUtility.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/UTILITY/CPDBUtility.vb>) | c29edff1c125e9fa515e69b88245baf89af1325bc28266e3db6f070e7aeb7b6a |
| [CPFileDialogButton.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/COMPONENT/CPFileDialogButton.vb>) | 192b106575e7591b62be38c398593a526ef16c254955617c09d32901fa6a16f4 |
| [CPValidate.xml](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/10_FRAMEWORK/CPValidate.xml>) | 55bcd43cec8a2cbfccd63d8fbc98bd880345c0840fd0381340e2622999442e22 |
| [Mcm1005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uScreen.Designer.vb>) | e672eca54c7652b3fe1e2e2a9da6f07fe192bd59d54a81fdde28cad195879935 |
| [Mcm1005uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uTabControl.Designer.vb>) | bc9d133e7e708fbf8b5c64214c03d340d0f3b893efdb50913ec97aaedf1dc410 |
| [Mcm1006u1Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1Screen.Designer.vb>) | 7338e7c661a9fba03dfbae317c72433a66a9f7a164b88095d7805c15e5fe48db |
| [Mcm1006u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2Screen.Designer.vb>) | e6eb808befa81e48b3bc1d4afd46d1d11a53ada9b7a5eececbc3946322f4fa71 |
| [Mcm1009uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uScreen.designer.vb>) | 764e887fed4e66315689463fddb0850071846aa6c7ba09fb2c2a71ad4fae7793 |
| [Mcm1010uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uScreen.designer.vb>) | d9deee99e29bb9feea94652190358dfb18b0316a07dd82a5c58f1933c037d84f |
| [Mcm1011uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uScreen.designer.vb>) | 43e0ef72ea45ef75163cd629feedcffbf911e2c92ef71f99e89d447dc9b2304f |
| [Mcm2006uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uScreen.Designer.vb>) | 26207144a25e1ef3967fac397fca36bcd8c093820973fdf9f07da84f66136bbc |
| [Mcm2006uTabControl.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uTabControl.Designer.vb>) | 68cc2d5441ab0dd05e136aeee1c7ef065f13c92f6c098fed64d3ee41282a764a |
| [Mcm2007u2Screen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2Screen.Designer.vb>) | fc1839b33c97125816628706cbe93dae68d43db21d87287b96942e89f9a5e291 |
| [Mcm2007uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uScreen.Designer.vb>) | 94a82dccafc9a9ba5fb2185140f9e7cd7ebc0c3226a1cc424304fb5a7916cd3e |
| [Mcm3002uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3002U/Mcm3002uScreen.Designer.vb>) | ec20bc68cfce9a731429f6fa977005fa454bb3ed2d4033527dbefbefbedb3ddb |
| [Mcm3003uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3003U/Mcm3003uScreen.Designer.vb>) | 063dbcc907f294d284d63873c434c5cdf1ca767b283c45b47c45ab24604e92d8 |
| [Mcm3004uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uScreen.Designer.vb>) | 3ce83e1c395124f2211f40c5f672f4a021dac44e2a6a40fd75ba5a82274142da |
| [Mcm3005uScreen.Designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uScreen.Designer.vb>) | d9eb7bd4c569d3da7b5c0a98428d3b8f1c57fdac3e444d8989a7a0a1ba8bff93 |
| [Mcm3007uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uScreen.designer.vb>) | a6f9cb771b45d50a5ef76214696759fabe975e5885bfaa3a5076ea714c053e1a |
| [Mcm3008uScreen.designer.vb](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uScreen.designer.vb>) | 60514c5ab0da7408c9e29b3e7019135dd2cfe872979148fa9beb8a8169aaa66a |
| [Mcm1005uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1005U/Mcm1005uDataSet.xsd>) | c05319d8b19f35b352b30c8ac75da1405fa5fd686fc3facceae8ca16caebb277 |
| [Mcm1006u1DataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u1DataSet.xsd>) | a77988d5db8055ce2e40902d4077dd02c8c26b3ca86bd8b806d8dbbd237a1590 |
| [Mcm1006u2DataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1006U/Mcm1006u2DataSet.xsd>) | b6e3336b5bf3058230656687f947dfb2c96fb62a2b48a9226a1f29547b319a13 |
| [Mcm1009uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1009U/Mcm1009uDataSet.xsd>) | 4513512d3a3942996a251fb25e49a13560887d73284bdc79ef1ebd5f25b8d39f |
| [Mcm1010uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1010U/Mcm1010uDataSet.xsd>) | aa184b3cbe25045a61da7d34c5af468da7d7cff23920afc22f2761e536e708a1 |
| [Mcm1011uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/20_CONSUMER/MCM1011U/Mcm1011uDataSet.xsd>) | fbe06972d4cea5f7e1febeb34755c14a0d62382f925ebadc07917d6019613ee3 |
| [Mcm2006uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2006U/Mcm2006uDataSet.xsd>) | a9ff9926372173d87094f39efa9f09d517c1273bfc6eddaa3ecdc00980933fd5 |
| [Mcm2007u2DataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007u2DataSet.xsd>) | 169eb7e80de875ae9ca92b4286ccd3b620e76cbc9112da86b117f835dfa68946 |
| [Mcm2007uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/30_CUSTOMER/MCM2007U/Mcm2007uDataSet.xsd>) | e8426a16c0ddd47483b7a88ea4c968863a92856171e05d03ba881750b67bced9 |
| [Mcm3002uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3002U/Mcm3002uDataSet.xsd>) | b58fcf52d8e68e42380f24d16e6fe5b28482ea2cea0dfe71e75ee602a999b9d0 |
| [Mcm3003uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3003U/Mcm3003uDataSet.xsd>) | e45248d809993d555d8f2eb94a4abfa08ace66d2a2519c3b3777abb3e10b37d5 |
| [Mcm3004uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3004U/Mcm3004uDataSet.xsd>) | e61565f61f18eeaf195e00b998baac1000e438ea86ec08150e889b5fd072f7a3 |
| [Mcm3005uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3005U/Mcm3005uDataSet.xsd>) | 128e2cb616825f4bf3c25ed967a43c5f60e8e1c0effe9f3ad8991e06b56d81cf |
| [Mcm3007uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3007U/Mcm3007uDataSet.xsd>) | 93215a4d782dda53d2dc1d49a4dfd2da83e26ca6265ed0dfebf434d98a20ab3f |
| [Mcm3008uDataSet.xsd](<../../参照用/MCM_javaコンバート用/MCM_作業用/MCM_作業用/30_APPLICATION/40_ETC/MCM3008U/Mcm3008uDataSet.xsd>) | 65471d75fd0e774d28863f4b0dc8eab61fd52d69da889bd16e14109ea82edb19 |
