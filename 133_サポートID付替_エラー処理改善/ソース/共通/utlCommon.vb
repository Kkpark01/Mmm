Imports System.IO

'共通処理モジュール
Public Module utlCommon

    '画面初期化
    Public Function GamenInit(ByRef dbManager As utlDbManager) As utlResult
        Dim res As New utlResult

        'DB接続チェック
        Dim dbm As New utlDbManager
        res = dbm.Connect
        dbm.DisConnect()
        dbm.Dispose()
        If res.Result = utlResult.ResultType.Err Then
            res.MsgTxt = "DB接続確認が失敗した為、画面起動を中止します。" & vbCrLf & vbCrLf & "【エラー内容】" & vbCrLf & res.MsgTxt
            GamenInit = res
            Exit Function
        End If

        'ログファイルチェック
        res = CheckLogFile()
        If res.Result = utlResult.ResultType.Err Then
            res.MsgTxt = "ログファイル確認が失敗した為、画面起動を中止します。" & vbCrLf & vbCrLf & "【エラー内容】" & vbCrLf & res.MsgTxt
            GamenInit = res
            Exit Function
        End If

        '画面初期化成功時
        res.SetSuccess("画面起動成功")
        GamenInit = res

    End Function

    'ログチェック
    Public Function CheckLogFile() As utlResult
        Dim res As New utlResult
        Dim fileInf As New FileInfo(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\" & "log.txt")
        Dim folder As String = fileInf.DirectoryName
        Dim file As String = fileInf.Name

        'フォルダ存在チェック
        If Not System.IO.Directory.Exists(folder) Then
            res.SetError("ログフォルダ「" & folder & "」が存在しません。")
        Else
            'ファイル使用中チェック
            If fileInf.Exists Then
                Try
                    'ファイル名を変更し、ファイル名を元に戻せれば使用中ではない
                    fileInf.MoveTo(Path.Combine(folder, file & ".bk"))
                    fileInf.MoveTo(Path.Combine(folder, file))
                    res.SetSuccess()
                Catch ex As Exception
                    'ファイル名が変更できないので、使用中
                    res.SetError("ログファイル[" & fileInf.FullName & "]が開かれている可能性があります。")
                End Try
            Else
                'ファイルが存在しなければ、使用中ではない
                res.SetSuccess()
            End If
        End If

        CheckLogFile = res
    End Function

    'メッセージクリア
    Public Sub ClearFormMsg()
        frmSupportIdTukekae.txtMessage.Text = ""
        frmSupportIdTukekae.txtMessage.Refresh()
    End Sub

    'メッセージ設定
    Public Sub SetFormMsg(ByVal msg As String)
        frmSupportIdTukekae.txtMessage.Text = msg
        frmSupportIdTukekae.txtMessage.Refresh()
    End Sub

    'ログ出力
    Public Function LogWrite(ByVal Message As String) As Integer
        Dim MsgWk As String
        Dim textFile As System.IO.StreamWriter
        MsgWk = System.DateTime.Now.ToString & "【" & Message & "】"
        textFile = New System.IO.StreamWriter(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\" & "log.txt", True, System.Text.Encoding.Default)
        textFile.WriteLine(MsgWk)
        textFile.Close()
        Return True
    End Function

    '確認メッセージ表示
    Public Function ShowMsgQuestion(ByVal syori As String) As Global.System.Windows.Forms.DialogResult
        ShowMsgQuestion = MessageBox.Show(syori & "してよろしいですか？", syori & "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question)
    End Function

    '成功メッセージ表示
    Public Sub ShowMsgSuccess(ByVal msg As String)
        MessageBox.Show(msg, "正常終了", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    'エラーメッセージ表示
    Public Sub ShowMsgError(ByVal msg As String)
        MessageBox.Show(msg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    '必須チェック
    Public Function CheckInputRequired(ByRef target As Object) As Boolean
        Dim check As Boolean = True

        '空かどうか
        If target Is Nothing Then
            check = False
        End If

        'NULLかどうか
        If IsDBNull(target) Then
            check = False
        End If

        '入力なしまたはブランクかどうか
        If target.ToString.Trim = "" Then
            check = False
        End If

        CheckInputRequired = check
    End Function

End Module

