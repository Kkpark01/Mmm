'保守契約管理システムサポートID付け替えフォーム
Public Class frmSupportIdTukekae

    '旧画面の枠線をWindows Formsの描画で維持する。
    Protected Overrides Sub OnPaint(e As Global.System.Windows.Forms.PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.DrawRectangle(Global.System.Drawing.Pens.Black, 23, 98, 593, 84)
    End Sub

    'DBマネージャー
    Public Shared Property dbManager() As utlDbManager

    '開始ボタン押下時処理
    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        Dim res As New utlResult
        Dim syori As String = sender.Text

        ClearFormMsg()
        SetFormMsg(syori & "確認中です")
        If ShowMsgQuestion(syori) = Global.System.Windows.Forms.DialogResult.OK Then

            'DB接続確認
            dbManager = New utlDbManager
            res = dbManager.Connect
            If res.Result = utlResult.ResultType.Err Then
                res.MsgTxt = "DB接続確認が失敗した為、処理を中止します。" & vbCrLf & vbCrLf & "【エラー内容】" & vbCrLf & res.MsgTxt
            Else
                'サポートID付け替え実行
                ExcuteTukekae(Me.txtMotoSupportId, Me.txtSakiSupportId, Me.txtMotoNonyuSaki, Me.txtSakiNonyuSaki)
            End If

        Else
            SetFormMsg(syori & "をキャンセルしました")
        End If

        'DB切断
        If frmSupportIdTukekae.dbManager IsNot Nothing Then frmSupportIdTukekae.dbManager.DisConnect()

    End Sub

    '終了ボタン押下時処理
    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        Dim syori As String = sender.Text
        ClearFormMsg()
        SetFormMsg(syori & "確認中です")
        If ShowMsgQuestion(syori) = Global.System.Windows.Forms.DialogResult.OK Then
            '画面を閉じる
            Me.Close()
        Else
            SetFormMsg(syori & "をキャンセルしました")
        End If
    End Sub

    '付け替え元サポートIDフォーカス喪失時処理
    Private Sub txtMotoSupportId_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMotoSupportId.Validating
        Dim res As New utlResult

        'DB接続確認
        dbManager = New utlDbManager
        res = dbManager.Connect
        If res.Result = utlResult.ResultType.Err Then
            res.MsgTxt = "DB接続確認が失敗した為、処理を中止します。" & vbCrLf & vbCrLf & "【エラー内容】" & vbCrLf & res.MsgTxt
        Else
            '付け替え元サポートIDチェック
            ClearFormMsg()
            CheckMotoSupportId(txtMotoSupportId, txtMotoNonyuSaki)
        End If

        'DB切断
        If frmSupportIdTukekae.dbManager IsNot Nothing Then frmSupportIdTukekae.dbManager.DisConnect()

    End Sub

    '付け替え先サポートIDフォーカス喪失時処理
    Private Sub txtSakiSupportId_Validating(sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSakiSupportId.Validating
        Dim res As New utlResult

        'DB接続確認
        dbManager = New utlDbManager
        res = dbManager.Connect
        If res.Result = utlResult.ResultType.Err Then
            res.MsgTxt = "DB接続確認が失敗した為、処理を中止します。" & vbCrLf & vbCrLf & "【エラー内容】" & vbCrLf & res.MsgTxt
        Else
            '付け替え先サポートIDチェック
            ClearFormMsg()
            CheckSakiSupportId(txtMotoSupportId, txtSakiSupportId, txtSakiNonyuSaki)
        End If

        'DB切断
        If frmSupportIdTukekae.dbManager IsNot Nothing Then frmSupportIdTukekae.dbManager.DisConnect()

    End Sub

End Class