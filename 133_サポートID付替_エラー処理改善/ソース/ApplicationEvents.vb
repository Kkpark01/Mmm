Namespace My

    'アプリケーションイベントクラス
    Partial Friend Class MyApplication

        'アプリケーション起動
        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            Dim res As New utlResult

            '画面初期化
            res = GamenInit(frmSupportIdTukekae.dbManager)
            If res.Result = utlResult.ResultType.Err Then
                res.ShowMsg()
                e.Cancel = True
            End If

            '画面中央に表示
            frmSupportIdTukekae.StartPosition = FormStartPosition.CenterScreen
        End Sub

        'アプリケーション二重起動禁止
        Private Sub MyApplication_StartupNextInstance(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.StartupNextInstanceEventArgs) Handles Me.StartupNextInstance
            ShowMsgError("すでに起動しています。")
        End Sub

        '補足されなかった例外エラー発生時
        Private Sub MyApplication_UnhandledException(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs) Handles Me.UnhandledException
            'e.ExitApplicationをTrueにすると、アプリケーションが終了する。
            e.ExitApplication = False
            ShowMsgError("予期せぬエラーが発生しました。システム管理者へ連絡してください。" & vbCr & vbCrLf & e.Exception.Message)
        End Sub

        'アプリケーション終了時
        Private Sub MyApplication_Shutdown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shutdown
            If frmSupportIdTukekae.dbManager IsNot Nothing Then frmSupportIdTukekae.dbManager.Dispose()
        End Sub

    End Class

End Namespace

