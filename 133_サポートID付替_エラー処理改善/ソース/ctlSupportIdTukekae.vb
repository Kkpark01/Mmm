'コントローラーモジュール
Public Module ctlSupportIdTukekae

    '付け替え処理実行
    Public Sub ExcuteTukekae(ByRef txtMotoSupportId As TextBox, ByRef txtSakiSupportId As TextBox, ByRef txtMotoNonyuSaki As TextBox, ByRef txtSakiNonyuSaki As TextBox)
        Dim res As New utlResult
        Dim dao As New daoSupportIdTukekae

        'ログファイルチェック
        res = CheckLogFile()
        If res.Result = utlResult.ResultType.Err Then
            res.MsgTxt = "ログファイル確認が失敗した為、処理を中止します。" & vbCrLf & vbCrLf & "【エラー内容】" & vbCrLf & res.MsgTxt
            SetFormMsg(res.MsgTxt)
            Exit Sub
        End If

        '付け替え元サポートIDチェック
        res = CheckMotoSupportId(txtMotoSupportId, txtMotoNonyuSaki)
        If res.Result = utlResult.ResultType.Err Then
            Exit Sub
        End If

        '付け替え先サポートIDチェック
        res = CheckSakiSupportId(txtMotoSupportId, txtSakiSupportId, txtSakiNonyuSaki)
        If res.Result = utlResult.ResultType.Err Then
            Exit Sub
        End If

        'サポートID更新
        res = dao.UpdateSupportId(txtMotoSupportId.Text, txtSakiSupportId.Text)
        If res.Result = utlResult.ResultType.Err Then
            Exit Sub
        End If

    End Sub

    '納入先取得処理実行
    Public Function ExcuteGetNonyuSaki(ByRef txtSupportId As TextBox, ByRef txtNonyuSaki As TextBox) As String
        Dim res As New utlResult
        Dim dao As New daoSupportIdTukekae
        Dim nonyuSaki As String

        '納入先取得
        txtNonyuSaki.Text = ""
        nonyuSaki = dao.GetNonyuSaki(txtSupportId.Text, txtNonyuSaki.Text)
        txtNonyuSaki.Text = nonyuSaki
        ExcuteGetNonyuSaki = nonyuSaki

    End Function

    '付け替え元サポートIDチェック
    Public Function CheckMotoSupportId(ByRef txtMotoSupportId As TextBox, ByRef txtMotoNonyuSaki As TextBox) As utlResult
        Dim res As New utlResult

        '付け替え元サポートID必須入力チェック
        If Not CheckInputRequired(txtMotoSupportId.Text) Then
            txtMotoSupportId.Select()
            res.SetError("付け替え元サポートIDを入力してください")
            CheckMotoSupportId = res
            Exit Function
        End If

        '付け替え元サポートIDに対する納入先存在チェック
        Try
            If ExcuteGetNonyuSaki(txtMotoSupportId, txtMotoNonyuSaki) = "" Then
                txtMotoSupportId.Select()
                res.SetError("データとして存在する付け替え元サポートIDを入力してください ")
                CheckMotoSupportId = res
                Exit Function
            End If
        Catch ex As Exception
            res.SetError(ex.Message)
            CheckMotoSupportId = res
            Exit Function
        End Try

        '処理成功時
        res.SetSuccess()
        CheckMotoSupportId = res
    End Function

    '付け替え先サポートIDチェック
    Public Function CheckSakiSupportId(ByRef txtMotoSupportId As TextBox, ByRef txtSakiSupportId As TextBox, ByRef txtSakiNonyuSaki As TextBox) As utlResult
        Dim res As New utlResult

        '付け替え元サポートID必須入力チェック
        If Not CheckInputRequired(txtMotoSupportId.Text) Then
            txtMotoSupportId.Select()
            res.SetError("付け替え元サポートIDを入力してください")
            CheckSakiSupportId = res
            Exit Function
        End If

        '付け替え先サポートID必須入力チェック
        If Not CheckInputRequired(txtSakiSupportId.Text) Then
            txtSakiSupportId.Select()
            res.SetError("付け替え先サポートIDを入力してください")
            CheckSakiSupportId = res
            Exit Function
        End If

        '付け替え元サポートIDと付け替え先サポートIDの同一チェック
        If txtMotoSupportId.Text = txtSakiSupportId.Text Then
            txtSakiSupportId.Select()
            res.SetError("付け替え先サポートIDが付け替え元と同一です")
            CheckSakiSupportId = res
            Exit Function
        End If

        '付け替え先サポートIDに対する納入先取得
        Try
            ExcuteGetNonyuSaki(txtSakiSupportId, txtSakiNonyuSaki)
        Catch ex As Exception
            res.SetError(ex.Message)
            CheckSakiSupportId = res
            Exit Function
        End Try

        '処理成功時
        res.SetSuccess()
        CheckSakiSupportId = res
    End Function

End Module

