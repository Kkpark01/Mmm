'処理結果クラス
Public Class utlResult

    '属性宣言
    Property Result As ResultType
    Property MsgTitle As String
    Property MsgTxt As String
    Property MsgCd As String
    Property MsgBtnType As MessageBoxButtons
    Property MsgIcon As MessageBoxIcon

    '列挙体宣言
    Public Enum ResultType
        Success     '成功
        Fail        '失敗
        Err         'エラー
    End Enum

    '処理成功を設定する
    Public Sub SetSuccess(Optional ByVal _msgtxt As String = "", Optional ByVal _msgcd As String = "")
        Result = ResultType.Success
        MsgTitle = "情報"
        MsgTxt = _msgtxt
        MsgCd = _msgcd
        MsgBtnType = MessageBoxButtons.OK
        MsgIcon = MessageBoxIcon.Information
    End Sub

    'エラー発生を設定する
    Public Sub SetError(Optional ByVal _msgtxt As String = "", Optional ByVal _msgcd As String = "")
        Result = ResultType.Err
        MsgTitle = "エラー"
        MsgTxt = _msgtxt
        MsgCd = _msgcd
        MsgBtnType = MessageBoxButtons.OK
        MsgIcon = MessageBoxIcon.Error
        SetFormMsg(MsgTxt)
    End Sub

    '処理結果を画面表示する
    Public Sub ShowMsg()
        MessageBox.Show(MsgTxt, MsgTitle, MsgBtnType, MsgIcon)
    End Sub

    '処理結果をログ出力する
    Public Sub WriteLogMsg()
        LogWrite(MsgTxt)
    End Sub

End Class
