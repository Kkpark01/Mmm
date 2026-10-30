Imports System.Data.SqlClient

'データアクセスクラス
Public Class daoSupportIdTukekae

    ''納入先取得引数型
    Private Structure GetNonyuSakiParamsType
        Dim SupportId As String
    End Structure

    'サポートID更新引数型
    Private Structure UpdateSupportIdParamsType
        Dim MotoSupportId As String
        Dim SakiSupportId As String
        Dim MotoNonyuSaki As String
        Dim SakiNonyuSaki As String
    End Structure

    ''納入先取得引数
    Private GetNonyuSakiPrm As GetNonyuSakiParamsType

    'サポートID更新引数
    Private UpdateSupportIdPrm As UpdateSupportIdParamsType

    '納入先取得処理を実行する
    Public Function GetNonyuSaki(ByRef txtSupportId As String, ByRef txtNonyuSaki As String) As String
        Dim res As New utlResult

        '納入先取得引数を設定する
        With GetNonyuSakiPrm
            .SupportId = txtSupportId
        End With

        '納入先取得処理実行
        Dim sql As String = ""
        Dim prm(0) As utlDbManager.ParametersType
        Dim reader As DataTable
        Dim nonyuSaki As String = ""

        Try
            'SQL設定
            sql = ""
            sql &= " SELECT "
            sql &= "     MAA.NONYUSAKI_ID "
            sql &= "   , MAA.NONYUSAKI_NK "
            sql &= " FROM "
            sql &= "     MCM.MCM_MA_NONYUSAKI MAA "
            sql &= " INNER JOIN "
            sql &= "     MCM.MCM_MA_PLANT MAB "
            sql &= " ON "
            sql &= "     MAA.NONYUSAKI_ID = MAB.NONYUSAKI_ID "
            sql &= " WHERE "
            sql &= "     MAB.SUPPORT_ID = ? "
            sql &= " ORDER BY "
            sql &= "     MAA.NONYUSAKI_ID ASC "

            'サポートID設定
            prm(0).Name = "@SUPPORT_ID"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = GetNonyuSakiPrm.SupportId

            'SELECT実行
            reader = frmSupportIdTukekae.dbManager.GetDataTable(sql, prm)
            For Each row As DataRow In reader.Rows
                If nonyuSaki <> "" Then
                    nonyuSaki &= ", "
                End If
                nonyuSaki &= row(0).ToString() & "：" & row(1).ToString()
            Next

        Catch ex As Exception
            res.SetError("納入先取得時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        GetNonyuSaki = nonyuSaki

    End Function

    'サポートID更新処理を実行する
    Public Function UpdateSupportId(ByRef txtMotoSupportId As String, ByRef txtSakiSupportId As String) As utlResult
        Dim res As New utlResult
        Dim committed As Boolean = False

        'サポートID更新引数を設定する
        With UpdateSupportIdPrm
            .MotoSupportId = txtMotoSupportId
            .SakiSupportId = txtSakiSupportId
        End With

        'サポートID更新処理実行
        Try
            SetFormMsg("【開始】処理実行中")

            '処理開始ログ出力
            LogWrite(" START ")

            'トランザクション開始
            frmSupportIdTukekae.dbManager.BeginTransaction()

            'DTS連携フラグ更新処理開始ログ出力
            LogWrite("DTS連携フラグ更新開始")

            'プラントマスタ(MAB)DTS連携フラグ更新
            Using reader As DataTable = Get_MCM_MA_PLANT()
                Dim cnt As Long = 0
                For Each row As DataRow In reader.Rows
                    cnt += UpdateDtsRenkeiFlg_MCM_MA_PLANT(row.Item("PLANT_ID"))
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_MA_PLANT " & "更新対象なし")
                End If
            End Using

            '納入先マスタ(MAA)DTS連携フラグ更新
            Using reader As DataTable = Get_MCM_MA_PLANT()
                Dim cnt As Long = 0
                Dim updatedNonyusakiIds As New HashSet(Of Long)()
                For Each row As DataRow In reader.Rows
                    '同一納入先が複数プラントに紐づく場合も、更新日時は一度だけ進める。
                    Dim nonyusakiId As Long = Convert.ToInt64(row.Item("NONYUSAKI_ID"))
                    If updatedNonyusakiIds.Add(nonyusakiId) Then
                        cnt += UpdateDtsRenkeiFlg_MCM_MA_NONYUSAKI(nonyusakiId)
                    End If
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_MA_NONYUSAKI " & "更新対象なし")
                End If
            End Using

            'DTS連携フラグ更新処理終了ログ出力
            LogWrite("DTS連携フラグ更新終了")

            'サポートID更新処理開始ログ出力
            LogWrite("サポートID更新開始")

            'プラントマスタ(MAB)サポートID更新
            Using reader As DataTable = Get_MCM_MA_PLANT()
                Dim cnt As Long = 0
                For Each row As DataRow In reader.Rows
                    cnt += UpdateSupportId_MCM_MA_PLANT(row.Item("PLANT_ID"))
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_MA_PLANT " & "更新対象なし" & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
                End If
            End Using

            '店舗基本見積(UMA)サポートID更新
            Using reader As DataTable = Get_MCM_UM_KIHON_MITSUMORI()
                Dim cnt As Long = 0
                For Each row As DataRow In reader.Rows
                    cnt += UpdateSupportId_MCM_UM_KIHON_MITSUMORI(row.Item("UM_KIHON_MITSUMORI_ID"), row.Item("PLANT_ID"))
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_UM_KIHON_MITSUMORI " & "更新対象なし" & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
                End If
            End Using

            '取引先見積(TMA)サポートID更新
            Using reader As DataTable = Get_MCM_TM_MITSUMORI()
                Dim cnt As Long = 0
                For Each row As DataRow In reader.Rows
                    cnt += UpdateSupportId_MCM_TM_MITSUMORI(row.Item("TM_IRAI_ID"), row.Item("PLANT_ID"))
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_TM_MITSUMORI " & "更新対象なし" & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
                End If
            End Using

            '取引先契約期間(TKB)サポートID更新
            Using reader As DataTable = Get_MCM_TK_KIKAN()
                Dim cnt As Long = 0
                For Each row As DataRow In reader.Rows
                    cnt += UpdateSupportId_MCM_TK_KIKAN(row.Item("TK_KIKAN_ID"), row.Item("PLANT_ID"))
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_TK_KIKAN " & "更新対象なし" & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
                End If
            End Using

            'ユーザ契約期間(UKB)サポートID更新
            Using reader As DataTable = Get_MCM_UK_KIKAN()
                Dim cnt As Long = 0
                For Each row As DataRow In reader.Rows
                    cnt += UpdateSupportId_MCM_UK_KIKAN(row.Item("UK_KIKAN_ID"), row.Item("PLANT_ID"))
                Next
                If cnt = 0 Then
                    '更新対象なしの場合
                    LogWrite("テーブル：MCM_UK_KIKAN " & "更新対象なし" & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
                End If
            End Using

            'サポートID更新処理終了ログ出力
            LogWrite("サポートID更新終了")

            'コミット
            frmSupportIdTukekae.dbManager.Commit()
            committed = True

            '処理終了ログ出力
            LogWrite("  END  ")

            '成功設定
            res.SetSuccess()

            '画面メッセージ出力
            SetFormMsg("【開始】処理を終了しました")

        Catch ex As Exception
            'システムエラー発生時

            If committed Then
                res.SetError("データの更新は完了しましたが、後処理でエラーが発生しました。再実行前に更新結果を確認してください。" & vbCrLf & ex.Message)
            Else
                Try
                    frmSupportIdTukekae.dbManager.Rollback()
                    res.SetError(ex.Message)
                Catch rollbackEx As Exception
                    res.SetError("処理とロールバックの確認に失敗しました。DBの状態を確認してください。" & vbCrLf & ex.Message & vbCrLf & rollbackEx.Message)
                End Try
            End If

        End Try

        UpdateSupportId = res

    End Function

    'プラントマスタ(MAB)データ取得
    Private Function Get_MCM_MA_PLANT() As DataTable
        Dim sql As String = ""
        Dim prm(0) As utlDbManager.ParametersType
        Dim reader As DataTable

        'SQL設定
        sql = ""
        sql &= " SELECT "
        sql &= "     PLANT_ID "
        sql &= "   , NONYUSAKI_ID " '納入先マスタのDTS連携フラグ更新の為に取得
        sql &= " FROM "
        sql &= "     MCM.MCM_MA_PLANT WITH (UPDLOCK, HOLDLOCK) "
        sql &= " WHERE "
        sql &= "     SUPPORT_ID = ? "
        sql &= " ORDER BY "
        sql &= "     PLANT_ID ASC "

        '引数付け替え元サポートID設定
        prm(0).Name = "@MOTO_SUPPORT_ID"
        prm(0).Type = SqlDbType.NVarChar
        prm(0).Value = UpdateSupportIdPrm.MotoSupportId

        'SELECT実行
        reader = frmSupportIdTukekae.dbManager.GetDataTable(sql, prm)

        Get_MCM_MA_PLANT = reader
    End Function

    '店舗基本見積(UMA)データ取得
    Private Function Get_MCM_UM_KIHON_MITSUMORI() As DataTable
        Dim sql As String = ""
        Dim prm(0) As utlDbManager.ParametersType
        Dim reader As DataTable

        'SQL設定
        sql = ""
        sql &= " SELECT "
        sql &= "     UM_KIHON_MITSUMORI_ID "
        sql &= "   , PLANT_ID "
        sql &= " FROM "
        sql &= "     MCM.MCM_UM_KIHON_MITSUMORI WITH (UPDLOCK, HOLDLOCK) "
        sql &= " WHERE "
        sql &= "     SUPPORT_ID = ? "
        sql &= " ORDER BY "
        sql &= "     UM_KIHON_MITSUMORI_ID ASC "

        '引数付け替え元サポートID設定
        prm(0).Name = "@MOTO_SUPPORT_ID"
        prm(0).Type = SqlDbType.NVarChar
        prm(0).Value = UpdateSupportIdPrm.MotoSupportId

        'SELECT実行
        reader = frmSupportIdTukekae.dbManager.GetDataTable(sql, prm)

        Get_MCM_UM_KIHON_MITSUMORI = reader
    End Function

    '取引先見積(TMA)データ取得
    Private Function Get_MCM_TM_MITSUMORI() As DataTable
        Dim sql As String = ""
        Dim prm(0) As utlDbManager.ParametersType
        Dim reader As DataTable

        'SQL設定
        sql = ""
        sql &= " SELECT "
        sql &= "     TM_IRAI_ID "
        sql &= "   , PLANT_ID "
        sql &= " FROM "
        sql &= "     MCM.MCM_TM_MITSUMORI WITH (UPDLOCK, HOLDLOCK) "
        sql &= " WHERE "
        sql &= "     SUPPORT_ID = ? "
        sql &= " ORDER BY "
        sql &= "     TM_IRAI_ID ASC "

        '引数付け替え元サポートID設定
        prm(0).Name = "@MOTO_SUPPORT_ID"
        prm(0).Type = SqlDbType.NVarChar
        prm(0).Value = UpdateSupportIdPrm.MotoSupportId

        'SELECT実行
        reader = frmSupportIdTukekae.dbManager.GetDataTable(sql, prm)

        Get_MCM_TM_MITSUMORI = reader
    End Function

    '取引先契約期間(TKB)データ取得
    Private Function Get_MCM_TK_KIKAN() As DataTable
        Dim sql As String = ""
        Dim prm(0) As utlDbManager.ParametersType
        Dim reader As DataTable

        'SQL設定
        sql = ""
        sql &= " SELECT "
        sql &= "     TK_KIKAN_ID "
        sql &= "   , PLANT_ID "
        sql &= " FROM "
        sql &= "     MCM.MCM_TK_KIKAN WITH (UPDLOCK, HOLDLOCK) "
        sql &= " WHERE "
        sql &= "     SUPPORT_ID = ? "
        sql &= " ORDER BY "
        sql &= "     TK_KIKAN_ID ASC "

        '引数付け替え元サポートID設定
        prm(0).Name = "@MOTO_SUPPORT_ID"
        prm(0).Type = SqlDbType.NVarChar
        prm(0).Value = UpdateSupportIdPrm.MotoSupportId

        'SELECT実行
        reader = frmSupportIdTukekae.dbManager.GetDataTable(sql, prm)

        Get_MCM_TK_KIKAN = reader
    End Function

    'ユーザ契約期間(UKB)データ取得
    Private Function Get_MCM_UK_KIKAN() As DataTable
        Dim sql As String = ""
        Dim prm(0) As utlDbManager.ParametersType
        Dim reader As DataTable

        'SQL設定
        sql = ""
        sql &= " SELECT "
        sql &= "     UK_KIKAN_ID "
        sql &= "   , PLANT_ID "
        sql &= " FROM "
        sql &= "     MCM.MCM_UK_KIKAN WITH (UPDLOCK, HOLDLOCK) "
        sql &= " WHERE "
        sql &= "     SUPPORT_ID = ? "
        sql &= " ORDER BY "
        sql &= "     UK_KIKAN_ID ASC "

        '引数付け替え元サポートID設定
        prm(0).Name = "@MOTO_SUPPORT_ID"
        prm(0).Type = SqlDbType.NVarChar
        prm(0).Value = UpdateSupportIdPrm.MotoSupportId

        'SELECT実行
        reader = frmSupportIdTukekae.dbManager.GetDataTable(sql, prm)

        Get_MCM_UK_KIKAN = reader
    End Function

    'プラントマスタ(MAB)DTS連携フラグ更新
    Private Function UpdateDtsRenkeiFlg_MCM_MA_PLANT(ByVal key As Long) As Integer
        Dim sql As String = ""
        Dim prm(1) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_MA_PLANT "
            sql &= " SET "
            sql &= "     DTSRENKEI_FLG = 0 "  'DTS連携フラグを「連携なし」にする
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     PLANT_ID = ? "

            '更新キー設定
            prm(1).Name = "@KEY"
            prm(1).Type = SqlDbType.BigInt
            prm(1).Value = key

            prm(0).Name = "@UPDATE_BY"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_MA_PLANT PLANT_ID=" & key & " DTS連携フラグ更新")
            End If

        Catch ex As Exception
            Throw New ApplicationException("プラントマスタ(MAB)DTS連携フラグ更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateDtsRenkeiFlg_MCM_MA_PLANT = count

    End Function

    '納入先マスタ(MAA)DTS連携フラグ更新
    Private Function UpdateDtsRenkeiFlg_MCM_MA_NONYUSAKI(ByVal key As Long) As Integer
        Dim sql As String = ""
        Dim prm(1) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_MA_NONYUSAKI "
            sql &= " SET "
            sql &= "     DTSRENKEI_FLG = 0 "  'DTS連携フラグを「連携なし」にする
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     NONYUSAKI_ID = ? "

            '更新キー設定
            prm(1).Name = "@KEY"
            prm(1).Type = SqlDbType.BigInt
            prm(1).Value = key

            prm(0).Name = "@UPDATE_BY"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_MA_NONYUSAKI" & " NONYUSAKI_ID=" & key & " DTS連携フラグ更新")
            End If

        Catch ex As Exception
            Throw New ApplicationException("納入先マスタ(MAA)DTS連携フラグ更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateDtsRenkeiFlg_MCM_MA_NONYUSAKI = count

    End Function

    'プラントマスタ(MAB)サポートID更新
    Private Function UpdateSupportId_MCM_MA_PLANT(ByVal key As Long) As Integer
        Dim sql As String = ""
        Dim prm(2) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_MA_PLANT "
            sql &= " SET "
            sql &= "     SUPPORT_ID = ? "
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     PLANT_ID = ? "

            '引数付け替え先サポートID設定
            prm(0).Name = "@SAKI_SUPPORT_ID"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = UpdateSupportIdPrm.SakiSupportId

            '更新キー設定
            prm(2).Name = "@KEY"
            prm(2).Type = SqlDbType.BigInt
            prm(2).Value = key

            prm(1).Name = "@UPDATE_BY"
            prm(1).Type = SqlDbType.NVarChar
            prm(1).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_MA_PLANT PLANT_ID=" & key & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
            End If

        Catch ex As Exception
            Throw New ApplicationException("プラントマスタ(MAB)サポートID更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateSupportId_MCM_MA_PLANT = count

    End Function

    '店舗基本見積(UMA)サポートID更新
    Private Function UpdateSupportId_MCM_UM_KIHON_MITSUMORI(ByVal key As Long, ByVal plantid As String) As Long
        Dim sql As String = ""
        Dim prm(2) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_UM_KIHON_MITSUMORI "
            sql &= " SET "
            sql &= "     SUPPORT_ID = ? "
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     UM_KIHON_MITSUMORI_ID = ? "

            '引数付け替え先サポートID設定
            prm(0).Name = "@SAKI_SUPPORT_ID"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = UpdateSupportIdPrm.SakiSupportId

            '更新キー設定
            prm(2).Name = "@KEY"
            prm(2).Type = SqlDbType.BigInt
            prm(2).Value = key

            prm(1).Name = "@UPDATE_BY"
            prm(1).Type = SqlDbType.NVarChar
            prm(1).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_UM_KIHON_MITSUMORI" & " PLANT_ID=" & plantid & " UM_KIHON_MITSUMORI_ID=" & key & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
            End If

        Catch ex As Exception
            Throw New ApplicationException("店舗基本見積(UMA)サポートID更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateSupportId_MCM_UM_KIHON_MITSUMORI = count

    End Function

    '取引先見積(TMA)サポートID更新
    Private Function UpdateSupportId_MCM_TM_MITSUMORI(ByVal key As Long, ByVal plantid As String) As Long
        Dim sql As String = ""
        Dim prm(2) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_TM_MITSUMORI "
            sql &= " SET "
            sql &= "     SUPPORT_ID = ? "
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     TM_IRAI_ID = ? "

            '引数付け替え先サポートID設定
            prm(0).Name = "@SAKI_SUPPORT_ID"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = UpdateSupportIdPrm.SakiSupportId

            '更新キー設定
            prm(2).Name = "@KEY"
            prm(2).Type = SqlDbType.BigInt
            prm(2).Value = key

            prm(1).Name = "@UPDATE_BY"
            prm(1).Type = SqlDbType.NVarChar
            prm(1).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_TM_MITSUMORI" & " PLANT_ID=" & plantid & " TM_IRAI_ID=" & key & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
            End If

        Catch ex As Exception
            Throw New ApplicationException("取引先見積(TMA)サポートID更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateSupportId_MCM_TM_MITSUMORI = count
    End Function

    '取引先契約期間(TKB)サポートID更新
    Private Function UpdateSupportId_MCM_TK_KIKAN(ByVal key As Long, ByVal plantid As String) As Long
        Dim sql As String = ""
        Dim prm(2) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_TK_KIKAN "
            sql &= " SET "
            sql &= "     SUPPORT_ID = ? "
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     TK_KIKAN_ID = ? "

            '引数付け替え先サポートID設定
            prm(0).Name = "@SAKI_SUPPORT_ID"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = UpdateSupportIdPrm.SakiSupportId

            '更新キー設定
            prm(2).Name = "@KEY"
            prm(2).Type = SqlDbType.BigInt
            prm(2).Value = key

            prm(1).Name = "@UPDATE_BY"
            prm(1).Type = SqlDbType.NVarChar
            prm(1).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_TK_KIKAN" & " PLANT_ID=" & plantid & " TK_KIKAN_ID=" & key & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
            End If

        Catch ex As Exception
            Throw New ApplicationException("取引先契約期間(TKB)サポートID更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateSupportId_MCM_TK_KIKAN = count
    End Function

    'ユーザ契約期間(UKB)サポートID更新
    Private Function UpdateSupportId_MCM_UK_KIKAN(ByVal key As Long, ByVal plantid As String) As Long
        Dim sql As String = ""
        Dim prm(2) As utlDbManager.ParametersType
        Dim count As Integer = 0

        Try
            'SQL設定
            sql = ""
            sql &= " UPDATE "
            sql &= "     MCM.MCM_UK_KIKAN "
            sql &= " SET "
            sql &= "     SUPPORT_ID = ? "
            'Webの楽観的排他が更新を検知できるよう、日時を必ず進める。
            sql &= "   , LASTUPDATE_DT = CASE WHEN LASTUPDATE_DT >= CONVERT(datetime2(0), SYSDATETIME()) THEN DATEADD(SECOND, 1, LASTUPDATE_DT) ELSE CONVERT(datetime2(0), SYSDATETIME()) END "
            sql &= "   , LASTUPDATE_BY = ? "
            sql &= " WHERE "
            sql &= "     UK_KIKAN_ID = ? "

            '引数付け替え先サポートID設定
            prm(0).Name = "@SAKI_SUPPORT_ID"
            prm(0).Type = SqlDbType.NVarChar
            prm(0).Value = UpdateSupportIdPrm.SakiSupportId

            '更新キー設定
            prm(2).Name = "@KEY"
            prm(2).Type = SqlDbType.BigInt
            prm(2).Value = key

            prm(1).Name = "@UPDATE_BY"
            prm(1).Type = SqlDbType.NVarChar
            prm(1).Value = frmSupportIdTukekae.dbManager.UpdateBy

            'UPDATE実行
            count = frmSupportIdTukekae.dbManager.ExcuteSql(sql, prm)
            If count > 0 Then
                '処理成功時
                LogWrite("テーブル：MCM_UK_KIKAN" & " PLANT_ID=" & plantid & " UK_KIKAN_ID=" & key & " 旧SUPPORT_ID=" & UpdateSupportIdPrm.MotoSupportId & " 新SUPPORT_ID=" & UpdateSupportIdPrm.SakiSupportId & "")
            End If

        Catch ex As Exception
            LogWrite("ユーザ契約期間(UKB)サポートID更新時にエラー発生。" & vbCrLf & ex.Message)
        End Try

        UpdateSupportId_MCM_UK_KIKAN = count
    End Function

End Class
