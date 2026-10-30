Attribute VB_Name = "Get_MCM_DATA"
Option Explicit
Dim exportWS As Worksheet                       '// 「取込データ」シート
Dim PilotWS As Worksheet                        '// 「操作シート」
Dim rngDest As String                           '// csvファイルフルパス
Dim Dest As String                              '// csvファイル名格納
Dim rng As Range                                '// csv取込開始場所指定用
'*******************************
Sub Get_MCM()
AdoGet_MCM "ALL"
End Sub
'*******************************
Sub exportdata_tran()
'csv取込
'*******************************
Dim nm As Name                              '//csv接続情報削除用
With exportWS.QueryTables.Add(Connection:="TEXT;" & rngDest, Destination:=rng)
.Name = "export"
.FieldNames = True
.RowNumbers = False
.FillAdjacentFormulas = False
.PreserveFormatting = True
.RefreshOnFileOpen = False
.RefreshStyle = xlInsertDeleteCells
.SavePassword = False
.SaveData = True
.AdjustColumnWidth = True
.RefreshPeriod = 0
.TextFilePromptOnRefresh = False
.TextFilePlatform = 932
.TextFileStartRow = 1
.TextFileParseType = xlDelimited
.TextFileTextQualifier = xlTextQualifierDoubleQuote
.TextFileConsecutiveDelimiter = False
.TextFileTabDelimiter = False
.TextFileSemicolonDelimiter = False
.TextFileCommaDelimiter = True
.TextFileSpaceDelimiter = False
.TextFileColumnDataTypes = Array(1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 1, 1, 1, _
1, 1, 1, 1, 1, 1, 1, 2, 2, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1, _
1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
'// 文字列取込対象（サポートID,シリアル番号,取引先契約番号,HW製番,SW製番,機器ID）
.TextFileTrailingMinusNumbers = True
.Refresh BackgroundQuery:=False
End With
'CSV接続情報削除
Do While ActiveWorkbook.Connections.Count > 0
ActiveWorkbook.Connections.item(ActiveWorkbook.Connections.Count).Delete
Loop
For Each nm In ActiveWorkbook.Names
On Error Resume Next
nm.Delete
Next nm
End Sub
'*******************************
Sub Data_Check()
'csv取込結果確認
'データ内に改行コードが含まれている場合、アラート表示＆処理終了
'*******************************
Dim CtrlUpRow As Long                       '// リスト最終行取得用
Dim CtrlDownRow As Long                     '// リスト中断行取得用
Dim ER_Col As Long                          '// エラーセル取得用
Dim LastCol As Long                         '// リスト最終列取得用
Dim chRow As Long                           '// 列右ズレデータ検出用
Dim i As Long                               '// ループカウンタ
LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
'//■■↓　セル内改行検知用　↓■■
'// 「Ctrl+↑」で最下段よりリスト最終行取得
CtrlUpRow = Cells(Rows.Count, LastCol).End(xlUp).Row
If CtrlUpRow = 1 Then Exit Sub              '// 取込結果が項目のみ場合、check終了
'// 「Ctrl+↓」で1行目からリスト最終行取得
CtrlDownRow = Cells(1, LastCol).End(xlDown).Row
If CtrlUpRow <> CtrlDownRow Then
ER_Col = Cells(CtrlDownRow + 1, Columns.Count).End(xlToLeft).Column
With Cells(CtrlDownRow + 1, ER_Col)     '// エラーセルにフォント設定
.Select
.Interior.Color = RGB(200, 200, 200)
.Font.Color = RGB(255, 0, 0)
.Font.Bold = True
.Font.Italic = True
End With
Application.ScreenUpdating = True
Cells(CtrlDownRow + 1, ER_Col).Activate
MsgBox "【サポートID:" & Cells(CtrlDownRow + 1, 4) & "】に不正なデータが含まれてます。" _
& vbCrLf & vbCrLf & Dest & "を確認して、ＭＣＭを修正してください。"
End
End If
'//■■↑　セル内改行検知用　↑■■
'//■■↓　セル内「”」&「,」検知用（セル右ズレ検知）　↓■■
For i = LastCol + 1 To LastCol + 30
CtrlUpRow = Cells(Rows.Count, i).End(xlUp).Row
If CtrlUpRow <> 1 Then
ER_Col = Cells(CtrlUpRow, Columns.Count).End(xlToLeft).Column
With Cells(CtrlUpRow, ER_Col)     '// エラーセルにフォント設定
.Select
.Interior.Color = RGB(200, 200, 200)
.Font.Color = RGB(255, 0, 0)
.Font.Bold = True
.Font.Italic = True
End With
Application.ScreenUpdating = True
Cells(CtrlUpRow, ER_Col).Activate
MsgBox "【サポートID:" & Cells(CtrlUpRow, 4) & "】に不正なデータが含まれてます。" _
& vbCrLf & vbCrLf & Dest & "を確認して、ＭＣＭを修正してください。"
End
End If
Next i
'//■■↑　セル内「”」&「,」検知用（セル右ズレ検知）　↑■■
End Sub
'*******************************
Sub Duplicate_item_check()
'重複した明細をチェック
'リスト上段より、countif(相対参照:絶対参照,機器管理ID)を行う
'*******************************
Dim item()                                  '// 機器管理ID配列
Dim LastRow As Long                         '// データ最終行格納用
Dim i As Long                               '// 検索値用ループカウンタ
Dim buf As Variant                          '// 検索機器管理ID格納用
Dim Si As Long                              '// 参照用ループカウンタ
Dim cnt As Long                             '// 重複データカウンタ
Dim item_cnt()                              '// 重複カウンタ配列
Dim fastKeys As Boolean, seen As Object
LastRow = Cells(Rows.Count, 2).End(xlUp).Row
ReDim item_cnt(1 To LastRow, 1 To 1)
'// 機器管理ID格納
item = Range(Cells(1, 1), Cells(LastRow, 2))
' SQLの機器管理IDは文字列。文字列以外を含む場合は従来の比較を維持する。
fastKeys = True
For i = 2 To LastRow
If VarType(item(i, 2)) <> vbString Then fastKeys = False: Exit For
Next i
If fastKeys Then
Set seen = CreateObject("Scripting.Dictionary")
seen.CompareMode = vbBinaryCompare
' 後方から数え、最後の行に1、それより前の同じIDに2を設定する。
For i = LastRow To 2 Step -1
If (LastRow - i + 1) Mod 200 = 0 Or i = LastRow Or i = 2 Then
Application.StatusBar = "重複データ確認中...(" & LastRow - i + 1 & "/" & LastRow & "件)"
DoEvents
End If
buf = item(i, 2)
If seen.Exists(buf) Then
item_cnt(i, 1) = 2
Else
seen.Add buf, True
item_cnt(i, 1) = 1
End If
Next i
Else
' CSV等で型が異なる場合は変更前と同じ比較を行う。
For i = 2 To LastRow
If i Mod 200 = 0 Or i = 2 Or i = LastRow Then
Application.StatusBar = "重複データ確認中...(" & i & "/" & LastRow & "件)"
DoEvents
End If
buf = item(i, 2)
For Si = i To LastRow
If buf = item(Si, 2) Then cnt = cnt + 1
If cnt = 2 Then Exit For
Next Si
item_cnt(i, 1) = cnt
cnt = 0
Next i
End If
'// 重複カウンタ書き出し
Range(Cells(1, 1), Cells(LastRow, 1)) = item_cnt
Application.StatusBar = False
End Sub
'*******************************
Sub col_Del()
'重複データ行削除
'*******************************
Dim LastRow As Long                         '// データ最終行格納用
Dim i As Long                               '// ループカウンタ
'// 重複カウンタ結果で昇順ソート
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(1, 1), Cells(LastRow, 114)).Sort _
key1:=Range("A1"), order1:=xlAscending, Header:=xlYes
'// リスト最終行が重複データか確認
If Cells(LastRow, 1) <> 1 Then
'// リスト最終行より、重複カウンタ「1」行を検索
For i = LastRow To 1 Step -1
If Cells(i, 1) = 1 Then Exit For
Next i
'// 上記結果がリスト最終行より上段であることを確認
If i < LastRow Then
i = i + 1
Rows(i & ":" & LastRow).Delete
End If
End If
'// ①「納入先コード」　②「サポートID」　③「システム名」順でソート
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(1, 1), Cells(LastRow, 114)).Sort _
key1:=Range("E1"), order1:=xlAscending, _
key2:=Range("D1"), order2:=xlAscending, _
key3:=Range("J1"), order3:=xlAscending, Header:=xlYes
End Sub
'*******************************
Sub VLOOKUP_TENPO()
'店舗リストより担当店舗当て込み
'*******************************
Dim tempWS As Worksheet                     '//店舗リストシート
Dim LastRow As Long                         '//取込データ最終行取得用
Dim tempRow As Long                         '//店舗リスト最終行取得用
Dim i As Long                               '//ループカウンタ
Dim item()                                  '//サポートID、店舗名セット用配列
Set tempWS = Worksheets("店舗リスト")
LastRow = exportWS.Cells(Rows.Count, 1).End(xlUp).Row
item = Range(exportWS.Cells(2, 3), exportWS.Cells(LastRow, 4))  '//店舗列、サポートIDを配列格納
tempRow = tempWS.Cells(Rows.Count, 1).End(xlUp).Row
On Error GoTo List_add
For i = 1 To LastRow - 1
Application.StatusBar = "担当店舗引き当て中...(" & i & "/" & LastRow & "件)"
DoEvents
item(i, 1) = WorksheetFunction.VLookup(item(i, 2), Range(tempWS.Cells(1, 1), tempWS.Cells(tempRow, 2)), 2, False)
Re:     '//新規サポートID書き出し後のリターンラベル
Next i
Application.StatusBar = False
exportWS.Range(Cells(2, 3), Cells(LastRow, 3)) = item           '//引き当て結果書き出し
'// ①「店舗名」　②「納入先コード」　③「サポートID」順でソート
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(1, 1), Cells(LastRow, 114)).Sort _
key1:=Range("C1"), order1:=xlAscending, _
key2:=Range("E1"), order2:=xlAscending, _
key3:=Range("D1"), order3:=xlAscending, Header:=xlYes
Exit Sub
'//新規サポートIDを店舗リストに書き出し
List_add:
tempWS.Cells(tempRow + 1, 1).Value = item(i, 2)
tempRow = tempRow + 1
Resume Re
End Sub
'*******************************
Sub comment_create()
'備考欄作成、明細単位の終息情報、契約時間情報書き出し
'*******************************
Dim LastRow As Long                         '//取込データ最終行取得用
Dim item()                                  '//備考内容セット用配列
Dim comment_item()                          '//各備考情報セット用配列
Dim LDay As Date                            '//終息日、延長終息日一時保管用
Dim sep As Long                             '//保守契約時間情報取得用
Dim i As Long                               '//ループカウンタ
exportWS.Activate
Range(Columns(49), Columns(52)).Insert      '//各備考情報列追加
Cells(1, 49).Value = "終息"
Cells(1, 50).Value = "延長終息"
Cells(1, 51).Value = "契約時間"
Cells(1, 52).Value = "その他"
Cells(1, 53).Value = "備考"
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
item = Range(Cells(2, 59), Cells(LastRow, 59))  '//備考内容セット
ReDim comment_item(1 To LastRow - 1, 1 To 3)
For i = 1 To LastRow - 1
If i Mod 200 = 0 Or i = 1 Or i = LastRow - 1 Then
Application.StatusBar = "備考編集中...(" & i & "/" & LastRow & "件)"
DoEvents
End If
If item(i, 1) <> Empty Then
If InStr(item(i, 1), "★終息") <> 0 Then
If InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") <> 0 Then
If IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") + 1, 10)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") + 1, 10)
ElseIf IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") + 1, 9)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") + 1, 9)
ElseIf IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") + 1, 8)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "」") + 1, 8)
End If
Else
If IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "≫") + 1, 10)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "≫") + 1, 10)
ElseIf IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "≫") + 1, 9)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "≫") + 1, 9)
ElseIf IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "≫") + 1, 8)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息") + 2, item(i, 1), "≫") + 1, 8)
End If
End If
If LDay <> 0 Then _
comment_item(i, 1) = LDay
LDay = Empty
End If
If InStr(item(i, 1), "★終息 →") <> 0 Then
If IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息 →") + 2, item(i, 1), "→") + 1, 10)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息 →") + 2, item(i, 1), "→") + 1, 10)
ElseIf IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息 →") + 2, item(i, 1), "→") + 1, 9)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息 →") + 2, item(i, 1), "→") + 1, 9)
ElseIf IsDate(Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息 →") + 2, item(i, 1), "→") + 1, 8)) Then
LDay = Mid(item(i, 1), InStr(InStr(item(i, 1), "★終息 →") + 2, item(i, 1), "→") + 1, 8)
End If
If LDay <> 0 Then _
comment_item(i, 2) = LDay
LDay = Empty
End If
If InStr(item(i, 1), "★保守契約時間") <> 0 Then
If InStr(InStr(item(i, 1), "★保守契約時間") + 7, item(i, 1), "/") <> 0 Then
sep = InStr(InStr(item(i, 1), "★保守契約時間") + 7, item(i, 1), "/") - InStr(item(i, 1), "★保守契約時間") - 7
comment_item(i, 3) = Mid(item(i, 1), InStr(item(i, 1), "★保守契約時間") + 8, sep - 2)
Else
comment_item(i, 3) = Mid(item(i, 1), InStr(item(i, 1), "★保守契約時間") + 8, 100)
End If
sep = Empty
End If
End If
Next i
Range(Cells(2, 49), Cells(LastRow, 51)).Value = comment_item
Application.StatusBar = False
End Sub
'*******************************
Sub Machine_Class()
'分類記号引き当て関数入力
'*******************************
Dim mainWS As Worksheet                     '//関数セット先シート用
Dim tempWS As Worksheet                     '//参照先リストシート用
Dim tempRow As Long                         '//参照先リスト範囲下段取得用
Dim LastRow As Long                         '//入力対象セル範囲下段取得用
Dim WSname As String                        '//参照先リストシート名入力用
WSname = "機器分類表"                       '//参照先シート名をセット
Set mainWS = Worksheets("取込データ")       '//関数セット先シート名をセット
Set tempWS = Worksheets(WSname)
LastRow = mainWS.Cells(Rows.Count, 2).End(xlUp).Row
tempRow = tempWS.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(2, 1), Cells(LastRow, 1)).FormulaR1C1 = "=if(index(" & WSname & "!R2C1:R" & tempRow & "C1,match(RC54," & WSname & "!R2C1:R" & tempRow & "C1,1),1)=RC54,VLOOKUP(RC54," & WSname & "!R2C1:R" & tempRow & "C5,5,true),NA())"
'■関数入力範囲　　　　　　　　　　　　　　　　　　　　　　　■検索範囲（R2C2=セルB2)　　　　　　↑■検索値(RC3=同一行のC列)　　　　　　　　　　　↑■検索値(RC3=同一行のC列)　　　↑■VLOOK範囲B2:F最終行　(C6=6列目Fを指定)
'IF(INDEX(機器分類表!$B$2:$B$9326,MATCH($C2,機器分類表!$B$2:$B$9326,1),1)=$C2,VLOOKUP($C2,機器分類表!$B$2:$F$9326,5,TRUE),NA())
'これを入力している
End Sub
'*******************************
Sub some_machine_class()
'個別機器分類引き当て
'*******************************
Dim rng As Range                            '//対象管理IDセル位置取得用
Dim LastRow As Long                         '//個別機器分類表リスト範囲下段取得用
Dim ListRow As Long                         '//取込データ最終行取得用
Dim i As Long                               '//ループカウンタ
exportWS.Activate
ListRow = exportWS.Cells(Rows.Count, 2).End(xlUp).Row
With Worksheets("個別機器分類表")
LastRow = .Cells(Rows.Count, 1).End(xlUp).Row
If LastRow = 1 Then Exit Sub            '//個別調整対象の有無確認
For i = LastRow To 2 Step -1
'//調整対象の機器管理IDセル位置取得
Application.StatusBar = "機器分類個別調整中...(" & LastRow - i + 1 & "/" & LastRow & "件)"
Set rng = Range(exportWS.Cells(2, 2), exportWS.Cells(ListRow, 2)).Find(what:=.Cells(i, 1), LookIn:=xlValues, Lookat:=xlWhole)
If Not rng Is Nothing Then              '//対象セル判定
rng.Offset(0, -1) = .Cells(i, 2)    '//対象機器管理ID検出OK：分類を編集
Else
.Rows(i).Delete                     '//対象機器管理ID未検出：撤去等管理対象外として、個別機器分類表より削除
End If
Next i
End With
Application.StatusBar = False
PilotWS.Activate
End Sub
'*******************************
Sub Machine_Class_sort()
'機器分類再引き当て
'*******************************
Dim LastRow As Long                         '//機器分類表リスト範囲取得用
Dim classWS As Worksheet                    '//機器分類表シート
Set PilotWS = Worksheets("操作シート")
Set exportWS = Worksheets("取込データ")
Set classWS = Worksheets("機器分類表")
Application.ScreenUpdating = False
classWS.Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'// 機器分類表を機器ID順で再ソート
Range(Cells(1, 1), Cells(LastRow, 6)).Sort _
key1:=Range("A1"), order1:=xlAscending, Header:=xlYes
exportWS.Activate
Call Machine_Class                          '//関数入力処理へ
Call some_machine_class                     '//個別機器分類引き当て
Application.ScreenUpdating = True
PilotWS.Activate
MsgBox "機器分類再引き当て完了"
End Sub
'*******************************
Sub tenpo()
'店舗再引き当て
'*******************************
Dim LastRow As Long                         '//店舗リスト範囲取得用
Set PilotWS = Worksheets("操作シート")
Set exportWS = Worksheets("取込データ")
Application.ScreenUpdating = False
Worksheets("店舗リスト").Activate
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'//店舗リストをサポートID順でソート
Range(Cells(1, 1), Cells(LastRow, 2)).Sort _
key1:=Range("A1"), order1:=xlAscending, Header:=xlYes
exportWS.Activate
Call VLOOKUP_TENPO                          '//担当店舗引き当て処理へ
PilotWS.Activate
Application.ScreenUpdating = True
MsgBox "店舗再引き当て完了"
End Sub
'*******************************
Sub Tran_MCM()
'各csv作成済み状態で再開
'*******************************
Dim sh  As New IWshRuntimeLibrary.WshShell  '// WshShellクラスオブジェクト
Dim ex  As WshExec                          '// Execメソッド戻り値
Dim sArCmd(6)                               '// 実行コマンド配列
Dim sCmd                                    '// 実行コマンド
Dim i                                       '// ループカウンタ
Dim LastRow As Long                         '// データ最終行格納用
Dim st                                      '// 処理開始時間格納用
Dim New_SupportID As Long                   '// 新規サポートID件数表示用
McmEndData.BeginCsvMode
st = Now
Application.ScreenUpdating = False
Set exportWS = Worksheets("取込データ")
Set PilotWS = Worksheets("操作シート")
'// 前回取込データ削除
Application.DisplayAlerts = False
exportWS.Delete
Worksheets.Add after:=Worksheets(Worksheets.Count)
ActiveSheet.Name = "取込データ"
Set exportWS = Worksheets("取込データ")
ActiveWindow.Zoom = 80
Application.DisplayAlerts = True
PilotWS.Activate
ChDir ThisWorkbook.Path
'// 未契約csvデータ取込
exportWS.Activate
Set rng = Cells(1, 1)
Dest = "32機器一覧「未契約」.csv"
rngDest = ThisWorkbook.Path & "\32機器一覧「未契約」.csv"
Call exportdata_tran                        '//csv取込
Call Data_Check                             '//取込結果確認
'// 解約csvデータ取込
LastRow = Cells(Rows.Count, 2).End(xlUp).Row
Set rng = Cells(LastRow + 1, 1)
Dest = "35機器一覧「解約（途中解約）」.csv"
rngDest = ThisWorkbook.Path & "\35機器一覧「解約（途中解約）」.csv"
Call exportdata_tran                        '//csv取込
'// 解約csvデータ項目行削除
Rows(LastRow + 1).Delete
Call Data_Check                             '//取込結果確認
'// 契約csvデータ取込
LastRow = Cells(Rows.Count, 2).End(xlUp).Row
Set rng = Cells(LastRow + 1, 1)
Dest = "31機器一覧「契約」.csv"
rngDest = ThisWorkbook.Path & "\31機器一覧「契約」.csv"
Call exportdata_tran                        '//csv取込
'// 契約csvデータ項目行削除
Rows(LastRow + 1).Delete
Call Data_Check                             '//取込結果確認
Call Duplicate_item_check                   '//重複データ確認
Call col_Del                                '//重複データ削除
'    Call VLOOKUP_TENPO                          '//担当店舗引き当て
Call comment_create                         '//備考コメント引き当て
'    Call Machine_Class                          '//機器分類引き当て
'    Call some_machine_class                     '//個別機器分類引き当て
Call Edit_seiban                            '//保守製番引き当て
Worksheets("店舗リスト").Activate
'    LastRow = Cells(Rows.Count, 1).End(xlUp).Row                    '//サポートID件数取得
'    New_SupportID = LastRow - Cells(Rows.Count, 2).End(xlUp).Row    '//新規サポートID件数算出
Application.ScreenUpdating = True
If New_SupportID = 0 Then
PilotWS.Activate
MsgBox "MCMデータ取得完了" & vbCrLf & vbCrLf & "新規サポートIDの追加は　０件です。" _
& vbCrLf & vbCrLf & "start:　" & st & vbCrLf & "end  :　" & Now
Else
MsgBox "MCMデータ取得完了" & vbCrLf & vbCrLf & "新規サポートIDの追加は　【" & New_SupportID & "件】です。" _
& vbCrLf & vbCrLf & "start:　" & st & vbCrLf & "end  :　" & Now
End If
McmEndData.CommitResult
End Sub
'*******************************
Sub Edit_seiban()
'保守製番情報更新
'*******************************
Dim mainWS As Worksheet                     '//関数セット先シート用
Dim tempWS As Worksheet                     '//参照先リストシート用
Dim tempRow As Long                         '//参照先リスト範囲下段取得用
Dim LastRow As Long                         '//入力対象セル範囲下段取得用
Dim WSname As String                        '//参照先リストシート名入力用
Dim priorCalc As Long, calcNumber As Long, calcSource As String, calcDescription As String
WSname = "製番データ"                       '//参照先シート名をセット
Set mainWS = Worksheets("取込データ")       '//関数セット先シート名をセット
Set tempWS = Worksheets(WSname)
'// 前回製番データ削除
Application.DisplayAlerts = False
tempWS.Delete
Worksheets.Add after:=Worksheets(Worksheets.Count)
ActiveSheet.Name = WSname
Set tempWS = Worksheets(WSname)
ActiveWindow.Zoom = 80
Application.DisplayAlerts = True
Set rng = Cells(1, 1)
rngDest = ThisWorkbook.Path & "\36契約製番.csv"
Call seiban_tran
mainWS.Activate
LastRow = mainWS.Cells(Rows.Count, 2).End(xlUp).Row
tempRow = tempWS.Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(2, 46), Cells(LastRow, 47)).NumberFormatLocal = "G/標準"
priorCalc = Application.Calculation
On Error GoTo RestoreCalc
If priorCalc <> xlCalculationManual Then Application.Calculation = xlCalculationManual
'製番開始
Range(Cells(2, 44), Cells(LastRow, 44)).FormulaR1C1 = "=iferror(if(index(" & WSname & "!R2C2:R" & tempRow & "C2,match(RC[38]," & WSname & "!R2C2:R" & tempRow & "C2,1),1)=RC[38],VLOOKUP(RC[38]," & WSname & "!R2C2:R" & tempRow & "C6,4,true),NA()),"""")"
'製番終了
Range(Cells(2, 45), Cells(LastRow, 45)).FormulaR1C1 = "=iferror(if(index(" & WSname & "!R2C2:R" & tempRow & "C2,match(RC[37]," & WSname & "!R2C2:R" & tempRow & "C2,1),1)=RC[37],VLOOKUP(RC[37]," & WSname & "!R2C2:R" & tempRow & "C6,5,true),NA()),"""")"
'HW製番
Range(Cells(2, 46), Cells(LastRow, 46)).FormulaR1C1 = "=iferror(if(index(" & WSname & "!R2C2:R" & tempRow & "C2,match(RC[36]," & WSname & "!R2C2:R" & tempRow & "C2,1),1)=RC[36],VLOOKUP(RC[36]," & WSname & "!R2C2:R" & tempRow & "C6,2,true),NA()),"""")"
'SW製番
Range(Cells(2, 47), Cells(LastRow, 47)).FormulaR1C1 = "=iferror(if(index(" & WSname & "!R2C2:R" & tempRow & "C2,match(RC[35]," & WSname & "!R2C2:R" & tempRow & "C2,1),1)=RC[35],VLOOKUP(RC[35]," & WSname & "!R2C2:R" & tempRow & "C6,3,true),NA()),"""")"
If priorCalc <> xlCalculationManual Then Range(Cells(2, 44), Cells(LastRow, 47)).Calculate
Range(Cells(2, 44), Cells(LastRow, 47)).Select
Selection.Value = Selection.Value
Application.Calculation = priorCalc
On Error GoTo 0
Exit Sub
RestoreCalc:
calcNumber = Err.Number: calcSource = Err.Source: calcDescription = Err.Description
On Error Resume Next
Application.Calculation = priorCalc
On Error GoTo 0
Err.Raise calcNumber, calcSource, calcDescription
End Sub
'*******************************
Sub seiban_tran()
'保守製番データ取込
'*******************************
Dim nm As Name                              '//csv接続情報削除用
If Not McmEndData.IsCsvMode() Then
McmEndData.CopySnapshot "36契約製番.csv", rng
Else
With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & rngDest, Destination:=rng)
.Name = "export"
.FieldNames = True
.RowNumbers = False
.FillAdjacentFormulas = False
.PreserveFormatting = True
.RefreshOnFileOpen = False
.RefreshStyle = xlInsertDeleteCells
.SavePassword = False
.SaveData = True
.AdjustColumnWidth = True
.RefreshPeriod = 0
.TextFilePromptOnRefresh = False
.TextFilePlatform = 932
.TextFileStartRow = 1
.TextFileParseType = xlDelimited
.TextFileTextQualifier = xlTextQualifierDoubleQuote
.TextFileConsecutiveDelimiter = False
.TextFileTabDelimiter = False
.TextFileSemicolonDelimiter = False
.TextFileCommaDelimiter = True
.TextFileSpaceDelimiter = False
.TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1)
.TextFileTrailingMinusNumbers = True
.Refresh BackgroundQuery:=False
End With
End If
'CSV接続情報削除
Do While ActiveWorkbook.Connections.Count > 0
ActiveWorkbook.Connections.item(ActiveWorkbook.Connections.Count).Delete
Loop
For Each nm In ActiveWorkbook.Names
On Error Resume Next
nm.Delete
Next nm
End Sub
private sub adoget_mcm(byval mode as string)
static running as boolean
dim started as date, oldstatus as variant, oldinteractive as boolean
dim description as string, newcount as long, last as long
if running then exit sub
running = true
started = now: oldstatus = application.statusbar: oldinteractive = application.interactive
on error goto failed
application.interactive = false: application.screenupdating = false
set pilotws = thisworkbook.worksheets("操作シート")
if not mcmenddata.fetchmain(mode) then goto finished
set exportws = thisworkbook.worksheets("取込データ")
thisworkbook.activate: exportws.activate
' 旧加工処理のendはエラーハンドラーを通らないため、入力停止を取得区間だけに限定する。
application.interactive = oldinteractive
' adoは列単位で取得するため、csvの引用符・改行による列ずれ検査は不要。
call duplicate_item_check
call col_del
call comment_create
call edit_seiban
with thisworkbook.worksheets("店舗リスト")
last = .cells(.rows.count, 1).end(xlup).row
newcount = last - .cells(.rows.count, 2).end(xlup).row
end with
mcmenddata.commitresult
pilotws.activate
msgbox "MCMデータ取得完了" & vbcrlf & vbcrlf & "新規サポートIDの追加は　【" & newcount & "件】です。" & _
vbcrlf & vbcrlf & "start:　" & started & vbcrlf & "end  :　" & now, vbinformation
finished:
application.interactive = oldinteractive: application.statusbar = oldstatus
application.screenupdating = true: application.displayalerts = true
exit sub
failed:
description = err.description
on error resume next
mcmenddata.invalidate
application.interactive = oldinteractive: application.statusbar = oldstatus
application.screenupdating = true: application.displayalerts = true
on error goto 0
msgbox "新MCMの取得・加工を中断しました。出力には使用せず、原因を確認して取得し直してください。" & vbcrlf & description, vbexclamation
end sub
