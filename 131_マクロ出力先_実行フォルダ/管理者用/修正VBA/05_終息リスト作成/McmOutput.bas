Attribute VB_Name = "McmOutput"
Option Explicit
' 出力は実行用Excelと同じフォルダに保存する。
Public Function PathFor(ByVal fileName As String) As String
    Dim fs As Object, bad As Variant
    If Len(ThisWorkbook.path) = 0 Then Err.Raise vbObjectError + 860, "McmOutput", "実行用Excelを保存してから実行してください。"
    If Len(fileName) = 0 Then Err.Raise vbObjectError + 861, "McmOutput", "出力ファイル名が空です。"
    If fileName = "." Or fileName = ".." Or Right$(fileName, 1) = "." Or Right$(fileName, 1) = " " Then _
        Err.Raise vbObjectError + 862, "McmOutput", "出力ファイル名の末尾にピリオド・空白は使用できません。"
    For Each bad In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(fileName, CStr(bad)) > 0 Then Err.Raise vbObjectError + 862, "McmOutput", "出力ファイル名に使用できない文字があります。"
    Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    PathFor = fs.BuildPath(ThisWorkbook.Path, fileName)
End Function
' 支払い一覧は元ブック自体を加工するため、別名保存で設定の参照元を変えない。
Public Function SaveCopyReport(ByVal fileName As String) As Boolean
    Dim target As String, fs As Object, book As Workbook
    target = PathFor(fileName)
    For Each book In Application.Workbooks
        If StrComp(book.FullName, target, vbTextCompare) = 0 Then _
            Err.Raise vbObjectError + 864, "McmOutput", "同名の出力ファイルを閉じてから実行してください。"
    Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FileExists(target) Then
        If MsgBox("同名の出力ファイルが存在します。上書きしますか？" & vbCrLf & target, _
                  vbYesNoCancel + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    ThisWorkbook.SaveCopyAs target
    SaveCopyReport = True
End Function


