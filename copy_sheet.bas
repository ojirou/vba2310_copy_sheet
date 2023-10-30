Attribute VB_Name = "Module1"
Sub CopySheet()
    Dim ws As Worksheet, ws1 As Worksheet
    ' シート「コピー元」が存在するか確認
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("コピー元")
    Set ws1 = ThisWorkbook.Sheets("コピー先")
    On Error GoTo 0
    ' シート「コピー元」が存在する場合、コピーして「コピー先」という名前で保存
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' 警告を非表示にする
        ws.Copy Before:=ThisWorkbook.Sheets(1)
        Application.DisplayAlerts = True ' 警告を非表示にする
        ' 新しく作成されたシートに名前を設定
        If Not ws1 Is Nothing Then
            copyDate = Format(Now, "yyyymmdd_hhmmss")
            ActiveSheet.Name = "コピー先_" & copyDate
            MsgBox "シート「コピー元」が「コピー先_" & copyDate & "」としてコピーされました。"
        Else
            ActiveSheet.Name = "コピー先"
            MsgBox "シート「コピー元」が「コピー先」としてコピーされました。"
        End If
    Else
        MsgBox "シート「コピー元」が存在しません。"
    End If
End Sub
