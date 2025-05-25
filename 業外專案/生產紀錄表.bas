Attribute VB_Name = "CreateNewMacros"
Sub InsertRowAtRow8()
    Dim ws As Worksheet
    Dim dValue As String
    Dim baseText As String
    Dim serialNumber As Long
    Dim newSerial As String
    Set ws = ThisWorkbook.Sheets("生產紀錄表")
    dValue = ws.Cells(6, 4).Value
    ws.Rows("13:17").Insert Shift:=xlDown
    ws.Cells(13, 1).Value = "VBA_INSERT"
    ws.Range("A6:M10").Copy
    ws.Range("A13").PasteSpecial Paste:=xlPasteValues
    ws.Range("D6:M10").ClearContents
    Application.CutCopyMode = False
    ThisWorkbook.Save
End Sub
Sub ProtectAndAutoExpand()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("生產紀錄表")
    On Error Resume Next
    ws.Unprotect Password:="123456"

    ws.Cells.Locked = False
    ws.Range("A1:M4").Locked = True

    ws.Protect Password:="123456", _
                AllowInsertingRows:=True, _
                AllowDeletingRows:=False, _
                AllowFormattingCells:=True
    MsgBox "範圍鎖定！", vbInformation, "完成"
End Sub
Sub FINDMIN()
    Dim fileName As String
    Dim fullFileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim foundWorkbook As Boolean
    Dim sheetName As String
    Dim currentCell As Range
    Dim minCount As Long
    Dim endRow As Long
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    
    Set sourceWB = ThisWorkbook
    Set sourceWS = sourceWB.Sheets("外購明細")
    fileName = InputBox("請輸入檔案名稱")
    sheetName = InputBox("請輸入料號")
    If fileName = "" Then
        MsgBox "未輸入檔案名稱，已取消。"
        Exit Sub
    End If
    Set wbTarget = Workbooks(fileName)
    Set wsTarget = wbTarget.Sheets(sheetName)
    fullFileName = fileName & ".xlsx"
    foundWorkbook = False
    For Each wb In Application.Workbooks
        If wb.Name = fullFileName Then
            foundWorkbook = True
            Exit For
        End If
    Next wb
    If Not foundWorkbook Then
        MsgBox "檔案「" & fullFileName & "」尚未開啟！"
        Exit Sub
    End If
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "找不到工作表「" & sheetName & "」"
        Exit Sub
    End If
    Set currentCell = ws.Range("B36")
    minCount = 0
    Do While currentCell.Value <> "MIN" And Not IsEmpty(currentCell)
        minCount = minCount + 1
        Set currentCell = currentCell.Offset(1, 0)
    Loop
    If currentCell.Value = "MIN" Then
    Else
        MsgBox "在該範圍中找不到 MIN"
    End If
    With sourceWB
        .Activate
        .Sheets("外購明細").Activate
    End With
    sourceWS.Rows("2:" & 1 + minCount).Insert Shift:=xlDown
    
    Dim sourceRange As Range
    Set sourceRange = wsTarget.Range("B36").Resize(minCount, 1)
    Dim targetRange As Range
    Set targetRange = sourceWS.Range("B2").Resize(minCount, 1)
    targetRange.Value = sourceRange.Value
        
    Dim sourceRange2 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "W" Then
            Set sourceRange2 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange2 Is Nothing Then
        Dim targetRange2 As Range
        Set targetRange2 = sourceWS.Range("D2").Resize(minCount, 1)
        targetRange2.Value = sourceRange2.Value
    End If
    
    Dim sourceRange3 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P1" Then
            Set sourceRange3 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange3 Is Nothing Then
        Dim targetRange3 As Range
        Set targetRange3 = sourceWS.Range("E2").Resize(minCount, 1)
        targetRange3.Value = sourceRange3.Value
    End If
    
    Dim sourceRange4 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "E" Then
            Set sourceRange4 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange4 Is Nothing Then
        Dim targetRange4 As Range
        Set targetRange4 = sourceWS.Range("F2").Resize(minCount, 1)
        targetRange4.Value = sourceRange4.Value
    End If
    
    Dim sourceRange5 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "F" Then
            Set sourceRange5 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange5 Is Nothing Then
        Dim targetRange5 As Range
        Set targetRange5 = sourceWS.Range("G2").Resize(minCount, 1)
        targetRange5.Value = sourceRange5.Value
    End If
    
    Dim sourceRange6 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "D0" Then
            Set sourceRange6 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange6 Is Nothing Then
        Dim targetRange6 As Range
        Set targetRange6 = sourceWS.Range("H2").Resize(minCount, 1)
        targetRange6.Value = sourceRange6.Value
    End If
    
    Dim sourceRange7 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "D1" Then
            Set sourceRange7 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange7 Is Nothing Then
        Dim targetRange7 As Range
        Set targetRange7 = sourceWS.Range("I2").Resize(minCount, 1)
        targetRange7.Value = sourceRange7.Value
    End If

    
    Dim sourceRange8 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P0" Then
            Set sourceRange8 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange8 Is Nothing Then
        Dim targetRange8 As Range
        Set targetRange8 = sourceWS.Range("J2").Resize(minCount, 1)
        targetRange8.Value = sourceRange8.Value
    End If

    Dim sourceRange9 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P0 x 10" Then
            Set sourceRange9 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange9 Is Nothing Then
        Dim targetRange9 As Range
        Set targetRange9 = sourceWS.Range("K2").Resize(minCount, 1)
        targetRange9.Value = sourceRange9.Value
    End If

    Dim sourceRange10 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P2" Then
            Set sourceRange10 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange10 Is Nothing Then
        Dim targetRange10 As Range
        Set targetRange10 = sourceWS.Range("L2").Resize(minCount, 1)
        targetRange10.Value = sourceRange10.Value
    End If

    Dim sourceRange11 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "A0" Then
            Set sourceRange11 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    
    If Not sourceRange11 Is Nothing Then
        Dim targetRange11 As Range
        Set targetRange11 = sourceWS.Range("M2").Resize(minCount, 1)
        targetRange11.Value = sourceRange11.Value
    End If

    Dim sourceRange12 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "A1" Then
            Set sourceRange12 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange12 Is Nothing Then
        Dim targetRange12 As Range
        Set targetRange12 = sourceWS.Range("N2").Resize(minCount, 1)
        targetRange12.Value = sourceRange12.Value
    End If

    Dim sourceRange13 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "B0" Then
            Set sourceRange13 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange13 Is Nothing Then
        Dim targetRange13 As Range
        Set targetRange13 = sourceWS.Range("O2").Resize(minCount, 1)
        targetRange13.Value = sourceRange13.Value
    End If

    Dim sourceRange14 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "B1" Then
            Set sourceRange14 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange14 Is Nothing Then
        Dim targetRange14 As Range
        Set targetRange14 = sourceWS.Range("P2").Resize(minCount, 1)
        targetRange14.Value = sourceRange14.Value
    End If

    Dim sourceRange15 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "K0" Then
            Set sourceRange15 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange15 Is Nothing Then
        Dim targetRange15 As Range
        Set targetRange15 = sourceWS.Range("Q2").Resize(minCount, 1)
        targetRange15.Value = sourceRange15.Value
    End If

    Dim sourceRange16 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "K1" Then
            Set sourceRange16 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange16 Is Nothing Then
        Dim targetRange16 As Range
        Set targetRange16 = sourceWS.Range("R2").Resize(minCount, 1)
        targetRange16.Value = sourceRange16.Value
    End If

    Dim sourceRange17 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "t" Then
            Set sourceRange17 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange17 Is Nothing Then
        Dim targetRange17 As Range
        Set targetRange17 = sourceWS.Range("S2").Resize(minCount, 1)
        targetRange17.Value = sourceRange17.Value
    End If

    Dim sourceRange18 As Range
    Set sourceRange18 = wsTarget.Range("C7").Resize(1, 1)
    
    Dim targetRange18 As Range
    Set targetRange18 = sourceWS.Range("C2").Resize(minCount, 1)
    
    targetRange18.Value = sourceRange18.Value
        
    If wsTarget.Range("M7").Value <> "" Then
        Set sourceRange19 = wsTarget.Range("M7")
    ElseIf wsTarget.Range("N7").Value <> "" Then
        Set sourceRange19 = wsTarget.Range("N7")
    Else
        MsgBox "M7 和 N7 有問題，無法複製", vbExclamation
        Exit Sub
    End If
    Dim targetRange19 As Range
    Set targetRange19 = sourceWS.Range("A2").Resize(minCount, 1)
    targetRange19.Value = sourceRange19.Value
    With sourceWB
        .Activate
        .Sheets("外購明細").Activate
    End With
    ThisWorkbook.Save
End Sub
Sub 鎖定特定儲存格()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("品保專用") '
    ws.Unprotect
    ws.Cells.Locked = False
    ws.Range("B3:K7").Locked = True
    ws.Range("B10:N14").Locked = True
    ws.Protect
End Sub
Sub InsertRowAtRow9()
    Dim ws As Worksheet
    Dim dValue As String
    Dim baseText As String
    Dim serialNumber As Long
    Dim newSerial As String
    Set ws = ThisWorkbook.Sheets("外購明細 (2)")
    dValue = ws.Cells(6, 4).Value
    ws.Rows("15:20").Insert Shift:=xlDown
    ws.Range("A15:AB20").Interior.ColorIndex = xlNone
    ws.Cells(15, 1).Value = "VBA_INSERT"
    ws.Range("A7:AB12").Copy
    ws.Range("A15").PasteSpecial Paste:=xlPasteValues
    ws.Range("A7:AB12").ClearContents
    Application.CutCopyMode = False
    ThisWorkbook.Save
End Sub
Sub FINDMIN2()
    Dim fileName As String
    Dim fullFileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim foundWorkbook As Boolean
    Dim sheetName As String
    Dim currentCell As Range
    Dim minCount As Long
    Dim endRow As Long
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    
    Set sourceWB = ThisWorkbook
    Set sourceWS = sourceWB.Sheets("外購明細 (2)")
    
    fileName = InputBox("請輸入檔案名稱")
    sheetName = InputBox("請輸入料號")
    If fileName = "" Then
        MsgBox "未輸入檔案名稱，已取消。"
        Exit Sub
    End If
    Set wbTarget = Workbooks(fileName)
    Set wsTarget = wbTarget.Sheets(sheetName)

    fullFileName = fileName & ".xlsx"
    foundWorkbook = False
    
    For Each wb In Application.Workbooks
        If wb.Name = fullFileName Then
            foundWorkbook = True
            Exit For
        End If
    Next wb
    
    If Not foundWorkbook Then
        MsgBox "檔案「" & fullFileName & "」未開啟！"
        Exit Sub
    End If
    
    
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "找不到工作表「" & sheetName & "」"
        Exit Sub
    End If
    
    Set currentCell = ws.Range("B36")
    minCount = 0
    
    Do While currentCell.Value <> "MIN" And Not IsEmpty(currentCell)
        minCount = minCount + 1
        Set currentCell = currentCell.Offset(1, 0)
    Loop
    
    If currentCell.Value = "MIN" Then

    Else
        MsgBox "在該範圍中找不到 MIN"
    End If
    
    


    
    With sourceWB
        .Activate
        .Sheets("外購明細 (2)").Activate
    End With

    sourceWS.Rows("15:" & 14 + minCount).Insert Shift:=xlDown
    
    sourceWS.Range("A15:AB" & (15 + minCount - 1)).Interior.ColorIndex = xlNone

    
    
    Dim sourceRange As Range
    Set sourceRange = wsTarget.Range("B36").Resize(minCount, 1)
    Dim targetRange As Range
    Set targetRange = sourceWS.Range("D15").Resize(minCount, 1)
    targetRange.Value = sourceRange.Value
        
        
        
    Dim sourceRange2 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "W" Then
            Set sourceRange2 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange2 Is Nothing Then
        Dim targetRange2 As Range
        Set targetRange2 = sourceWS.Range("F15").Resize(minCount, 1)
        targetRange2.Value = sourceRange2.Value
    End If
    
    Dim sourceRange3 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P1" Then
            Set sourceRange3 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange3 Is Nothing Then
        Dim targetRange3 As Range
        Set targetRange3 = sourceWS.Range("G15").Resize(minCount, 1)
        targetRange3.Value = sourceRange3.Value
    End If
    
    Dim sourceRange4 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "E" Then
            Set sourceRange4 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange4 Is Nothing Then
        Dim targetRange4 As Range
        Set targetRange4 = sourceWS.Range("H15").Resize(minCount, 1)
        targetRange4.Value = sourceRange4.Value
    End If
    
    Dim sourceRange5 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "F" Then
            Set sourceRange5 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange5 Is Nothing Then
        Dim targetRange5 As Range
        Set targetRange5 = sourceWS.Range("I15").Resize(minCount, 1)
        targetRange5.Value = sourceRange5.Value
    End If
    
    Dim sourceRange6 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "D0" Then
            Set sourceRange6 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange6 Is Nothing Then
        Dim targetRange6 As Range
        Set targetRange6 = sourceWS.Range("J15").Resize(minCount, 1)
        targetRange6.Value = sourceRange6.Value
    End If
    
    Dim sourceRange7 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "D1" Then
            Set sourceRange7 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange7 Is Nothing Then
        Dim targetRange7 As Range
        Set targetRange7 = sourceWS.Range("K15").Resize(minCount, 1)
        targetRange7.Value = sourceRange7.Value
    End If

    
    Dim sourceRange8 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P0" Then
            Set sourceRange8 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange8 Is Nothing Then
        Dim targetRange8 As Range
        Set targetRange8 = sourceWS.Range("L15").Resize(minCount, 1)
        targetRange8.Value = sourceRange8.Value
    End If

    Dim sourceRange9 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P0 x 10" Then
            Set sourceRange9 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange9 Is Nothing Then
        Dim targetRange9 As Range
        Set targetRange9 = sourceWS.Range("M15").Resize(minCount, 1)
        targetRange9.Value = sourceRange9.Value
    End If

    Dim sourceRange10 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "P2" Then
            Set sourceRange10 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange10 Is Nothing Then
        Dim targetRange10 As Range
        Set targetRange10 = sourceWS.Range("N15").Resize(minCount, 1)
        targetRange10.Value = sourceRange10.Value
    End If

    Dim sourceRange11 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "A0" Then
            Set sourceRange11 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange11 Is Nothing Then
        Dim targetRange11 As Range
        Set targetRange11 = sourceWS.Range("O15").Resize(minCount, 1)
        targetRange11.Value = sourceRange11.Value
    End If

    Dim sourceRange12 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "A1" Then
            Set sourceRange12 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange12 Is Nothing Then
        Dim targetRange12 As Range
        Set targetRange12 = sourceWS.Range("P15").Resize(minCount, 1)
        targetRange12.Value = sourceRange12.Value
    End If

    Dim sourceRange13 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "B0" Then
            Set sourceRange13 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange13 Is Nothing Then
        Dim targetRange13 As Range
        Set targetRange13 = sourceWS.Range("Q15").Resize(minCount, 1)
        targetRange13.Value = sourceRange13.Value
    End If

    Dim sourceRange14 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "B1" Then
            Set sourceRange14 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange14 Is Nothing Then
        Dim targetRange14 As Range
        Set targetRange14 = sourceWS.Range("R15").Resize(minCount, 1)
        targetRange14.Value = sourceRange14.Value
    End If

    Dim sourceRange15 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "K0" Then
            Set sourceRange15 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange15 Is Nothing Then
        Dim targetRange15 As Range
        Set targetRange15 = sourceWS.Range("S15").Resize(minCount, 1)
        targetRange15.Value = sourceRange15.Value
    End If

    Dim sourceRange16 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "K1" Then
            Set sourceRange16 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange16 Is Nothing Then
        Dim targetRange16 As Range
        Set targetRange16 = sourceWS.Range("T15").Resize(minCount, 1)
        targetRange16.Value = sourceRange16.Value
    End If

    Dim sourceRange17 As Range
    For col = 1 To wsTarget.Cells(31, wsTarget.Columns.Count).End(xlToLeft).Column
        If wsTarget.Cells(31, col).Value = "t" Then
            Set sourceRange17 = wsTarget.Cells(36, col).Resize(minCount, 1)
            Exit For
        End If
    Next col
    If Not sourceRange17 Is Nothing Then
        Dim targetRange17 As Range
        Set targetRange17 = sourceWS.Range("U15").Resize(minCount, 1)
        targetRange17.Value = sourceRange17.Value
    End If

    Dim sourceRange18 As Range
    Set sourceRange18 = wsTarget.Range("C7").Resize(1, 1)
    
    Dim targetRange18 As Range
    Set targetRange18 = sourceWS.Range("E15").Resize(minCount, 1)
    
    targetRange18.Value = sourceRange18.Value
        
    If wsTarget.Range("N7").Value <> "" Then
        Set sourceRange19 = wsTarget.Range("N7")
    Else
        MsgBox "M7 和 N7有問題，無法複製", vbExclamation
        Exit Sub
    End If
    
    Dim targetRange19 As Range
    Set targetRange19 = sourceWS.Range("C15").Resize(minCount, 1)
    targetRange19.Value = sourceRange19.Value
    

    Set sourceRange20 = wsTarget.Range("N5")
    Dim targetRange20 As Range
    Set targetRange20 = sourceWS.Range("A15").Resize(minCount, 1)
    targetRange20.Value = sourceRange20.Value

    With sourceWB
        .Activate
        .Sheets("外購明細 (2)").Activate
    End With
    ThisWorkbook.Save
End Sub
