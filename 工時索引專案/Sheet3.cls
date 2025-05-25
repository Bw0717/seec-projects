VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim wsIndex As Worksheet, wsData As Worksheet
    Dim keyword As String, searchMode As String
    Dim lastRow As Long, i As Long, outRow As Long
    Dim targetType As String

    Set wsIndex = ThisWorkbook.Sheets("索引")
    Set wsData = ThisWorkbook.Sheets("資料庫")

    keyword = Trim(wsIndex.OLEObjects("TextBox1").Object.text)
    If wsIndex.OLEObjects("OptionButton1").Object.Value = True Then
        searchMode = "TYPE"
    ElseIf wsIndex.OLEObjects("OptionButton2").Object.Value = True Then
        searchMode = "件號"
    ElseIf wsIndex.OLEObjects("OptionButton3").Object.Value = True Then
        searchMode = "件號搜尋TYPE"
    Else
        MsgBox "請選擇搜尋模式", vbExclamation: Exit Sub
    End If


    wsIndex.Range("B10:F1000").ClearContents

    If keyword = "" Then
        MsgBox "請輸入查詢關鍵字", vbExclamation: Exit Sub
    End If

    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    With wsIndex.Range("B10:G10")
        .Value = Array("TYPE", "生產件號", "摘要", "資源", "資源備註", "連結")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 255)
        .Borders.LineStyle = xlContinuous
    End With
    outRow = 11


    If UCase(searchMode) = "TYPE" Then
        targetType = keyword
        If targetType = "" Then
            MsgBox "找不到對應的TYPE", vbExclamation: Exit Sub
        End If

    ElseIf UCase(searchMode) = "件號" Then
        For i = 2 To lastRow
            If Trim(wsData.Cells(i, 2).Value) = Trim(keyword) Then
                wsIndex.Cells(outRow, 2).Resize(1, 5).Value = wsData.Range(wsData.Cells(i, 1), wsData.Cells(i, 5)).Value
                outRow = outRow + 1
            End If
        Next i
        If outRow = 11 Then
            MsgBox "找不到對應的件號", vbExclamation: Exit Sub
        End If
    ElseIf UCase(searchMode) = "件號搜尋TYPE" Then
        targetType = ""
        For i = 2 To lastRow
            If Trim(wsData.Cells(i, 2).Value) = Trim(keyword) Then
                targetType = wsData.Cells(i, 1).Value
                Exit For
            End If
        Next i
        If targetType = "" Then
            MsgBox "找不到對應的件號", vbExclamation: Exit Sub
        End If

    Else
        MsgBox "請在 B1 輸入搜尋模式（TYPE 或 件號）", vbExclamation: Exit Sub
    End If

    For i = 2 To lastRow
        If wsData.Cells(i, 1).Value = targetType Then
            wsIndex.Cells(outRow, 2).Resize(1, 5).Value = wsData.Range(wsData.Cells(i, 1), wsData.Cells(i, 5)).Value
            outRow = outRow + 1
        End If
    Next i
    If outRow > 11 Then
        wsIndex.Range("G11:G" & outRow - 1).FormulaR1C1 = _
            "=IFERROR(HYPERLINK(VLOOKUP(RC[-4], Hyperlink!C1:C2, 2, FALSE), ""連結""), """")"
    End If
    Application.CutCopyMode = False
    wsIndex.Columns.AutoFit
    MsgBox "完成查詢，共列出 " & outRow - 11 & " 筆資料", vbInformation
End Sub



Private Sub OptionButton1_Click()

End Sub

Private Sub OptionButton2_Click()

End Sub

Private Sub OptionButton3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub
Private Sub Workbook_Open()
    With Sheets("索引").OLEObjects("TextBox1").Object
        .text = "請輸入件號或TYPE"
        .ForeColor = RGB(150, 150, 150)
    End With
End Sub

Private Sub TextBox1_Enter()
    With Me.OLEObjects("TextBox1").Object
        If .text = "請輸入件號或TYPE" Then
            .text = ""
            .ForeColor = RGB(0, 0, 0)
        End If
    End With
End Sub

