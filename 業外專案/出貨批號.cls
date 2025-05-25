VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub RunSelectedShipmentMode()
    With ThisWorkbook.Sheets("出貨批號")
        On Error GoTo OptionError
        If .OLEObjects("OptionButton2").Object.Value = True Then
            Call ProcessAllShipments
        ElseIf .OLEObjects("OptionButton1").Object.Value = True Then
            Call ProcessAllShipments2
        Else
            MsgBox "請先選擇一個處理模式 (OptionButton1 或 OptionButton2)", vbExclamation
        End If
        Exit Sub
OptionError:
        MsgBox "找不到 OptionButton1 或 OptionButton2。", vbCritical
    End With
End Sub
Sub ProcessAllShipments()
    Dim wsLotList As Worksheet: Set wsLotList = ThisWorkbook.Sheets("出貨批號")
    Dim currentRow As Long, sectionStart As Long, sectionEnd As Long
    Dim sectionNum As Long: sectionNum = 0
    Dim lastRow As Long: lastRow = wsLotList.Cells(wsLotList.Rows.Count, "D").End(xlUp).Row
    Dim fileName As String

    currentRow = 4
    Do While currentRow <= lastRow
        If Trim(wsLotList.Cells(currentRow, "D").Value) <> "" Then
            sectionStart = currentRow
            Do While currentRow <= lastRow And Trim(wsLotList.Cells(currentRow, "D").Value) <> ""
                currentRow = currentRow + 1
            Loop
            sectionEnd = currentRow - 1
            sectionNum = sectionNum + 1

            fileName = InputBox("請輸入第 " & sectionNum & " 段的出貨報告")
            If Trim(fileName) <> "" Then
                Call Output_OneShipment(wsLotList, sectionStart, sectionEnd, fileName)
            Else
                MsgBox "第 " & sectionNum & " 段未輸入檔名，略過。"
            End If
        Else
            currentRow = currentRow + 1
        End If
    Loop
    MsgBox "完成 " & sectionNum & " 組出貨報表"
End Sub
Sub Output_OneShipment(wsLotList As Worksheet, sectionStart As Long, sectionEnd As Long, fileName As String)
    Dim lotList As Collection: Set lotList = New Collection
    Dim i As Long, lot As String

    For i = sectionStart To sectionEnd
        lot = Trim(wsLotList.Cells(i, "D").Value)
        If lot <> "" Then lotList.Add lot
    Next i

    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    On Error Resume Next
    Set wbTarget = Workbooks(fileName)
    On Error GoTo 0
    If wbTarget Is Nothing Then
        MsgBox "找不到已開啟的活頁簿：" & fileName, vbExclamation
        Exit Sub
    End If
    Set wsTarget = wbTarget.Sheets(1)

    Dim wsSource As Worksheet: Set wsSource = ThisWorkbook.Sheets("外購明細 (2)")
    Dim lastRow As Long: lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    Dim sourceData As Collection: Set sourceData = New Collection
    For i = 15 To lastRow
        If InCollection(lotList, Trim(wsSource.Cells(i, "D").Value)) Then
            Dim rowData As Variant
            ReDim rowData(1 To 21)
            Dim j As Long
            For j = 1 To 21
                rowData(j) = wsSource.Cells(i, j).Value
            Next j
            sourceData.Add rowData
        End If
    Next i

    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    colMap.Add 4, "D": colMap.Add 6, "W": colMap.Add 7, "P1": colMap.Add 8, "E"
    colMap.Add 9, "F": colMap.Add 10, "D0": colMap.Add 11, "D1"
    colMap.Add 12, "P0": colMap.Add 13, "P0 x 10": colMap.Add 14, "P2"
    colMap.Add 15, "A0": colMap.Add 16, "A1": colMap.Add 17, "B0"
    colMap.Add 18, "B1": colMap.Add 19, "K0": colMap.Add 20, "K1": colMap.Add 21, "t"

    Dim nameColDict As Object: Set nameColDict = CreateObject("Scripting.Dictionary")
    Dim headerCol As Long
    For headerCol = 1 To wsTarget.Cells(30, wsTarget.Columns.Count).End(xlToLeft).Column
        Dim colTitle As String: colTitle = Trim(wsTarget.Cells(30, headerCol).Value)
        If Len(colTitle) > 0 Then nameColDict(colTitle) = headerCol
    Next headerCol

    Dim startRow As Long: startRow = 35
    Dim endRow As Long: endRow = startRow + sourceData.Count - 1
    For i = 1 To sourceData.Count
        Dim dataItem As Variant
        dataItem = sourceData(i)
        wsTarget.Cells(startRow + i - 1, 2).Value = dataItem(4)
        wsTarget.Cells(startRow + i - 1, 19).Value = "OK"

        Dim key
        For Each key In colMap.Keys
            If nameColDict.Exists(colMap(key)) Then
                With wsTarget.Cells(startRow + i - 1, nameColDict(colMap(key)))
                    .Value = Round(dataItem(key), 3)
                    .NumberFormat = "0.000"
                End With
            End If
        Next key
    Next i

    Dim firstC As String, firstA As String, firstB As String, e4 As String, h4 As String
    firstC = wsLotList.Cells(sectionStart, "C").Value
    firstA = wsLotList.Cells(sectionStart, "A").Value
    firstB = wsLotList.Cells(sectionStart, "B").Value
    e4 = wsLotList.Cells(sectionStart, "E").Value
    h4 = wsLotList.Cells(sectionStart, "H").Value

    wsTarget.Range("H5").Value = firstC
    wsTarget.Range("H7").Value = e4
    wsTarget.Range("K48").Value = h4
    wsTarget.Range("N5").Value = firstA
    wsTarget.Range("C5").Value = firstB


End Sub

Function InCollection(col As Collection, item As String) As Boolean
    Dim x
    For Each x In col
        If x = item Then
            InCollection = True
            Exit Function
        End If
    Next x
    InCollection = False
End Function
Sub ProcessAllShipments2()
    Dim wsLotList As Worksheet: Set wsLotList = ThisWorkbook.Sheets("出貨批號")
    Dim currentRow As Long, sectionStart As Long, sectionEnd As Long
    Dim sectionNum As Long: sectionNum = 0
    Dim lastRow As Long: lastRow = wsLotList.Cells(wsLotList.Rows.Count, "D").End(xlUp).Row
    Dim fileName As String

    currentRow = 4
    Do While currentRow <= lastRow
        If Trim(wsLotList.Cells(currentRow, "D").Value) <> "" Then
            sectionStart = currentRow
            Do While currentRow <= lastRow And Trim(wsLotList.Cells(currentRow, "D").Value) <> ""
                currentRow = currentRow + 1
            Loop
            sectionEnd = currentRow - 1
            sectionNum = sectionNum + 1

            fileName = InputBox("請輸入第 " & sectionNum & " 段的報表檔名（不含副檔名，請先開啟）")
            If Trim(fileName) <> "" Then
                Call ComputeStatsForLotGroup(wsLotList, sectionStart, sectionEnd, fileName)
            Else
                MsgBox "第 " & sectionNum & " 段未輸入檔名，略過。"
            End If
        Else
            currentRow = currentRow + 1
        End If
    Loop
    MsgBox "完成 " & sectionNum & " 組出貨報告"
End Sub

Sub ComputeStatsForLotGroup(wsLotList As Worksheet, sectionStart As Long, sectionEnd As Long, fileName As String)
    Dim lotList As Collection: Set lotList = New Collection
    Dim i As Long, lot As String

    For i = sectionStart To sectionEnd
        lot = Trim(wsLotList.Cells(i, "D").Value)
        If lot <> "" Then lotList.Add lot
    Next i

    Dim wbTarget As Workbook, wsTarget As Worksheet
    On Error Resume Next
    Set wbTarget = Workbooks(fileName)
    On Error GoTo 0
    If wbTarget Is Nothing Then
        MsgBox "出貨報告沒打開：" & fileName, vbExclamation
        Exit Sub
    End If
    Set wsTarget = wbTarget.Sheets(1)

    Dim firstC As String, firstA As String, firstB As String, e4 As String
    firstC = wsLotList.Cells(sectionStart, "C").Value
    firstA = wsLotList.Cells(sectionStart, "A").Value
    firstB = wsLotList.Cells(sectionStart, "B").Value
    e4 = wsLotList.Cells(sectionStart, "E").Value
    
    wsTarget.Range("C5").Value = firstB
    wsTarget.Range("H5").Value = firstC
    wsTarget.Range("H7").Value = e4
    wsTarget.Range("N5").Value = firstA

    Dim wsSource As Worksheet: Set wsSource = ThisWorkbook.Sheets("外購明細 (2)")
    Dim lastRow As Long: lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row

    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    colMap.Add 6, "W": colMap.Add 7, "P1": colMap.Add 8, "E"
    colMap.Add 9, "F": colMap.Add 10, "D0": colMap.Add 11, "D1"
    colMap.Add 12, "P0": colMap.Add 13, "P0 x 10": colMap.Add 14, "P2"
    colMap.Add 15, "A0": colMap.Add 16, "A1": colMap.Add 17, "B0"
    colMap.Add 18, "B1": colMap.Add 19, "K0": colMap.Add 20, "K1": colMap.Add 21, "t"

    Dim nameColDict As Object: Set nameColDict = CreateObject("Scripting.Dictionary")
    Dim headerCol As Long
    For headerCol = 1 To wsTarget.Cells(30, wsTarget.Columns.Count).End(xlToLeft).Column
        Dim colTitle As String: colTitle = Trim(wsTarget.Cells(30, headerCol).Value)
        If Len(colTitle) > 0 Then nameColDict(colTitle) = headerCol
    Next headerCol

    Dim srcCol As Variant, tgtColTitle As String
    Dim values As Collection
    Dim statMin As Double, statMax As Double, statAvg As Double, total As Double
    Dim rowVal As Variant

    For Each srcCol In colMap.Keys
        tgtColTitle = colMap(srcCol)
        If nameColDict.Exists(tgtColTitle) Then
            Set values = New Collection
            For i = 15 To lastRow
                lot = Trim(wsSource.Cells(i, "D").Value)
                If InCollection(lotList, lot) Then
                    rowVal = wsSource.Cells(i, srcCol).Value
                    If IsNumeric(rowVal) Then values.Add CDbl(rowVal)
                End If
            Next i

            If values.Count > 0 Then
                statMin = values(1): statMax = values(1): total = 0
                For Each rowVal In values
                    If rowVal < statMin Then statMin = rowVal
                    If rowVal > statMax Then statMax = rowVal
                    total = total + rowVal
                Next
                statAvg = total / values.Count

                Dim tgtCol As Long: tgtCol = nameColDict(tgtColTitle)
                wsTarget.Cells(35, tgtCol).Value = Round(statMin, 3)
                wsTarget.Cells(36, tgtCol).Value = Round(statMax, 3)
                wsTarget.Cells(37, tgtCol).Value = Round(statAvg, 3)
            End If
        End If
    Next srcCol

    Dim h4 As String: h4 = wsLotList.Cells(sectionStart, "H").Value
    wsTarget.Range("K38").Value = h4
End Sub

