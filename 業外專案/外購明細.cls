VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Output()
    Dim wbSource As Workbook, wbTarget As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet, wsLotList As Worksheet
    Dim x As String, lot As String
    Dim i As Long, j As Long

    Set wbSource = ThisWorkbook
    x = InputBox("請輸入目標檔案名稱")

    On Error Resume Next
    Set wbTarget = Workbooks(x)
    On Error GoTo 0
    If wbTarget Is Nothing Then
        MsgBox "請確認 '" & x & "' 已開啟。", vbExclamation
        Exit Sub
    End If

    If Not WorksheetExists("外購明細 (2)", wbSource) Then
        MsgBox "找不到來源工作表 '外購明細 (2)'", vbExclamation
        Exit Sub
    End If
    If Not WorksheetExists("出貨批號", wbSource) Then
        MsgBox "找不到來源工作表 '出貨批號'", vbExclamation
        Exit Sub
    End If

    Set wsSource = wbSource.Sheets("外購明細 (2)")
    Set wsLotList = wbSource.Sheets("出貨批號")
    Set wsTarget = wbTarget.Sheets(1)


    Dim lotList As Collection
    Set lotList = New Collection
    Dim lastLotRow As Long: lastLotRow = wsLotList.Cells(wsLotList.Rows.Count, "D").End(xlUp).Row
    For i = 4 To lastLotRow
        lot = Trim(wsLotList.Cells(i, "D").Value)
        If lot <> "" Then lotList.Add lot
    Next i


    Dim lastRow As Long: lastRow = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    Dim sourceData As Collection: Set sourceData = New Collection
    For i = 15 To lastRow
        lot = Trim(wsSource.Cells(i, "D").Value)
        If lot <> "" Then
            If ExistsInCollection(lotList, lot) Then
                Dim rowData(1 To 21) As Variant
                For j = 1 To 21
                    rowData(j) = wsSource.Cells(i, j).Value
                Next j
                sourceData.Add rowData
            End If
        End If
    Next i

    If sourceData.Count = 0 Then
        MsgBox "找不到任何符合的批號資料", vbExclamation
        Exit Sub
    End If


    Dim colMap As Object: Set colMap = CreateObject("Scripting.Dictionary")
    colMap.Add 4, "D": colMap.Add 6, "W": colMap.Add 7, "P1": colMap.Add 8, "E"
    colMap.Add 9, "F": colMap.Add 10, "D0": colMap.Add 11, "D1"
    colMap.Add 12, "P0": colMap.Add 13, "P0 x 10": colMap.Add 14, "P2"
    colMap.Add 15, "A0": colMap.Add 16, "A1": colMap.Add 17, "B0"
    colMap.Add 18, "B1": colMap.Add 19, "K0": colMap.Add 20, "K1"
    colMap.Add 21, "t"

    Dim nameColDict As Object: Set nameColDict = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = wsTarget.Cells(30, wsTarget.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        Dim colTitle As String: colTitle = Trim(wsTarget.Cells(30, i).Value)
        If Len(colTitle) > 0 Then nameColDict(colTitle) = i
    Next i


    For i = 1 To sourceData.Count
        Dim dataItem() As Variant
        dataItem = sourceData(i)
        wsTarget.Cells(34 + i, 2).Value = dataItem(4)
        wsTarget.Cells(34 + i, 19).Value = "OK"
        Dim key
        For Each key In colMap.Keys
            If nameColDict.Exists(colMap(key)) Then
                With wsTarget.Cells(34 + i, nameColDict(colMap(key)))
                    .Value = Round(dataItem(key), 3)
                    .NumberFormat = "0.000"
                End With
            End If
        Next key
    Next i

    Dim firstRow As Long
    firstRow = 4
    
    Dim firstC As String, firstA As String, firstB As String, e4 As String
    
    Do While wsLotList.Cells(firstRow, "C").Value = "" And firstRow <= lastLotRow
        firstRow = firstRow + 1
    Loop
    
    If firstRow <= lastLotRow Then
        firstC = wsLotList.Cells(firstRow, "C").Value
        firstA = wsLotList.Cells(firstRow, "A").Value
        firstB = wsLotList.Cells(firstRow, "B").Value
    Else
        firstC = "": firstA = "": firstB = ""
    End If
    
    e4 = wsLotList.Range("E4").Value
    h4 = wsLotList.Range("H4").Value
    wsTarget.Range("H5").Value = firstC
    wsTarget.Range("H7").Value = e4
    wsTarget.Range("K48").Value = h4
    wsTarget.Range("N5").Value = firstA
    wsTarget.Range("C5").Value = firstB

    MsgBox "共貼上 " & sourceData.Count & " 筆資料", vbInformation
End Sub

Function ExistsInCollection(col As Collection, item As String) As Boolean
    Dim x
    For Each x In col
        If x = item Then
            ExistsInCollection = True
            Exit Function
        End If
    Next x
    ExistsInCollection = False
End Function
Function WorksheetExists(sheetName As String, wb As Workbook) As Boolean
    On Error Resume Next
    WorksheetExists = Not wb.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

