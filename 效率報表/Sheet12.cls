VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call Command
    Call 合併指定工作表到新活頁簿
    Call CopyEmptyRowsToNewFileAndDelete
    Call 樞紐
End Sub

Sub Command()
    Dim txtFiles As Variant
    Dim keywords As Variant
    Dim i As Long, j As Long
    Dim ws As Worksheet
    Dim txtFile As String
    Dim TxtLine As String
    Dim RowNum As Long
    Dim stream As Object
    Dim RowCount As Long
    Dim deleteRange As Range
    Dim lastRow As Long
    Dim keyword As String
    
    txtFiles = Array("G:\效率\ADG.txt", "G:\效率\AEV.txt", "G:\效率\AM1.txt", _
                     "G:\效率\AM2.txt", "G:\效率\AMG.txt", "G:\效率\ARD.txt", "G:\效率\ARG.txt")
    keywords = Array("ADG", "AEV", "AM1", "AM2", "AMG", "ARD", "ARG")
    For i = LBound(txtFiles) To UBound(txtFiles)
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "匯入資料_" & keywords(i)
        txtFile = txtFiles(i)
        keyword = keywords(i)
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile txtFile
        RowNum = 1
        Do Until stream.EOS
            TxtLine = stream.ReadText(-2)
            ws.Cells(RowNum, 1).Value = TxtLine
            RowNum = RowNum + 1
        Loop
        stream.Close
        Set stream = Nothing
        widths = Array(5, 11, 21, 17, 17, 17, 9, 11, 12, 11, 11, 10, 10, 10, 9, 9, 8)
        For j = 1 To RowNum - 1
            TxtLine = ws.Cells(j, 1).Value
            startPos = 1
            Dim k As Integer
            For k = LBound(widths) To UBound(widths)
                ws.Cells(j, k + 1).Value = Mid(TxtLine, startPos, widths(k))
                startPos = startPos + widths(k)
            Next k
        Next j
        ws.Range("A8:Q8").Value = Array("組織", "班別", "日期", "工單號碼", "組件", "資源名稱", "繳庫量", "資源單位工時", "前一年資源單位工時", _
                                        "繳庫", "低減工時", "外包內製", "間接支援", "C工時", "修理工時", "試作工時", "換模換線")
        Dim cell As Range
        For Each cell In ws.Range("A9:Q" & RowNum - 1)
            cell.Value = Replace(cell.Value, " ", "")
        Next cell
        ws.Rows("1:7").Delete Shift:=xlUp
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For j = 1 To lastRow
            If ws.Cells(j, 1).Value = "PL/SQ" Then
                RowCount = j
                Exit For
            End If
        Next j
        If RowCount > 0 Then
            Set deleteRange = Nothing
            For j = 2 To RowCount
                If ws.Cells(j, 1).Value <> keyword Then
                    If deleteRange Is Nothing Then
                        Set deleteRange = ws.Rows(j)
                    Else
                        Set deleteRange = Union(deleteRange, ws.Rows(j))
                    End If
                End If
            Next j
            If Not deleteRange Is Nothing Then
                deleteRange.Delete Shift:=xlUp '
            End If
        End If
        ws.Columns("A:Q").AutoFit
    Next i
End Sub
Sub 合併指定工作表到新活頁簿()
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim wb As Workbook
    Dim sourceWorkbook As Workbook
    Dim sheetNames As Variant
    Dim sheetName As Variant
    Dim headerAdded As Boolean
    Dim i As Long
    
    Set wb = Workbooks.Add
    Set wsNew = wb.Sheets(1)
    wsNew.Name = "合併結果"
    Set sourceWorkbook = ThisWorkbook
    sheetNames = Array("匯入資料_ARG", "匯入資料_ARD", "匯入資料_AMG", "匯入資料_AM2", "匯入資料_AM1", "匯入資料_AEV", "匯入資料_ADG")
    headerAdded = False
    For Each sheetName In sheetNames
        On Error Resume Next
        Set ws = sourceWorkbook.Sheets(sheetName)
        On Error GoTo 0

        If Not ws Is Nothing Then
            If Not headerAdded Then
                ws.Rows(1).Copy wsNew.Rows(1)
                headerAdded = True
            End If
            lastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row + 1
            ws.Range("A2", ws.Cells(ws.Rows.Count, ws.UsedRange.Columns.Count).End(xlUp)).Copy _
                wsNew.Cells(lastRow, 1)
        End If
        Set ws = Nothing
    Next sheetName
    wsNew.Columns.AutoFit
    wb.SaveAs Filename:="G:\效率\彙整.xlsx"
    wb.Close SaveChanges:=True
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next
        Set ws = sourceWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        Set ws = Nothing
    Next i
End Sub
Sub CopyEmptyRowsToNewFileAndDelete()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim newWb As Workbook
    Dim lastRow As Long
    Dim destRow As Long
    Dim i As Long
    Dim sourceFile As String
    Dim deleteRow As Long
    Dim titlesToGroup As Variant
    Dim title As String
    Dim col As Long
    Dim titleRow As Long
    Dim j As Long
    Dim FormulR As String
    On Error GoTo ErrorHandler
    
    sourceFile = "G:\效率\彙整.xlsx"
    Set wb = Workbooks.Open(sourceFile)
    On Error Resume Next
    Set wsSource = wb.Sheets("合併結果")
    On Error GoTo ErrorHandler
    If wsSource Is Nothing Then
        MsgBox "工作表 '合併結果' 未找到！", vbCritical
        wb.Close False
        Exit Sub
    End If
    Set newWb = Workbooks.Add
    Set wsDest = newWb.Sheets(1)
    wsDest.Name = "C工時"
    wsSource.Rows(1).Copy
    wsDest.Rows(1).PasteSpecial Paste:=xlPasteValues
    lastRow = wsSource.Cells(wsSource.Rows.Count, "E").End(xlUp).Row
    destRow = 2
    For i = 2 To lastRow
        If IsEmpty(wsSource.Cells(i, 5).Value) Then
            wsSource.Rows(i).Copy
            wsDest.Rows(destRow).PasteSpecial Paste:=xlPasteValues
            destRow = destRow + 1
        End If
    Next i
    
    newWb.SaveAs "G:\效率\C工時.xlsx"
    newWb.Close
    
    For i = lastRow To 2 Step -1
        If IsEmpty(wsSource.Cells(i, 5).Value) Then
            wsSource.Rows(i).Delete
        End If
    Next i
    
    wsSource.Range("R1").Value = "工時"
    wsSource.Range("S1").Value = "繳庫工時"
    wsSource.Range("T1").Value = "實際班別"
    wsSource.Range("U1").Value = "判別"
    wsSource.Range("V1").Value = "備註"
    targetRow = wsSource.Cells(wsSource.Rows.Count, "F").End(xlUp).Row
    
    formulaR = "=IFERROR(INDEX('G:\資源用料表\資料處理專用\[AG.xlsx]Sheet1'!J:J,MATCH(F2,'G:\資源用料表\資料處理專用\[AG.xlsx]Sheet1'!F:F,0)),0)"
    wsSource.Range("R2").Formula = formulaR
    wsSource.Range("R2:R" & targetRow - 1).FillDown
    
    wsSource.Range("S2").Formula = "=R2*G2"
    wsSource.Range("S2:S" & targetRow - 1).FillDown
    
    formulaT = "=IFERROR(INDEX('G:\人力報表\[@R12資源清單-各單位彙總.xlsx]資源清單彙總'!H:H,MATCH(F2,'G:\人力報表\[@R12資源清單-各單位彙總.xlsx]資源清單彙總'!A:A,0)),0)"
    wsSource.Range("T2").Formula = formulaT
    wsSource.Range("T2:T" & targetRow - 1).FillDown
    
    wsSource.Range("U2").Formula = "=IF(LEFT(F2,2)='換模','OK',IF(LEFT(S2,4)='繳庫為0','OK',IF(ISBLANK(V2),IF(AND(ABS(H2)=ABS(ROUND((R2),4)),ABS(J2)=ABS(ROUND((S2),2))),'OK',IF(ABS(H2)<>ABS(ROUND((R2),4)),IF(ABS(J2)<>ABS(ROUND((S2),2)),'工時與繳庫錯誤','工時錯誤'),'繳庫錯誤')),'OK')))"
    wsSource.Range("U2:U" & targetRow - 1).FillDown
    
    wsSource.Columns.AutoFit
    titlesToGroup = Array("低減工時", "外包內製", "間接支援", "C工時", "修理工時", "試作工時", "換模換線")
    titleRow = 1
    
    For j = LBound(titlesToGroup) To UBound(titlesToGroup)
        title = titlesToGroup(j)
        col = 0
        On Error Resume Next
        col = Application.Match(title, wsSource.Rows(titleRow), 0)
        On Error GoTo ErrorHandler
        If col > 0 Then
            wsSource.Columns(col).Group
            wsSource.Columns(col).Hidden = True
        End If
    Next j
    wb.Save
    wb.Close False
    Exit Sub
ErrorHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
    If Not wb Is Nothing Then wb.Close False
End Sub

Sub 樞紐()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim wb As Workbook
    Dim filePath As String
    Dim pivotSheet As Worksheet
    Dim pivotTable As pivotTable
    
    filePath = "G:\效率\彙整.xlsx"
    On Error GoTo ErrorHandler
    Set wb = Workbooks.Open(filePath)
    On Error Resume Next
    Set ws = wb.Sheets("合併結果")
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        MsgBox "工作表 '合併結果' 未找到!", vbCritical
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set pivotSheet = Sheets.Add(after:=Sheets(Sheets.Count))
    pivotSheet.Name = "Pivot"
    Set pivotTable = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        ws.Range("A1:V" & lastRow)).CreatePivotTable(TableDestination:=pivotSheet.Range("A3"), TableName:="PivotTable")
        
    On Error Resume Next
    With pivotTable.PivotFields("組織")
        If Err.Number <> 0 Then
            MsgBox "欄位 '組織' 不存在!", vbCritical
            Exit Sub
        End If
        .Orientation = xlRowField
        .Position = 1
        .Caption = "部門"
    End With
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    With pivotTable.PivotFields("實際班別")
        If Err.Number <> 0 Then
            MsgBox "欄位 '實際班別' 不存在!", vbCritical
            Exit Sub
        End If
        .Orientation = xlRowField
        .Position = 2
        .Caption = "實際班別"
    End With
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    pivotTable.AddDataField pivotTable.PivotFields("繳庫"), "加總：繳庫", xlSum
    If Err.Number <> 0 Then
        MsgBox "欄位 '繳庫' 不存在!", vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    pivotTable.RowAxisLayout xlTabularRow
    wb.Save
    wb.Close
    MsgBox "完成", vbInformation
    Exit Sub
ErrorHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical
End Sub


    
    

