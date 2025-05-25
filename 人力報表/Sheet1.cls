VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ToggleButton1_Click()
    Dim xmlFilePaths As Variant
    Dim xmlFilePath As Variant
    Dim xlsxFilePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastColumn As Long
    Dim headerRange As Range
    Dim i As Long
    Dim headers As Variant
    Dim newHeaders As Variant
    Dim newColumnIndex As Long
    Dim columnFound As Boolean
    Dim fileName As String
    Dim resourceColumn As Long
    Dim standardCode As String
    Dim tempWs As Worksheet
    Dim totalAssemblyColumn As Long
    Dim totalAssemblyLastRow As Long


    xmlFilePaths = Array( _
        "C:\Users\14574\Desktop\資料處理專用\AEV.XML", _
        "C:\Users\14574\Desktop\資料處理專用\ADG.XML", _
        "C:\Users\14574\Desktop\資料處理專用\AM1.XML", _
        "C:\Users\14574\Desktop\資料處理專用\AM2.XML", _
        "C:\Users\14574\Desktop\資料處理專用\AMG.XML", _
        "C:\Users\14574\Desktop\資料處理專用\ARG.XML", _
        "C:\Users\14574\Desktop\資料處理專用\ARD.XML" _
    )
    For Each xmlFilePath In xmlFilePaths
        xlsxFilePath = Replace(xmlFilePath, ".XML", ".xlsx")
        Set wb = Workbooks.OpenXML(fileName:=xmlFilePath, LoadOption:=xlXmlLoadImportToList)
        Set ws = wb.Worksheets(1)
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headers = Array("C_FLEXITEM", "C_ITEM_DESCRIPTION", "C_DEPARTMENT_CODE", "C_RESOURCE_CODE", "C_FLEXDATA", "C_SETBACK_DAYS", "C_ASSEMBLY_USAGE", "C_REQUIRED_HOURS", "C_RESOURCE_UNITS", "C_TOTAL_HOURS")
        newHeaders = Array("總成件號", "總成摘要", "班別", "資源", "來源件號", "倒推日期", "使用單位", "需求時數", "資源單位", "總計時數")
        Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastColumn))
        For newColumnIndex = LBound(headers) To UBound(headers)
            columnFound = False
            For i = 1 To lastColumn
                If ws.Cells(1, i).Value = headers(newColumnIndex) Then
                    ws.Cells(1, i).Value = newHeaders(newColumnIndex)
                    columnFound = True
                    Exit For
                End If
            Next i
            If Not columnFound Then Exit Sub
        Next newColumnIndex
        Set tempWs = wb.Worksheets.Add(after:=ws)
        For newColumnIndex = LBound(newHeaders) To UBound(newHeaders)
            For i = 1 To lastColumn
                If ws.Cells(1, i).Value = newHeaders(newColumnIndex) Then
                   
                    tempWs.Cells(1, newColumnIndex + 1).Value = newHeaders(newColumnIndex)
                   
                    ws.Columns(i).Copy Destination:=tempWs.Columns(newColumnIndex + 1)
                    Exit For
                End If
            Next i
        Next newColumnIndex
        ws.Cells.Clear
        tempWs.UsedRange.Copy Destination:=ws.Cells(1, 1)
        fileName = Mid(xmlFilePath, InStrRev(xmlFilePath, "\") + 1)
        fileName = Replace(fileName, ".XML", "")
        ws.Columns(1).Insert Shift:=xlToRight
        ws.Cells(1, 1).Value = "部門"
        totalAssemblyColumn = Application.Match("總成件號", ws.Rows(1), 0)
        If Not IsError(totalAssemblyColumn) Then
            totalAssemblyLastRow = ws.Cells(ws.Rows.Count, totalAssemblyColumn).End(xlUp).Row
            ws.Range(ws.Cells(2, 1), ws.Cells(totalAssemblyLastRow, 1)).Value = fileName
        End If
        resourceColumn = Application.Match("資源", ws.Rows(1), 0)
        If Not IsError(resourceColumn) Then
            ws.Columns(resourceColumn).Insert Shift:=xlToRight
            ws.Cells(1, resourceColumn).Value = "標準代號"
            For i = 2 To ws.Cells(ws.Rows.Count, resourceColumn + 1).End(xlUp).Row
                If InStr(ws.Cells(i, resourceColumn + 1).Value, "換模") > 0 Then
                    ws.Cells(i, resourceColumn).Value = "-"
                Else
                    Dim resourceValue As String
                    resourceValue = ws.Cells(i, resourceColumn + 1).Value
                    Dim totalLength As Integer
                    totalLength = Len(resourceValue)
                    If totalLength = 10 Then
                        standardCode = Mid(resourceValue, 2, 1) & _
                                       Mid(resourceValue, 4, 1) & _
                                       Mid(resourceValue, 6, 1) & _
                                       Mid(resourceValue, 7, 1)
                    ElseIf totalLength = 9 Then
                        standardCode = Mid(resourceValue, 2, 1) & _
                                       Mid(resourceValue, 3, 1) & _
                                       Mid(resourceValue, 5, 1) & _
                                       Mid(resourceValue, 6, 1)
                    Else
                        standardCode = ""
                    End If
            
                    ws.Cells(i, resourceColumn).Value = standardCode
                End If
            Next i
        End If
        ws.Cells.HorizontalAlignment = xlHAlignLeft
        Application.DisplayAlerts = False
        tempWs.Delete
        Application.DisplayAlerts = True
        wb.SaveAs fileName:=xlsxFilePath, FileFormat:=xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
    Next xmlFilePath
    MsgBox "成功"
End Sub

Private Sub ToggleButton2_Click()
    Call ConsolidateXLSXFiles
    Call UpdateColumnI
    Call AG樞紐
End Sub
Sub ConsolidateXLSXFiles()
    Dim sourceFilePaths As Variant
    Dim sourceFilePath As Variant
    Dim targetWb As Workbook
    Dim sourceWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    
    sourceFilePaths = Array( _
        "C:\Users\14574\Desktop\資料處理專用\AEV.xlsx", _
        "C:\Users\14574\Desktop\資料處理專用\ADG.xlsx", _
        "C:\Users\14574\Desktop\資料處理專用\AM1.xlsx", _
        "C:\Users\14574\Desktop\資料處理專用\AM2.xlsx", _
        "C:\Users\14574\Desktop\資料處理專用\AMG.xlsx", _
        "C:\Users\14574\Desktop\資料處理專用\ARG.xlsx", _
        "C:\Users\14574\Desktop\資料處理專用\ARD.xlsx" _
    )
    Set targetWb = Workbooks.Add
    Set targetWs = targetWb.Worksheets(1)
    targetRow = 1
    For Each sourceFilePath In sourceFilePaths
        On Error Resume Next
        Set sourceWb = Workbooks.Open(sourceFilePath)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        Set sourceWs = sourceWb.Worksheets(1)
        lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 1 Then
            If targetRow = 1 Then
                sourceWs.Range("A1:L" & lastRow).Copy Destination:=targetWs.Cells(targetRow, 1)
                targetRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row + 1
            Else
                If lastRow > 1 Then
                    sourceWs.Range("A2:L" & lastRow).Copy Destination:=targetWs.Cells(targetRow, 1)
                    targetRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row + 1
                End If
            End If
        End If
        sourceWb.Close SaveChanges:=False
NextFile:
    Next sourceFilePath
    targetWs.Range("N1").Value = "資源備註"
    
  
    targetWs.Range("M1").Value = "實際班別"
    
   
    targetWs.Range("O1").Value = "預測量(次1月)"
    
    
    targetWs.Range("P1").Value = "工時需求(次1月)"
    
   
    targetWs.Range("Q1").Value = "預測量(次2月)"
    
  
    targetWs.Range("R1").Value = "工時需求(次2月)"
    
  
    targetWs.Range("S1").Value = "預測量(次3月)"
    
 
    targetWs.Range("T1").Value = "工時需求(次3月)"
    
   
    targetWs.Range("U1").Value = "委託數量"
    
   
    targetWs.Range("V1").Value = "委託工時"
    Formula = "=IF(ISERROR(VLOOKUP(B2,'[建議數量-人力計算.xlsx]Pivot'!$A:$D,2,FALSE)),0,VLOOKUP(B2,'[建議數量-人力計算.xlsx]Pivot'!$A:$D,2,FALSE))"
    targetWs.Range("O2").Formula = Formula
    targetWs.Range("O2:O" & targetRow - 1).FillDown
    
 
    formulaM = "=VLOOKUP(TRIM(F2),'\\172.16.20.111\af-am製造管理處\AAD製造管理部\AAPS能率課\工時人力管理組\1.標準工時\@資源代碼資料庫\[@R12資源清單-各單位彙總.xlsx]資源清單彙總'!$A:$H,8,FALSE)"
    targetWs.Range("M2").Formula = formulaM
    targetWs.Range("M2:M" & targetRow - 1).FillDown
    
 
    formulaP = "=I2*J2*O2"
    targetWs.Range("P2").Formula = formulaP
    targetWs.Range("P2:P" & targetRow - 1).FillDown
    
  
    formulaQ = "=IF(ISERROR(VLOOKUP(B2,'[建議數量-人力計算.xlsx]Pivot'!$A:$D,3,FALSE)),0,VLOOKUP(B2,'[建議數量-人力計算.xlsx]Pivot'!$A:$D,3,FALSE))"
    targetWs.Range("Q2").Formula = formulaQ
    targetWs.Range("Q2:Q" & targetRow - 1).FillDown
    
 
    formulaR = "=I2*J2*Q2"
    targetWs.Range("R2").Formula = formulaR
    targetWs.Range("R2:R" & targetRow - 1).FillDown
    

    formulaS = "=IF(ISERROR(VLOOKUP(B2,'[建議數量-人力計算.xlsx]Pivot'!$A:$D,4,FALSE)),0,VLOOKUP(B2,'[建議數量-人力計算.xlsx]Pivot'!$A:$D,4,FALSE))"
    targetWs.Range("S2").Formula = formulaS
    targetWs.Range("S2:S" & targetRow - 1).FillDown
    

    formulaT = "=I2*J2*S2"
    targetWs.Range("T2").Formula = formulaT
    targetWs.Range("T2:T" & targetRow - 1).FillDown
    
  
    formulaU = "=IFERROR(INDEX('[部門間委託.xlsx]明細'!G:G,MATCH(B2,'[部門間委託.xlsx]明細'!D:D,0)),0)"
    targetWs.Range("U2").Formula = formulaU
    targetWs.Range("U2:U" & targetRow - 1).FillDown
    

    formulaV = "=I2*J2*U2"
    targetWs.Range("V2").Formula = formulaV
    targetWs.Range("V2:V" & targetRow - 1).FillDown
    targetWb.SaveAs fileName:="C:\Users\14574\Desktop\資料處理專用\AG.xlsx", FileFormat:=xlOpenXMLWorkbook
    targetWb.Close SaveChanges:=False
End Sub
Sub UpdateColumnI()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim wb As Workbook
    Dim filePath As String

    filePath = "C:\Users\14574\Desktop\資料處理專用\AG.xlsx"
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 5) = "AM106" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 5) = "AA383" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 6) = "AA301C" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 5) = "AA882" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 3) = "A4H" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 5) = "AM800" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 7).Value, 5) = "AM801" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    For i = 2 To lastRow
        If Left(ws.Cells(i, 4).Value, 3) = "ACF" Then
            ws.Cells(i, 9).Value = 0
        End If
    Next i
    wb.Save
    wb.Close
End Sub

Sub AG樞紐()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim wb As Workbook
    Dim filePath As String
    
    filePath = "C:\Users\14574\Desktop\資料處理專用\AG.xlsx"
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Pivot"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets("Sheet1").Range("A1:V" & lastRow)).CreatePivotTable TableDestination:= _
        Sheets("Pivot").Range("A3"), TableName:="PivotTable"
    With ActiveSheet.PivotTables("PivotTable").PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
        .Caption = "部門"
    End With
    With ActiveSheet.PivotTables("PivotTable").PivotFields("實際班別")
        .Orientation = xlRowField
        .Position = 2
        .Caption = "實際班別"
    End With
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("工時需求(次1月)"), _
        "加總：工時需求(次1月)", xlSum
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("工時需求(次2月)"), _
        "加總：工時需求(次2月)", xlSum
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("工時需求(次3月)"), _
        "加總：工時需求(次3月)", xlSum
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("委託工時"), _
        "加總：委託工時", xlSum
    ActiveSheet.PivotTables("PivotTable").RowAxisLayout xlTabularRow
    wb.Save
    wb.Close
    MsgBox "完成", vbInformation
End Sub
Private Sub ToggleButton3_Click()
    Call UpdateColumnI2
    Call Edit
End Sub
Sub UpdateColumnI2()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim wb As Workbook
    Dim filePath As String
    filePath = "C:\Users\14574\Desktop\資料處理專用\建議數量-人力計算.xlsx"
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets("預測(含安全)")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        If Left(ws.Cells(i, 2).Value, 3) = "A2X" Then
          
            ws.Cells(i, 6).Value = "AEV1"
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        If Left(ws.Cells(i, 2).Value, 5) = "A4A2X" Then
       
            ws.Cells(i, 6).Value = "AEV1"
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        If Left(ws.Cells(i, 2).Value, 2) = "AX" Then
            ws.Cells(i, 6).Value = "AEV1"
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        If Left(ws.Cells(i, 2).Value, 3) = "M9Q" Then
            ws.Cells(i, 6).Value = "AH1"
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        If Left(ws.Cells(i, 2).Value, 3) = "A2Q" Then
            ws.Cells(i, 6).Value = "AEV4"
        End If
    Next i
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    For i = 3 To lastRow
        If Left(ws.Cells(i, 2).Value, 2) = "AQ" Then
            ws.Cells(i, 6).Value = "AEV4"
        End If
    Next i
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Pivot"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Sheets("預測(含安全)").Range("B2:J" & lastRow)).CreatePivotTable TableDestination:= _
        Sheets("Pivot").Range("A3"), TableName:="PivotTable"
    With ActiveSheet.PivotTables("PivotTable").PivotFields("料號")
        .Orientation = xlRowField
        .Position = 1
        .Caption = "料號"
    End With
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("當月"), _
        "加總：當月", xlSum
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("次月"), _
        "加總：次月", xlSum
    ActiveSheet.PivotTables("PivotTable").AddDataField ActiveSheet.PivotTables("PivotTable").PivotFields("次2月"), _
        "加總：次2月", xlSum
    ActiveSheet.PivotTables("PivotTable").RowAxisLayout xlTabularRow
    wb.Save
    wb.Close
End Sub
Sub Edit()
    Dim ws As Worksheet
    Dim pt As PivotTable
    filePath = "C:\Users\14574\Desktop\資料處理專用\建議數量-人力計算.xlsx"
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets("含安全量樞紐1")
    Set pt = ws.PivotTables("樞紐分析表7")
    With pt
        .AddDataField .PivotFields("排定金額"), "加總：排定金額", xlSum
        .AddDataField .PivotFields("次月$"), "加總：次月$", xlSum
        .AddDataField .PivotFields("次2月$"), "加總：次2月$", xlSum
        .PivotFields("加總：排定金額").Orientation = xlDataField
        .PivotFields("加總：排定金額").Position = 1
        .PivotFields("加總：次月$").Orientation = xlDataField
        .PivotFields("加總：次月$").Position = 2
        .PivotFields("加總：次2月$").Orientation = xlDataField
        .PivotFields("加總：次2月$").Position = 3
    End With
    wb.Save
    wb.Close
    MsgBox "完成", vbInformation
End Sub
