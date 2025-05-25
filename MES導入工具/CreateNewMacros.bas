Attribute VB_Name = "CreateNewMacros"
Sub ImportAndConvertTxtToExcel()
    Dim FilePath As String
    Dim SavePath As String
    Dim fileNum As Integer
    Dim lineData As String
    Dim dataArray As Variant
    Dim rowNum As Integer
    Dim colNum As Integer
    Dim ws As Worksheet
    Dim newWb As Workbook
    Dim LastRow As Long, LastCol As Long
    Dim r As Long, c As Long


    FilePath = Application.GetOpenFilename("Text Files (*.tsv), *.tsv", , "選擇要匯入的 tsv 檔案")
    If FilePath = "False" Then
        MsgBox "未選擇檔案，程式終止。", vbExclamation
        Exit Sub
    End If


    Set newWb = Workbooks.Add

    Set ws = newWb.Sheets(1)
    ws.Name = "ImportedData"

    With ws.QueryTables.Add(Connection:="TEXT;" & FilePath, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With

    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column


    For r = 1 To LastRow
        For c = 1 To LastCol
            ws.Cells(r, c).Value = Replace(Trim(ws.Cells(r, c).Value), "-", "")
        Next c
    Next r
    
    For r = 1 To LastRow
        For c = 1 To LastCol
            ws.Cells(r, c).Value = Replace(Trim(ws.Cells(r, c).Value), " ", "")
        Next c
    Next r
    
    
    
    Set Rng = ws.Rows(1)
    
    deleteCols = Array("版次", "狀態", "工程料號", "料號序號", "作業序號", "替代結構", "工程用料表", "附註", "基準", "計畫百分比", "良品率", "小計數量", "有效性控制", "起始", "終止", "起始日期", "終止日期", "停用", "已導入", "工程變更單", "供給型態", "倉庫", "儲位", "計算成本", "單位成本", "小計數量", "小計成本", "作業序號", "製造", "延緩", "累積製造", "累積總計", "選擇性", "互斥", "ATP", "最小數量", "最大數量", "銷售訂單基準", "可出貨", "納入出貨文件", "出貨所需", "收入所需")
    For i = Rng.Columns.Count To 1 Step -1
        For Each colName In deleteCols
            If ws.Cells(1, i).Value = colName Then
                ws.Columns(i).Delete
                Exit For
            End If
        Next colName
    Next i
            
           
 
    ws.Columns.AutoFit

    SavePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="另存新檔")
    If SavePath <> "False" Then
        newWb.SaveAs Filename:=SavePath, FileFormat:=51
        MsgBox "轉檔完成！", vbInformation
    Else
        MsgBox "轉檔取消。", vbExclamation
    End If
End Sub

