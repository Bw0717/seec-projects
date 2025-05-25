VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim filePath As String
    Dim wb As Workbook, ws As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, i As Long, j As Long, k As Long
    Dim outputRow As Long, typeCounter As Long
    Dim startRow As Long, mergeEnd As Long, currentRow As Long, currentCounter As Long
    Dim typeRowCount As Long, lastDataRow As Long

    Dim dataArr As Variant
    Dim groupDict As Object, typeMap As Object, partResDict As Object, partSummaryDict As Object, typeReplaceDict As Object
    Dim keyArr As Variant, partKey As Variant, keyList As Variant
    Dim partNo As String, partSummary As String, res As String, note As String, srcPart As String
    Dim comboKey As String, typeName As String
    Dim oldType As String, newType As String
    Dim currentType As String, currentPart As String

    filePath = Application.GetOpenFilename("Excel 檔案 (*.xlsx), *.xlsx", , "選擇來源資料")
    If filePath = "False" Or filePath = "" Then Exit Sub

    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Sheets(1)
    lastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    dataArr = ws.Range("A2:E" & lastRow).Value

    Set groupDict = CreateObject("Scripting.Dictionary")
    Set typeMap = CreateObject("Scripting.Dictionary")
    Set partResDict = CreateObject("Scripting.Dictionary")
    Set partSummaryDict = CreateObject("Scripting.Dictionary")
    typeCounter = 1

    For i = 1 To UBound(dataArr)
        partNo = Trim(CStr(dataArr(i, 1)))
        partSummary = Trim(CStr(dataArr(i, 2)))
        res = Trim(CStr(dataArr(i, 3)))
        note = Trim(CStr(dataArr(i, 4)))
        srcPart = Trim(CStr(dataArr(i, 5)))


        
        If (IsAlphaDigitAlpha3(partNo) And partNo = srcPart) Or (Not IsAlphaDigitAlpha3(partNo)) Then
            comboKey = res & "|" & note
            If Not groupDict.exists(partNo) Then Set groupDict(partNo) = CreateObject("Scripting.Dictionary")
            groupDict(partNo)(comboKey) = True
            If Not partResDict.exists(partNo) Then Set partResDict(partNo) = CreateObject("Scripting.Dictionary")
            partResDict(partNo)(comboKey) = True
            If Not partSummaryDict.exists(partNo) Then partSummaryDict.Add partNo, partSummary
        End If
SkipRow:
    Next i

    keyArr = groupDict.keys
    For i = 0 To UBound(keyArr)
        partKey = keyArr(i)
        comboKey = Join(SortKeys(groupDict(partKey)), "||")
        If Not typeMap.exists(comboKey) Then
            typeMap.Add comboKey, "TEMP_TYPE_" & Format(typeCounter, "000")
            typeCounter = typeCounter + 1
        End If
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("分類結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.Name = "分類結果"
    With wsOut
        .Cells(1, 1).Value = "TYPE"
        .Cells(1, 2).Value = "生產件號"
        .Cells(1, 3).Value = "摘要"
        .Cells(1, 4).Value = "資源"
        .Cells(1, 5).Value = "資源備註"
    End With

    outputRow = 2
    keyArr = partResDict.keys
    For i = 0 To UBound(keyArr)
        partKey = keyArr(i)
        comboKey = Join(SortKeys(groupDict(partKey)), "||")
        typeName = typeMap(comboKey)
        keyList = partResDict(partKey).keys

        For k = LBound(keyList) To UBound(keyList)
            comboKey = keyList(k)
            res = Split(comboKey, "|")(0)
            note = Split(comboKey, "|")(1)
            wsOut.Cells(outputRow, 1).Value = typeName
            wsOut.Cells(outputRow, 2).Value = partKey
            wsOut.Cells(outputRow, 3).Value = partSummaryDict(partKey)
            wsOut.Cells(outputRow, 4).Value = res
            wsOut.Cells(outputRow, 5).Value = note
            outputRow = outputRow + 1
        Next k
    Next i

    Set typeReplaceDict = CreateObject("Scripting.Dictionary")
    currentCounter = 1
    lastDataRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    For currentRow = 2 To lastDataRow
        oldType = wsOut.Cells(currentRow, 1).Value
        If Not typeReplaceDict.exists(oldType) Then
            newType = "TYPE" & Format(currentCounter, "00")
            typeReplaceDict.Add oldType, newType
            currentCounter = currentCounter + 1
        End If
        wsOut.Cells(currentRow, 1).Value = typeReplaceDict(oldType)
    Next currentRow

    wsOut.Columns.AutoFit
    MsgBox "分類完成", vbInformation
    wb.Close SaveChanges:=False
End Sub

Function IsAlphaDigitAlpha3(text As String) As Boolean
    If Len(text) = 3 Then
        If Mid(text, 1, 1) Like "[A-Za-z]" And _
           Mid(text, 2, 1) Like "#" And _
           Mid(text, 3, 1) Like "[A-Za-z]" Then
            IsAlphaDigitAlpha3 = True
        Else
            IsAlphaDigitAlpha3 = False
        End If
    Else
        IsAlphaDigitAlpha3 = False
    End If
End Function

Function SortKeys(dict) As Variant
    Dim arr() As String
    Dim i As Long, j As Long
    Dim tmp As String
    Dim keys As Variant
    keys = dict.keys
    ReDim arr(0 To UBound(keys))
    For i = 0 To UBound(keys)
        arr(i) = keys(i)
    Next i
    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next j
    Next i
    SortKeys = arr
End Function

Sub ff()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim fDialog As FileDialog
    Dim fso As Object
    Dim folder As Object
    Dim searchDict As Object
    Dim matchedFiles As Object
    Dim lastRow As Long
    Dim i As Long, key As Variant
    Dim foundPath As String
    Dim rowsToDelete As Collection
    Set rowsToDelete = New Collection

    Set ws = ThisWorkbook.Sheets("索引清單")

    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "請選擇要搜尋的資料夾"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set searchDict = CreateObject("Scripting.Dictionary")
    Set matchedFiles = CreateObject("Scripting.Dictionary")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            searchDict(ws.Cells(i, 1).Value) = i
            matchedFiles(ws.Cells(i, 1).Value) = Array(#1/1/1970#, "")
        End If
    Next i

    Call 遞迴搜尋(folderPath, searchDict, matchedFiles)

    For Each key In matchedFiles.keys
        i = searchDict(key)
        foundPath = matchedFiles(key)(1)
        If foundPath <> "" Then
            ws.Cells(i, 2).Value = foundPath
        Else
            ws.Cells(i, 2).Value = "未找到"
            rowsToDelete.Add i
        End If
    Next key

    If rowsToDelete.Count > 0 Then
        Dim idx As Long
        For idx = rowsToDelete.Count To 1 Step -1
            ws.Rows(rowsToDelete(idx)).Delete
        Next idx
    End If

    MsgBox "搜尋完成", vbInformation
End Sub

Sub 遞迴搜尋(folderPath As String, searchDict As Object, matchedFiles As Object)
    Dim fso As Object, folder As Object, subFolder As Object
    Dim file As Object, fileName As String
    Dim key As Variant
    Dim fileTime As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    For Each file In folder.files
        fileName = fso.GetBaseName(file.Name)
        If (LCase(fso.GetExtensionName(file.Name)) = "pdf" Or LCase(fso.GetExtensionName(file.Name)) = "xlsx") And InStr(file.Name, "~") = 0 Then
            For Each key In searchDict.keys
                If InStr(fileName, key) > 0 Then
                    fileTime = file.DateLastModified
                    If fileTime > matchedFiles(key)(0) Then
                        matchedFiles(key) = Array(fileTime, file.Path)
                    End If
                End If
            Next key
        End If
    Next file

    For Each subFolder In folder.SubFolders
        Call 遞迴搜尋(subFolder.Path, searchDict, matchedFiles)
    Next subFolder
End Sub

Private Sub CommandButton2_Click()
    Dim fDialog As FileDialog
    Dim selectedFile As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim savePath As String
    Dim fileName As String
    Dim fileExt As String
    Dim newName As String

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .Title = "請選擇要處理的 Excel 檔案"
        .Filters.Clear
        .Filters.Add "Excel 檔案", "*.xlsx"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        selectedFile = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False


    Set wb = Workbooks.Open(fileName:=selectedFile)

    If wb.Sheets.Count >= 1 Then wb.Sheets(1).Delete


    Set ws = wb.Sheets(1)
    ws.Columns("C").Delete
    ws.Columns("C").Delete

    ws.Columns("L").Copy
    ws.Columns("D").PasteSpecial xlPasteValues
    Application.CutCopyMode = False

    fileName = Mid(selectedFile, InStrRev(selectedFile, "\") + 1)
    fileName = Left(fileName, InStrRev(fileName, ".") - 1)
    savePath = Left(selectedFile, InStrRev(selectedFile, "\")) & fileName & "_完成.xlsx"

    wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook ' .xlsx
    wb.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "處理完成，檔案已另存為：" & vbCrLf & savePath, vbInformation
End Sub

Private Sub CommandButton3_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dict As Object
    Set ws = ThisWorkbook.Sheets("索引清單")
    Set dict = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    

    For i = lastRow To 2 Step -1
        Dim val As String
        val = Trim(ws.Cells(i, "A").Value)
        
        If val <> "" Then
            If dict.exists(val) Then
                ws.Rows(i).Delete
            Else
                dict.Add val, True
            End If
        End If
    Next i
    
    Call ff
End Sub


