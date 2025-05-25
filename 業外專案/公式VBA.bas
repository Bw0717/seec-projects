Attribute VB_Name = "Module1"
Sub InsertFormulaToSecondWorkbook_Stable()

    Dim sourcePath As String, targetPath As String
    Dim sourceName As String, targetName As String
    Dim sourceExt As String, targetExt As String
    Dim newSourcePath As String, newTargetPath As String
    Dim sourceWB As Workbook, targetWB As Workbook
    Dim targetWS As Worksheet
    Dim wb As Workbook
    Dim sourceAlreadyOpen As Boolean, targetAlreadyOpen As Boolean
    Dim sheetName As String
    Dim fso As Object
    Dim sourceFolder As String
    Dim fullSourceRef As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    targetPath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "選擇【出貨報告】檔案")
    If targetPath = "False" Then Exit Sub
    targetExt = LCase(fso.GetExtensionName(targetPath))
    targetName = fso.GetFileName(targetPath)
    targetAlreadyOpen = False
    For Each wb In Application.Workbooks
        If wb.FullName = targetPath Then
            Set targetWB = wb
            targetAlreadyOpen = True
            Exit For
        End If
    Next wb
    If targetExt = "xls" Then
        If Not targetAlreadyOpen Then
            Set targetWB = Workbooks.Open(targetPath)
        End If
        newTargetPath = Left(targetPath, InStrRev(targetPath, ".")) & "xlsx"
        targetWB.SaveAs Filename:=newTargetPath, FileFormat:=xlOpenXMLWorkbook
        targetWB.Close SaveChanges:=False
        MsgBox "目標檔已轉為 .xlsx：" & vbCrLf & newTargetPath, vbInformation
        
        targetAlreadyOpen = False
        For Each wb In Application.Workbooks
            If wb.FullName = newTargetPath Then
                Set targetWB = wb
                targetAlreadyOpen = True
                Exit For
            End If
        Next wb
        If Not targetAlreadyOpen Then
            Set targetWB = Workbooks.Open(newTargetPath)
        End If
        
        targetPath = newTargetPath
        targetName = fso.GetFileName(newTargetPath)
    ElseIf Not targetAlreadyOpen Then
        Set targetWB = Workbooks.Open(targetPath)
    End If
    
    sourcePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "選擇【製造清單】檔案")
    If sourcePath = "False" Then Exit Sub
    sourceExt = LCase(fso.GetExtensionName(sourcePath))
    sourceName = fso.GetFileName(sourcePath)
    sourceAlreadyOpen = False
    
    For Each wb In Application.Workbooks
        If wb.FullName = sourcePath Then
            Set sourceWB = wb
            sourceAlreadyOpen = True
            Exit For
        End If
    Next wb
    
    If sourceExt = "xls" Then
        If Not sourceAlreadyOpen Then
            Set sourceWB = Workbooks.Open(sourcePath)
        End If
        
        newSourcePath = Left(sourcePath, InStrRev(sourcePath, ".")) & "xlsx"
        sourceWB.SaveAs Filename:=newSourcePath, FileFormat:=xlOpenXMLWorkbook
        sourceWB.Close SaveChanges:=False
        
        For Each wb In Application.Workbooks
            If wb.FullName = newSourcePath Then
                wb.Close SaveChanges:=False
                Exit For
            End If
        Next wb
        
        MsgBox "來源檔已轉為 .xlsx：" & vbCrLf & newSourcePath, vbInformation
        sourcePath = newSourcePath
        sourceName = fso.GetFileName(newSourcePath)
    End If

    sourceFolder = Left(sourcePath, InStrRev(sourcePath, "\"))
    fullSourceRef = "'" & sourceFolder & "[" & sourceName & "]T值'!"
    
    sheetName = "T值"
    On Error Resume Next
    Set targetWS = targetWB.Sheets(sheetName)
    On Error GoTo 0
    If targetWS Is Nothing Then
        MsgBox "在檔案 [" & targetName & "] 中找不到工作表 'T值'！", vbExclamation
        If Not targetAlreadyOpen Then targetWB.Close SaveChanges:=False
        Exit Sub
    End If

    With targetWS
        .Range("C7").Formula = "=IFERROR(INDEX(" & fullSourceRef & "$C:$C,MATCH(C6," & fullSourceRef & "$C:$C,0)+1),""沒有下一個值"")"
        .Range("C8").Formula = "=IFERROR(INDEX(" & fullSourceRef & "$C:$C,MATCH(C6," & fullSourceRef & "$C:$C,0)+2),""沒有下一個值"")"
        .Range("D6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,2,FALSE)"
        .Range("E6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,3,FALSE)"
        .Range("F6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,4,FALSE)"
        .Range("G6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,5,FALSE)"
        .Range("H6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,6,FALSE)"
        .Range("I6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,7,FALSE)"
        .Range("J6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,8,FALSE)"
        .Range("K6").Formula = "=VLOOKUP($C6," & fullSourceRef & "$C:$K,9,FALSE)"
        .Range("L6").Formula = "=MAX($D6:$K6)"
        .Range("M6").Formula = "=MIN($D6:$K6)"
        .Range("N6").Formula = "=AVERAGE($D6:$K6)"
        .Range("O6").Formula = "=$L6-$M6"
        
        .Range("D6:O6").Copy
        .Range("D7").PasteSpecial
        .Range("D8").PasteSpecial

    End With

    Application.CutCopyMode = False
    MsgBox "公式寫入完成", vbInformation

    If Not targetAlreadyOpen Then
        targetWB.Save
        targetWB.Close
    End If

End Sub

