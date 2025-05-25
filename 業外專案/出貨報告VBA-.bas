Attribute VB_Name = "Module1"

Sub 更新1()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[1.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[1.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新2()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[2.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[2.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新3()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[3.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[3.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新4()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[4.xlsm]輸入實測值'!R29C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[4.xlsm]輸入實測值'!R29C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新5()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[5.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[5.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新6()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[6.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[6.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新7()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[7.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[7.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新UB1()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[ub1.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[ub1.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新UB2()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[ub2.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[ub2.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub 更新UB3()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[ub3.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[ub3.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub
Sub AlignButtons_Horizontally()
    Dim ws As Worksheet
    Dim btnNames As Variant
    Dim btn As Shape
    Dim i As Integer
    Dim leftPos As Double
    Dim spacing As Double
    
    Set ws = ActiveSheet
    btnNames = Array("Oval 1", "Oval 3", "Oval 4", "Oval 16", "Oval 22", "橢圓 7")
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    
    btnNames = Array("Oval 5", "Oval 6", "Oval 17", "Oval 18", "Oval 23", "橢圓 6") ' 按鈕名稱列表
    
    leftPos = 100
    spacing = 120

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    btnNames = Array("Oval 9", "Oval 11", "Oval 12", "Oval 24", "Oval 25", "Oval 15") ' 按鈕名稱列表
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i



    btnNames = Array("Oval 13", "Oval 14", "Oval 19", "Oval 20", "Oval 26", "橢圓 1") ' 按鈕名稱列表
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    
    
    btnNames = Array("Oval 30", "DAEWON", "Text Box 882", "Text Box 27", "圓角矩形 13", "chipbond-1") ' 按鈕名稱列表
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    btnNames = Array("chipbondCM", "chipbondCM1") ' 按鈕名稱列表
    
    leftPos = 580
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66 + 66 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
End Sub

Sub LockAllButtonsAndProtectSheet()
    Dim ws As Worksheet
    Dim shp As Shape
    
    Set ws = ActiveSheet
    

    ws.Unprotect "123456"
    

    For Each shp In ws.Shapes
        shp.Placement = xlFreeFloating
        shp.Locked = True
    Next shp


    ws.Protect Password:="123456", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    
    MsgBox "所有按鈕已鎖定，無法移動！", vbInformation, "成功"
End Sub

Sub 更新11()
Application.ScreenUpdating = False
    x = InputBox("檔案名稱")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[1-1.xlsm]輸入實測值'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[1-1.xlsm]輸入實測值'!R34C2:R65536C19," & i & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub

Sub 自動帶入()

Application.ScreenUpdating = False
    y = InputBox("輸入料號")
    x = InputBox("檔案名稱")
Set wb = Workbooks(x)
Set ws = wb.Sheets(1)
Set formulaRange = ws.Range("M35:M46")
Set formulaRangel = ws.Range("N35:N46")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = _
    "=VLOOKUP(R35C2,'[" & y & ".xlsx]工作表1'!R4C1:R65536C23,3,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 12
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[" & y & ".xlsx]工作表1'!R4C1:R65536C23," & i + 1 & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i

    formulaText = "=IF(B35="""","""",VLOOKUP(B35,'[" & y & ".xlsx]工作表1'!$A$4:$W$32474,15,FALSE))"
    formulaRange.Formula = formulaText
    Application.Calculate
    formulaRange.Value = formulaRange.Value
    
    formulaText = "=IF(B35="""","""",VLOOKUP(B35,'[" & y & ".xlsx]工作表1'!$A$4:$W$32474,17,FALSE))"
    formulaRangel.Formula = formulaText
    Application.Calculate
    formulaRangel.Value = formulaRangel.Value
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub

