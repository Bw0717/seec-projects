Attribute VB_Name = "Module1"

Sub ��s1()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[1.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[1.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��s2()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[2.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[2.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��s3()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[3.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[3.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��s4()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[4.xlsm]��J�����'!R29C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[4.xlsm]��J�����'!R29C2:R65536C19," & i & ",0)"
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
Sub ��s5()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[5.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[5.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��s6()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[6.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[6.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��s7()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[7.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[7.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��sUB1()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[ub1.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[ub1.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��sUB2()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[ub2.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[ub2.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
Sub ��sUB3()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[ub3.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[ub3.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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
    btnNames = Array("Oval 1", "Oval 3", "Oval 4", "Oval 16", "Oval 22", "��� 7")
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    
    btnNames = Array("Oval 5", "Oval 6", "Oval 17", "Oval 18", "Oval 23", "��� 6") ' ���s�W�٦C��
    
    leftPos = 100
    spacing = 120

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    btnNames = Array("Oval 9", "Oval 11", "Oval 12", "Oval 24", "Oval 25", "Oval 15") ' ���s�W�٦C��
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i



    btnNames = Array("Oval 13", "Oval 14", "Oval 19", "Oval 20", "Oval 26", "��� 1") ' ���s�W�٦C��
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    
    
    btnNames = Array("Oval 30", "DAEWON", "Text Box 882", "Text Box 27", "�ꨤ�x�� 13", "chipbond-1") ' ���s�W�٦C��
    
    leftPos = 100
    spacing = 120
    

    For i = LBound(btnNames) To UBound(btnNames)
        Set btn = ws.Shapes(btnNames(i))
        btn.Top = 50 + 66 + 66 + 66 + 66
        btn.Left = leftPos
        leftPos = leftPos + spacing
    Next i
    
    btnNames = Array("chipbondCM", "chipbondCM1") ' ���s�W�٦C��
    
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
    
    MsgBox "�Ҧ����s�w��w�A�L�k���ʡI", vbInformation, "���\"
End Sub

Sub ��s11()
Application.ScreenUpdating = False
    x = InputBox("�ɮצW��")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = "=VLOOKUP(R35C2,'[1-1.xlsm]��J�����'!R34C2:R65536C19,2,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 18
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[1-1.xlsm]��J�����'!R34C2:R65536C19," & i & ",0)"
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

Sub �۰ʱa�J()

Application.ScreenUpdating = False
    y = InputBox("��J�Ƹ�")
    x = InputBox("�ɮצW��")
Set wb = Workbooks(x)
Set ws = wb.Sheets(1)
Set formulaRange = ws.Range("M35:M46")
Set formulaRangel = ws.Range("N35:N46")
    If Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = "" Or Workbooks(x).Sheets("sheet1").Cells(35, 3).Value = "" Then
    Workbooks(x).Sheets("sheet1").Cells(5, 8).FormulaR1C1 = _
    "=VLOOKUP(R35C2,'[" & y & ".xlsx]�u�@��1'!R4C1:R65536C23,3,0)"
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Value = Workbooks(x).Sheets("sheet1").Cells(5, 8).Value
    Workbooks(x).Sheets("sheet1").Cells(5, 8).Replace "#N/A", ""
    End If
    
    For A = 35 To 100
        If Workbooks(x).Sheets("sheet1").Cells(A, 2).Value <> "" And Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            For i = 3 To 12
            With Workbooks(x).Sheets("sheet1").Cells(A, i)
            .FormulaR1C1 = "=VLOOKUP(R" & A & "C2,'[" & y & ".xlsx]�u�@��1'!R4C1:R65536C23," & i + 1 & ",0)"
            .Value = .Value
            .Replace "#VALUE!", "0"
            .Replace "#N/A", ""
            End With
            Next i

    formulaText = "=IF(B35="""","""",VLOOKUP(B35,'[" & y & ".xlsx]�u�@��1'!$A$4:$W$32474,15,FALSE))"
    formulaRange.Formula = formulaText
    Application.Calculate
    formulaRange.Value = formulaRange.Value
    
    formulaText = "=IF(B35="""","""",VLOOKUP(B35,'[" & y & ".xlsx]�u�@��1'!$A$4:$W$32474,17,FALSE))"
    formulaRangel.Formula = formulaText
    Application.Calculate
    formulaRangel.Value = formulaRangel.Value
        ElseIf Workbooks(x).Sheets("sheet1").Cells(A, 3).Value = "" Then
            Exit Sub
        End If
    Next A
Application.ScreenUpdating = True
End Sub

