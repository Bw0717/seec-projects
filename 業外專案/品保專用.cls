VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub 匯出()
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim x As String


    x = InputBox("檔案名稱")
    If Not x Like "*.xlsm" Then x = x & ".xlsm"


    Set wbSource = ThisWorkbook
    On Error Resume Next
    Set wbTarget = Workbooks(x)
    On Error GoTo 0


    If wbTarget Is Nothing Then
        MsgBox "請確認 '" & x & "' 已開啟。", vbExclamation, "錯誤"
        Exit Sub
    End If


    Set wsSource = wbSource.Sheets("品保專用")
    Set wsTarget = wbTarget.Sheets("一般")


    wsTarget.Range("L7").Value = wsSource.Range("L4").Value
    wsTarget.Range("H5").Value = wsSource.Range("N4").Value
    wsTarget.Range("L5").Value = ""
    Set Rng = wsTarget.Range("L5")
    With Rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISBLANK(L5)")
        .Interior.Color = RGB(255, 255, 0)
    End With
    wsTarget.Range("F36").Value = wsSource.Range("C5").Value
    wsTarget.Range("F37").Value = wsSource.Range("C6").Value
    wsTarget.Range("F38").Value = wsSource.Range("C7").Value
    
    wsTarget.Range("J36").Value = wsSource.Range("D5").Value
    wsTarget.Range("J37").Value = wsSource.Range("D6").Value
    wsTarget.Range("J38").Value = wsSource.Range("D7").Value
    
    wsTarget.Range("K36").Value = wsSource.Range("E5").Value
    wsTarget.Range("K37").Value = wsSource.Range("E6").Value
    wsTarget.Range("K38").Value = wsSource.Range("E7").Value
    
    wsTarget.Range("L36").Value = wsSource.Range("F5").Value
    wsTarget.Range("L37").Value = wsSource.Range("F6").Value
    wsTarget.Range("L38").Value = wsSource.Range("F7").Value
    
    wsTarget.Range("M36").Value = wsSource.Range("G5").Value
    wsTarget.Range("M37").Value = wsSource.Range("G6").Value
    wsTarget.Range("M38").Value = wsSource.Range("G7").Value
    
    wsTarget.Range("N36").Value = wsSource.Range("H5").Value
    wsTarget.Range("N37").Value = wsSource.Range("H6").Value
    wsTarget.Range("N38").Value = wsSource.Range("H7").Value
    
    wsTarget.Range("O36").Value = wsSource.Range("I5").Value
    wsTarget.Range("O37").Value = wsSource.Range("I6").Value
    wsTarget.Range("O38").Value = wsSource.Range("I7").Value
    
    wsTarget.Range("E36").Value = wsSource.Range("J5").Value
    wsTarget.Range("E37").Value = wsSource.Range("J6").Value
    wsTarget.Range("E38").Value = wsSource.Range("J7").Value
    
    wsTarget.Range("H36").Value = IIf(IsError(wsSource.Range("K5").Value) Or wsSource.Range("K5").Value = "N/A", "NIL", wsSource.Range("K5").Value)
    wsTarget.Range("H37").Value = IIf(IsError(wsSource.Range("K6").Value) Or wsSource.Range("K6").Value = "N/A", "NIL", wsSource.Range("K6").Value)
    wsTarget.Range("H38").Value = IIf(IsError(wsSource.Range("K7").Value) Or wsSource.Range("K7").Value = "N/A", "NIL", wsSource.Range("K7").Value)
    
End Sub

Sub 匯出2()
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim x As String


    x = InputBox("檔案名稱")
    If Not x Like "*.xlsm" Then x = x & ".xlsm"


    Set wbSource = ThisWorkbook
    On Error Resume Next
    Set wbTarget = Workbooks(x)
    On Error GoTo 0


    If wbTarget Is Nothing Then
        MsgBox "請確認 '" & x & "' 已開啟。", vbExclamation, "錯誤"
        Exit Sub
    End If


    Set wsSource = wbSource.Sheets("品保專用")
    Set wsTarget = wbTarget.Sheets("一般")


    wsTarget.Range("L7").Value = wsSource.Range("L4").Value
    wsTarget.Range("H5").Value = wsSource.Range("N4").Value
    wsTarget.Range("L5").Value = ""
    Set Rng = wsTarget.Range("L5")

    wsTarget.Range("F36").Value = wsSource.Range("C5").Value    'F
    wsTarget.Range("F37").Value = wsSource.Range("C6").Value
    wsTarget.Range("F38").Value = wsSource.Range("C7").Value
    
    wsTarget.Range("J36").Value = wsSource.Range("D5").Value    'P0 x 10
    wsTarget.Range("J37").Value = wsSource.Range("D6").Value
    wsTarget.Range("J38").Value = wsSource.Range("D7").Value
    
    wsTarget.Range("K36").Value = wsSource.Range("E5").Value    'P2
    wsTarget.Range("K37").Value = wsSource.Range("E6").Value
    wsTarget.Range("K38").Value = wsSource.Range("E7").Value
    
    wsTarget.Range("L36").Value = wsSource.Range("F5").Value    'A0
    wsTarget.Range("L37").Value = wsSource.Range("F6").Value
    wsTarget.Range("L38").Value = wsSource.Range("F7").Value
    
    wsTarget.Range("M36").Value = wsSource.Range("G5").Value    'B0
    wsTarget.Range("M37").Value = wsSource.Range("G6").Value
    wsTarget.Range("M38").Value = wsSource.Range("G7").Value
    
    wsTarget.Range("N36").Value = wsSource.Range("H5").Value    'K0
    wsTarget.Range("N37").Value = wsSource.Range("H6").Value
    wsTarget.Range("N38").Value = wsSource.Range("H7").Value
    
    wsTarget.Range("O36").Value = wsSource.Range("I5").Value    'T
    wsTarget.Range("O37").Value = wsSource.Range("I6").Value
    wsTarget.Range("O38").Value = wsSource.Range("I7").Value
    
    wsTarget.Range("E36").Value = wsSource.Range("J5").Value    'E
    wsTarget.Range("E37").Value = wsSource.Range("J6").Value
    wsTarget.Range("E38").Value = wsSource.Range("J7").Value
    
    wsTarget.Range("H36").Value = IIf(IsError(wsSource.Range("K5").Value) Or wsSource.Range("K5").Value = "N/A", "NIL", wsSource.Range("K5").Value) 'D1
    wsTarget.Range("H37").Value = IIf(IsError(wsSource.Range("K6").Value) Or wsSource.Range("K6").Value = "N/A", "NIL", wsSource.Range("K6").Value)
    wsTarget.Range("H38").Value = IIf(IsError(wsSource.Range("K7").Value) Or wsSource.Range("K7").Value = "N/A", "NIL", wsSource.Range("K7").Value)
    
End Sub
