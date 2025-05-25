Attribute VB_Name = "CreateNewMacros"
Sub ImportTextFileWithSpaceAndEmptyValueHandling()
    Dim ws As Worksheet
    Dim txtFile As String
    Dim importRange As Range
    Dim txtContent As String
    Dim cleanedTxtFile As String
    Dim fso As Object
    Dim txtFileStream As Object
    Dim stream As Object

    txtFile = "G:\效率\ADG.txt"
    cleanedTxtFile = "G:\效率\Test.txt"
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile txtFile
    txtContent = stream.ReadText(-1)
    stream.Close
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFileStream = fso.CreateTextFile(cleanedTxtFile, True)
    txtContent = Replace(txtContent, "  ", " ")
    Do While InStr(txtContent, "  ") > 0
        txtContent = Replace(txtContent, "  ", " ")
    Loop
    txtContent = Replace(txtContent, "    ", "")
    txtFileStream.Write txtContent
    txtFileStream.Close
    Set ws = ThisWorkbook.Sheets(1)
    Set importRange = ws.Range("A1")
    With ws.QueryTables.Add(Connection:="TEXT;" & cleanedTxtFile, Destination:=importRange)
        .TextFileSpaceDelimiter = True
        .TextFileTabDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileOtherDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1)
        .Refresh BackgroundQuery:=False
    End With
End Sub
Sub ImportAndSplitTXT()
    Dim ws As Worksheet
    Dim filePath As String
    Dim txtFile As Integer
    Dim line As String
    Dim RowNum As Integer
    Dim cellValues() As String
    Dim i As Integer
    
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear
    filePath = "G:\效率\Test.txt"
    txtFile = FreeFile
    Open filePath For Input As txtFile
    RowNum = 1
    Do Until EOF(txtFile)
        Line Input #txtFile, line
        cellValues = Split(line, " ")
        For i = LBound(cellValues) To UBound(cellValues)
            ws.Cells(RowNum, i + 1).Value = cellValues(i)
        Next i
        RowNum = RowNum + 1
    Loop
    Close txtFile
End Sub
