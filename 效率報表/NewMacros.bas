Attribute VB_Name = "NewMacros"
Sub ImportMultipleTxtFiles()
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
    txtFiles = Array("G:\�Ĳv\ADG.txt", "G:\�Ĳv\AEV.txt", "G:\�Ĳv\AM1.txt", _
                     "G:\�Ĳv\AM2.txt", "G:\�Ĳv\AMG.txt", "G:\�Ĳv\ARD.txt", "G:\�Ĳv\ARG.txt")
    keywords = Array("ADG", "AEV", "AM1", "AM2", "AMG", "ARD", "ARG")
    For i = LBound(txtFiles) To UBound(txtFiles)
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "�פJ���_" & keywords(i)
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
        ws.Range("A8:Q8").Value = Array("��´", "�Z�O", "���", "�u�渹�X", "�ե�", "�귽�W��", "ú�w�q", "�귽���u��", "�e�@�~�귽���u��", _
                                        "ú�w", "�C��u��", "�~�]���s", "�����䴩", "C�u��", "�ײz�u��", "�է@�u��", "���Ҵ��u")
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
    MsgBox "���\�פJ", vbInformation
End Sub



