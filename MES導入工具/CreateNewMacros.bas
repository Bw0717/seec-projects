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


    FilePath = Application.GetOpenFilename("Text Files (*.tsv), *.tsv", , "��ܭn�פJ�� tsv �ɮ�")
    If FilePath = "False" Then
        MsgBox "������ɮסA�{���פ�C", vbExclamation
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
    
    deleteCols = Array("����", "���A", "�u�{�Ƹ�", "�Ƹ��Ǹ�", "�@�~�Ǹ�", "���N���c", "�u�{�ήƪ�", "����", "���", "�p�e�ʤ���", "�}�~�v", "�p�p�ƶq", "���ĩʱ���", "�_�l", "�פ�", "�_�l���", "�פ���", "����", "�w�ɤJ", "�u�{�ܧ��", "�ѵ����A", "�ܮw", "�x��", "�p�⦨��", "��즨��", "�p�p�ƶq", "�p�p����", "�@�~�Ǹ�", "�s�y", "���w", "�ֿn�s�y", "�ֿn�`�p", "��ܩ�", "����", "ATP", "�̤p�ƶq", "�̤j�ƶq", "�P��q����", "�i�X�f", "�ǤJ�X�f���", "�X�f�һ�", "���J�һ�")
    For i = Rng.Columns.Count To 1 Step -1
        For Each colName In deleteCols
            If ws.Cells(1, i).Value = colName Then
                ws.Columns(i).Delete
                Exit For
            End If
        Next colName
    Next i
            
           
 
    ws.Columns.AutoFit

    SavePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="�t�s�s��")
    If SavePath <> "False" Then
        newWb.SaveAs Filename:=SavePath, FileFormat:=51
        MsgBox "���ɧ����I", vbInformation
    Else
        MsgBox "���ɨ����C", vbExclamation
    End If
End Sub

