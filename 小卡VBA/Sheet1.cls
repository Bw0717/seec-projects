VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GenerateQRCodes_E10()
    Dim ws As Worksheet
    Dim i As Long
    Dim qrText As String
    Dim qrURL As String
    Dim targetCell As Range
    Dim pic As Picture
    Dim imgLeft As Double, imgTop As Double

    Set ws = ThisWorkbook.Sheets("托架與電樞組立")
    For i = 5 To 10000 Step 7
        qrText = Trim(ws.Range("C" & i).Value)
        If qrText <> "" Then
            qrURL = "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=" & URLEncode(qrText)
            Set targetCell = ws.Range("C" & i - 1)
            imgLeft = targetCell.Left
            imgTop = targetCell.Top

            Set pic = ws.Pictures.Insert(qrURL)
            With pic
                .Left = imgLeft + 28
                .Top = imgTop + 10
                .Width = 70
                .Height = 70
            End With
        End If
    Next i
End Sub
Function URLEncode(str As String) As String
    Dim i As Long
    Dim ch As String
    Dim encoded As String
    encoded = ""
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 97 To 122
                encoded = encoded & ch
            Case Else
                encoded = encoded & "%" & Hex(AscW(ch))
        End Select
    Next i
    URLEncode = encoded
End Function
Sub ClearAllPictures_WPS()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("托架與電樞組立")
    Dim shp As Shape
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If InStr(1, shp.Name, "Picture", vbTextCompare) > 0 Or _
           InStr(1, shp.Name, "圖片", vbTextCompare) > 0 Or _
           InStr(1, shp.Name, "Image", vbTextCompare) > 0 Then
            shp.Delete
        End If
    Next i
End Sub

Private Sub CommandButton1_Click()
Call ClearAllPictures_WPS
Call GenerateQRCodes_E10
End Sub










