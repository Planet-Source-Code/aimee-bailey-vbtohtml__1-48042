Attribute VB_Name = "colorCon"
Public Sub Hex2RGB(strHexColor As String, r As Byte, g As Byte, b As Byte)
    Dim HexColor As String
    Dim i As Byte
    On Error Resume Next
    ' make sure the string is 6 characters l
    '     ong
    ' (it may have been given in &H###### fo
    '     rmat, we want ######)
    strHexColor = Right((strHexColor), 6)
    ' however, it may also have been given a
    '     s or #***** format, so add 0's in front


    For i = 1 To (6 - Len(strHexColor))
        HexColor = HexColor & "0"
    Next
    HexColor = HexColor & strHexColor
    ' convert each set of 2 characters into
    '     bytes, using vb's cbyte function
    r = CByte("&H" & Right$(HexColor, 2))
    g = CByte("&H" & Mid$(HexColor, 3, 2))
    b = CByte("&H" & Left$(HexColor, 2))
End Sub


Public Function RGB2Hex(r As Byte, g As Byte, b As Byte) As String
    On Error Resume Next
    ' convert to long using vb's rgb functio
    '     n, then use the long2rgb function
    RGB2Hex = Long2Hex(RGB(r, g, b))
End Function


Public Sub Long2RGB(LongColor As Long, r As Byte, g As Byte, b As Byte)
    On Error Resume Next
    ' convert to hex using vb's hex function
    '     , then use the hex2rgb function
    Hex2RGB Hex(LongColor), r, g, b
End Sub


Public Function RGB2Long(r As Byte, g As Byte, b As Byte) As Long
    On Error Resume Next
    ' use vb's rgb function
    RGB2Long = RGB(r, g, b)
End Function


Public Function Long2Hex(LongColor As Long) As String
    On Error Resume Next
    ' use vb's hex function
    Long2Hex = Hex(LongColor)
End Function


Public Function Hex2Long(strHexColor As String) As Long
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    On Error Resume Next
    ' use the hex2rgb function to get the re
    '     d green and blue bytes
    Hex2RGB strHexColor, b, g, r
    ' convert to long using vb's rgb functio
    '     n
    Hex2Long = RGB(r, g, b)
End Function

