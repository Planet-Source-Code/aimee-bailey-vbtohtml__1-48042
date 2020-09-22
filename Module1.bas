Attribute VB_Name = "Module1"
Public Function GetFonts(combo As ComboBox)
For i = 0 To Screen.FontCount
combo.AddItem Screen.Fonts(i)
Next i
End Function

Public Function GotoFont(combo As ComboBox, fnt As String)
For i = 0 To combo.ListCount - 1
    If LCase(combo.list(i)) = LCase(fnt) Then
        combo.ListIndex = i
        Exit Function
    End If
Next i
End Function

Public Function Brack(Optional str As String) As String 'returns string brackets!
Brack = Chr(34) & str & Chr(34)
End Function
