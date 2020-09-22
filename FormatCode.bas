Attribute VB_Name = "FormatCode"
Public txtCommentCol As String
Public txtKeywordCol As String
Public txtNormalCol As String
Public Keywords As String
Public dcancel As Boolean

Public Function LoadDefaults()
txtCommentCol = "008000"
txtNormalCol = "000000"
txtKeywordCol = "0000A0"

x = "Alias|Event|Public|Private|Global|Declare|Type|Const|Dim|Variant|Sub|Function|"
x = x & "Property|End|Exit|Next|As|For|To|On|Error|Resume|Go|If|While|True|"
x = x & "False|Byte|CByte|New|Boolean|Currency|Date|Double|Long|String|"
x = x & "Integer|CLng|Single|Loop|DoEvents|Do|Then|Select|Case|Get|Let|"
x = x & "Set|Property|Print|Input|Output|Open|With|Until|Else|ByVal|Lib"
Keywords = x
End Function

Public Function ChangeColorScheme(scheme As String)
Select Case LCase(scheme)
    Case "default"
        LoadDefaults
    Case "custom..."
        Form2.Show
End Select
End Function

Public Function ConvertToList(str As String, list As ListBox, delimiter As String)
On Error GoTo finished
    For i = 0 To Len(str)
        list.AddItem Split(str, delimiter)(i)
    Next i
finished: End Function


Public Function Format(code As String, tmplist As ListBox, tmpList2 As ListBox) As String
Dim x As String

tmplist.Clear: tmpList2.Clear
ConvertToList code, tmplist, vbCrLf
ConvertToList Keywords, tmpList2, "|"
x = 0
aa:
    For a = 0 To tmplist.ListCount - 1
    DoEvents
    If dcancel = True Then Form3.Label1.Caption = "Canceled!": Exit Function
    Form3.PBar1.Max = tmplist.ListCount - 1
    Form3.PBar1.Value = a
        If InStr(1, tmplist.list(a), "Public Function") > 0 Then
            If tmplist.list(a - 1) <> "<HR>" Then
                tmplist.AddItem "<HR>", a: GoTo aa
            End If
        ElseIf InStr(1, tmplist.list(a), "Public Sub") > 0 Then
            If tmplist.list(a - 1) <> "<HR>" Then
                tmplist.AddItem "<HR>", a: GoTo aa
            End If
        ElseIf InStr(1, tmplist.list(a), "Public Property") > 0 Then
            If tmplist.list(a - 1) <> "<HR>" Then
                tmplist.AddItem "<HR>", a: GoTo aa
            End If
        ElseIf InStr(1, tmplist.list(a), "Private Function") > 0 Then
            If tmplist.list(a - 1) <> "<HR>" Then
                tmplist.AddItem "<HR>", a: GoTo aa
            End If
        ElseIf InStr(1, tmplist.list(a), "Private Sub") > 0 Then
            If tmplist.list(a - 1) <> "<HR>" Then
                tmplist.AddItem "<HR>", a: GoTo aa
            End If
        ElseIf InStr(1, tmplist.list(a), "Private Property") > 0 Then
            If tmplist.list(a - 1) <> "<HR>" Then
                tmplist.AddItem "<HR>", a: GoTo aa
            End If
        End If
    Next a
    For i = 0 To tmplist.ListCount - 1
    If dcancel = True Then Form3.Label1.Caption = "Canceled!": Exit Function
    Form3.PBar1.Max = tmplist.ListCount - 1
    Form3.PBar1.Value = i
            x = x & FormatLine(tmplist.list(i), tmpList2) & "<BR>" & vbCrLf
    Next i
    Format = x

tmplist.Clear: tmpList2.Clear
End Function



Public Function FormatLine(code As String, tmplist As ListBox) As String
    For j = 0 To tmplist.ListCount - 1
        code = DoMargins(code)
        code = DoComments(code)
        code = Replace(code, tmplist.list(j), "<FONT COLOR=" & Brack(txtKeywordCol) & ">" & tmplist.list(j) & "</FONT>")
    Next j
FormatLine = code
End Function

Public Function DoComments(str As String) As String
For i = Len(str) To 1 Step -1
    If Mid(str, i, 1) = Chr(39) Then
    If Right(str, 7) <> "</FONT>" Then
    com = Mid(str, i, Len(str) - i + 1)
    DoComments = Replace(str, com, "<FONT COLOR=" & Brack(txtCommentCol) & ">" & com & "</FONT>")
    GoTo err
    End If
    End If
Next i
DoComments = str
err:
End Function

Public Function DoMargins(str As String) As String
Dim spaces As Integer

If Mid(str, 1, 1) = " " Then
    For i = 1 To Len(str)
        If Mid(str, i, 1) = " " Then spaces = i Else GoTo aa
    Next i
aa:
    DoMargins = Replace(Mid(str, 1, spaces), " ", "&nbsp;") & " " & Trim(str)
Else
DoMargins = str
End If
End Function
