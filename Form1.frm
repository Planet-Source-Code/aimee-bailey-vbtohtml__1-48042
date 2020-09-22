VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SB-Soft - VB To HTML"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMD2 
      Left            =   3600
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3600
      TabIndex        =   18
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&GO!"
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Edit (.frm)"
      Height          =   6135
      Left            =   6480
      TabIndex        =   14
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Text2 
         Height          =   5775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Text            =   "Form1.frx":0E42
         Top             =   240
         Width           =   4575
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   1455
         Left            =   1200
         TabIndex        =   32
         Top             =   2280
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2566
         _Version        =   393217
         TextRTF         =   $"Form1.frx":0E48
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Edit >"
      Height          =   255
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Browse..."
      Height          =   255
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Text            =   ".frm"
      Top             =   360
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selected FileInfo"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   6255
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1800
         Picture         =   "Form1.frx":0EC8
         ScaleHeight     =   225
         ScaleWidth      =   3945
         TabIndex        =   12
         Top             =   720
         Width           =   3975
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "none"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1845
         Top             =   765
         Width           =   3975
      End
      Begin VB.Label Label7 
         Caption         =   "none"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "none"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Lines Of Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "FileType:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Name (Not Filename):"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Export Options..."
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   6255
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2040
         Top             =   840
      End
      Begin VB.ComboBox Combo2 
         Height          =   330
         ItemData        =   "Form1.frx":120A
         Left            =   3000
         List            =   "Form1.frx":1214
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   840
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   240
         Width           =   4575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   330
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "abc123"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "example:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Colour Scheme"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "FileName"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Font Face"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Contents"
      Height          =   2295
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   6255
      Begin MSComDlg.CommonDialog CMD 
         Left            =   4080
         Top             =   1200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList file_type 
         Left            =   5280
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":122C
               Key             =   "Form"
               Object.Tag             =   "frm"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":157E
               Key             =   "Module"
               Object.Tag             =   "bas"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":18D0
               Key             =   "Class Module"
               Object.Tag             =   "cls"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C22
               Key             =   "User Control"
               Object.Tag             =   "ctl"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1F74
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList code_type 
         Left            =   4680
         Top             =   1080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":22C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2618
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":296A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "code_type"
         SmallIcons      =   "code_type"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.ListBox List2 
      Height          =   1740
      Left            =   2160
      TabIndex        =   29
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   1440
      TabIndex        =   30
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   690
      Left            =   2040
      TabIndex        =   31
      Top             =   2040
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1740
      Left            =   4080
      TabIndex        =   28
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "By Steve Bailey (djhybridfactor@hotmail.com)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Select A File"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xDIM As Integer
Dim xFUNC As Integer
Dim xSUB As Integer
Dim xDEC As Integer
Dim xCODE As Integer
Dim xPROP As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        If Label6.Caption = "None" Then Check1.Value = 0
        Me.Width = 11475
        Frame3.Caption = "Edit (" & GetFile(Text1.Text) & ")"
        Text2.Text = ""
        Dim x As String
            For i = 0 To List1.ListCount - 1
                x = x & List1.list(i) & vbCrLf
            Next i
        Text2.Text = x
    Else
        Me.Width = 6555
    End If
x = Screen.Width / 2
y = Me.Width / 2
Me.Left = x - y
End Sub
Public Function GetFile(file As String) As String
Dim x As Integer
Dim i As Integer
For i = 1 To Len(file)
If Mid(file, i, 1) = "\" Then x = i
Next i
GetFile = Mid(file, x + 1, 255)
End Function



Private Sub Combo2_Click()
FormatCode.ChangeColorScheme Combo2.Text
End Sub

Private Sub Command1_Click()
On Error GoTo err
CMD.CancelError = True
CMD.Filter = "VB Files [*.frm;*.bas;*.cls;*.ctl]|*.frm;*.bas;*.cls;*.ctl*"
CMD.ShowOpen
Text1.Text = CMD.FileName
err:
End Sub

Private Sub Command2_Click()
Dim overwrite As Boolean
overwrite = CheckFile(Text3.Text)
If overwrite = False Then MsgBox "Canceled!", vbOKOnly, "VBtoHTML": FormatCode.dcancel = True: Exit Sub
Form3.Show

If Label7.Caption = "0" Then
    MsgBox "Cannot convert!" & vbCrLf & "reason: No Code To Convert!!!!!", vbCritical, "ERROR"
ElseIf Label7.Caption = "none" Then
    MsgBox "Cannot convert!" & vbCrLf & "reason: No Code To Convert!!!!!", vbCritical, "ERROR"
ElseIf Trim(Text3.Text) = "" Then
    MsgBox "Cannot convert!" & vbCrLf & "reason: No File To Convert To!!!!!", vbCritical, "ERROR"
Else
    Dim x As String
            For i = 0 To List1.ListCount - 1
                x = x & List1.list(i) & vbCrLf
            Next i
        If Me.Width = 6555 Then
        Text2.Text = x
        Else
            If Text2.Text = Trim("") Then
            Text2.Text = x
            End If
        End If
    Text4.Text = Text2.Text
    If FormatCode.dcancel = True Then Form3.Label1.Caption = "Canceled!": Exit Sub
    Text4.Text = Format(Text4.Text, List2, List3)
    If overwrite = True Then
        Text4.Text = HTMLHeader & Text4.Text & HTMLFooter
        'On Error GoTo err
        Open Text3.Text For Output As #1
        Print #1, Text4.Text
        Close #1
        Form3.Label1.Caption = "Complete!"
        Form3.Command2.Visible = False
    End If
End If
Exit Sub
err:
MsgBox "Cannot convert!" & vbCrLf & "reason: Error Writing To Output File!!!!!", vbCritical, "ERROR"
End Sub

Public Function CheckFile(file As String) As Boolean
On Error GoTo err
x = FileLen(file)
y = MsgBox(file & vbCrLf & " Already Exists!" & vbCrLf & "Overwrite?", vbYesNo)
If y = vbNo Then CheckFile = False: Exit Function
err:
CheckFile = True
End Function

Public Function HTMLHeader() As String
x = "<HTML>" & vbCrLf
x = x & "<HEAD>" & vbCrLf
x = x & "<TITLE>Code From " & Label6.Caption & "</TITLE>" & vbCrLf
x = x & "<HTML>" & vbCrLf
x = x & "<BODY>" & vbCrLf & vbCrLf
x = x & "<FONT FACE=" & Brack(Combo1.Text) & " BGCOLOR=" & Brack("#FFFFFF") & " TEXT=" & Brack(FormatCode.txtNormalCol) & ">" & vbCrLf
HTMLHeader = x
End Function

Public Function HTMLFooter() As String
x = vbCrLf & "</BODY>" & vbCrLf
x = x & "</HTML>"
HTMLFooter = x
End Function

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
On Error GoTo err
CMD2.CancelError = True
CMD2.Filter = "HTML Files [*.htm;*.html]|*.htm;*.html*"
CMD2.ShowSave
Text3.Text = CMD2.FileName
err:
End Sub

Private Sub Command5_Click()
Text2.Text = FormatCode.Format(Text2.Text, List2, List3)
End Sub

Private Sub Form_Load()
FormatCode.LoadDefaults
Combo2.ListIndex = 0
Me.Width = 6555
GetFonts Combo1 'Send All Fonts To combo1
GotoFont Combo1, "Arial" 'Choose Default Font
End Sub

Public Function LoadInfo(file As String)
'xtype(0) = module,class
'xtype(1) = form,user-control
'Open file For Input As #1
'Text2.Text = ""
'Do Until EOF(1)
'DoEvents
'Input #1, a$
'Text2.Text = Text2.Text & a$ & vbCrLf
'Loop
rtb.LoadFile file
Text2.Text = rtb.Text
FormatCode.ConvertToList Text2.Text, List1, vbCrLf
'Close #1
    StripHeader
    Label7.Caption = List1.ListCount - 1
    DoCounts
End Function
Public Function DoCounts()
' 1
' 2 Sub, function,
' 3 dim, public declare, public const
Dim line As String
xDIM = 0: xFUNC = 0: xSUB = 0: xDEC = 0: xCODE = 0: xPROP = 0
ListView1.ListItems.Clear
For i = 0 To List1.ListCount - 1
line = LCase(List1.list(i))
    If Mid(line, 1, 3) = "dim" Then
        xDIM = xDIM + 1
        
    ElseIf Mid(line, 1, 6) = "public" Then
            If Mid(line, 8, 8) = "function" Then
                xFUNC = xFUNC + 1
            ElseIf Mid(line, 8, 7) = "declare" Then
                xDEC = xDEC + 1
            ElseIf Mid(line, 8, 3) = "sub" Then
                xSUB = xSUB + 1
            ElseIf Mid(line, 8, 8) = "property" Then
                xPROP = xPROP + 1
            End If
    ElseIf Mid(line, 1, 7) = "private" Then
            If Mid(line, 8, 8) = "function" Then
                xFUNC = xFUNC + 1
            ElseIf Mid(line, 8, 7) = "declare" Then
                xDEC = xDEC + 1
            ElseIf Mid(line, 8, 3) = "sub" Then
                xSUB = xSUB + 1
            End If
    Else
        xCODE = xCODE + 1
    End If
Next i
    
    If xDIM > 0 Then
        ListView1.ListItems.Add , , "Dim Statements", 1, 1
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xDIM
    End If
    
    If xDEC > 0 Then
        ListView1.ListItems.Add , , "Declare Statements", 1, 1
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xDEC
    End If
    
    If xFUNC > 0 Then
        ListView1.ListItems.Add , , "Function Statements", 2, 2
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xFUNC
    End If
    
    If xSUB > 0 Then
        ListView1.ListItems.Add , , "Sub Statements", 2, 2
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xSUB
    End If
    
    If xPROP > 0 Then
        ListView1.ListItems.Add , , "Property Statements", 2, 2
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xPROP
    End If
    
    If xCODE > 0 Then
        ListView1.ListItems.Add , , "Other VB Line", 3, 3
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , xCODE
    Else
        ListView1.ListItems.Add , , "Other VB Line", 3, 3
        ListView1.ListItems(ListView1.ListItems.Count).ListSubItems.Add , , "0    Very Strange!!!"
    End If

End Function
Public Function StripHeader()
For i = 0 To List1.ListCount - 1
            If Mid(List1.list(0), 1, 17) = "Attribute VB_Name" Then
                Label6.Caption = Trim(Replace(Mid(List1.list(0), 20, 255), Chr(34), ""))
                For b = 0 To 20
                    If Mid(List1.list(0), 1, 9) <> "Attribute" Then GoTo err
                    List1.RemoveItem 0
                Next b
            Else
                List1.RemoveItem 0
            End If
        Next i
err:
End Function


Private Sub Text1_Change()
'On Error GoTo err
Command2.Enabled = False
Dim found As Boolean
found = False
Me.Width = 6555
If Mid(Right(Text1.Text, 4), 1, 1) = "." Then
DoEvents
    For i = 1 To file_type.ListImages.Count - 1
        If LCase(file_type.ListImages(i).Tag) = LCase(Right(Text1.Text, 3)) Then
        Picture1.Picture = file_type.ListImages(i).Picture
        Label5.Caption = file_type.ListImages(i).Key
        'Else
            'Load File INFO
            LoadInfo Text1.Text
            
        found = True
        End If
    Next i
    
End If

If found = False Then
    Picture1.Picture = file_type.ListImages(5).Picture
    Label5.Caption = "None"
End If

err:
Command2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label12.Font = Combo1.Text
End Sub
