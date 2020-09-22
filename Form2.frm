VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Custom Scheme"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Change"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0352
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   2040
      TabIndex        =   8
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Color"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.VScrollBar VScroll4 
      Height          =   1935
      Left            =   1660
      Max             =   255
      TabIndex        =   7
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll3 
      Height          =   1935
      Left            =   1130
      Max             =   255
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1935
      Left            =   600
      Max             =   255
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      Left            =   130
      Max             =   255
      TabIndex        =   3
      Top             =   360
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   240
      ScaleHeight     =   1395
      ScaleWidth      =   1515
      TabIndex        =   12
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Keywords??"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   5160
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   5160
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Types"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "B&&W"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "BLUE"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "GREEN"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "RED"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text1_Change()

End Sub

Private Sub Command1_Click()
FormatCode.txtNormalCol = "#" & ListView1.ListItems(1).ListSubItems(1).Text
FormatCode.txtCommentCol = "#" & ListView1.ListItems(1).ListSubItems(1).Text
FormatCode.txtKeywordCol = "#" & ListView1.ListItems(1).ListSubItems(1).Text
Unload Me
End Sub

Private Sub Command2_Click()
Form1.Combo2.ListIndex = 0
Unload Me
End Sub

Private Sub Command3_Click()
ListView1.SelectedItem.ListSubItems(1).ForeColor = Picture1.BackColor
ListView1.SelectedItem.ListSubItems(1).Text = Long2Hex(Picture1.BackColor)
End Sub


Private Sub Form_Load()
ListView1.ColumnHeaders(1).Width = 2190.047
ListView1.ColumnHeaders(2).Width = 854.9292
FormatCode.LoadDefaults
ListView1.ListItems.Add , , "Normal", 1, 1
ListView1.ListItems.Add , , "Comments", 1, 1
ListView1.ListItems.Add , , "Keywords", 2, 2
ListView1.ListItems(1).ListSubItems.Add , , FormatCode.txtNormalCol
ListView1.ListItems(1).ListSubItems(1).ForeColor = colorCon.Hex2Long(FormatCode.txtNormalCol)
ListView1.ListItems(2).ListSubItems.Add , , FormatCode.txtCommentCol
ListView1.ListItems(2).ListSubItems(1).ForeColor = colorCon.Hex2Long(FormatCode.txtCommentCol)
ListView1.ListItems(3).ListSubItems.Add , , FormatCode.txtKeywordCol
ListView1.ListItems(3).ListSubItems(1).ForeColor = colorCon.Hex2Long(FormatCode.txtKeywordCol)
Dim r As Byte
Dim g As Byte
Dim b As Byte
colorCon.Long2RGB CLng(ListView1.SelectedItem.ListSubItems(1).ForeColor), r, g, b
VScroll1.Value = r
VScroll2.Value = g
VScroll3.Value = b
End Sub

Private Sub Label5_Click()
MsgBox "Keywords are:" & vbCrLf & vbCrLf & FormatCode.Keywords & " etc...", vbInformation, "Question"
End Sub

Private Sub ListView1_Click()
Dim r As Byte
Dim g As Byte
Dim b As Byte
colorCon.Long2RGB CLng(ListView1.SelectedItem.ListSubItems(1).ForeColor), r, g, b
VScroll1.Value = r
VScroll2.Value = g
VScroll3.Value = b
End Sub

Private Sub VScroll1_Change()
Picture1.BackColor = colorCon.RGB2Long(VScroll1.Value, VScroll2.Value, VScroll3.Value)
DoMedian
End Sub
Private Sub VScroll2_Change()
Picture1.BackColor = colorCon.RGB2Long(VScroll1.Value, VScroll2.Value, VScroll3.Value)
DoMedian
End Sub
Private Sub VScroll3_Change()
Picture1.BackColor = colorCon.RGB2Long(VScroll1.Value, VScroll2.Value, VScroll3.Value)
DoMedian
End Sub

Public Function DoMedian()
VScroll4.Value = VScroll1.Value
End Function

Public Function EditCol()
On Error Resume Next
ListView1.SelectedItem.ListSubItems(1).ForeColor = Picture1.BackColor
ListView1.SelectedItem.ListSubItems(1).Text = Long2Hex(Picture1.BackColor)
End Function

Private Sub VScroll1_Scroll()
EditCol
End Sub
Private Sub VScroll2_Scroll()
EditCol
End Sub
Private Sub VScroll3_Scroll()
EditCol
End Sub

Private Sub VScroll4_Scroll()
Dim t As Integer
On Error Resume Next
t = VScroll4.Value

    VScroll1.Value = t
    VScroll2.Value = t
    VScroll3.Value = t
EditCol
End Sub
