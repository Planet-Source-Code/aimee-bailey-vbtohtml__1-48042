VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processing..."
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5025
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while calculating..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
FormatCode.dcancel = True
Command2.Visible = False
End Sub
