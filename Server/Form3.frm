VERSION 5.00
Object = "{6685E735-3BF6-11D1-A345-444553540000}#1.1#0"; "GRADIENTTITLE.OCX"
Begin VB.Form Form3 
   Caption         =   " "
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   ScaleHeight     =   3465
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close Window"
      Enabled         =   0   'False
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox URL 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin GradientTitle.graTitle graTitle1 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      GradEndColor    =   0
      InGradStartColor=   0
      InGradEndColor  =   0
      SysColor        =   0   'False
   End
   Begin VB.PictureBox ImageR 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   0
      MouseIcon       =   "Form3.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   2295
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      Begin VB.Timer Timer2 
         Left            =   480
         Top             =   360
      End
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

Private Sub ImageR_Click()
If URL.Text <> "" Then
ShellExecute hwnd, "open", URL.Text, vbNullString, vbNullString, 0
End If
End Sub

Private Sub Timer2_Timer()
Command1.Enabled = True
Timer2.Enabled = False
End Sub
