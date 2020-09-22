VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Download Location"
   ClientHeight    =   3345
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar NewDLSpot 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3090
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cance l"
      Height          =   1395
      Left            =   4800
      TabIndex        =   1
      Top             =   1650
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   1395
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
If NewDLSpot.SimpleText <> "" Then
Form1.DLLocation = NewDLSpot.SimpleText
Else
MsgBox "No Directory Has Been Selected"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
NewDLSpot.SimpleText = Dir1.Path
If Mid(NewDLSpot.SimpleText, Len(NewDLSpot.SimpleText), 1) = "/" Or Mid(NewDLSpot.SimpleText, Len(NewDLSpot.SimpleText), 1) = "\" Then
NewDLSpot.SimpleText = Right(NewDLSpot, Len(NewDLSpot.SimpleText) - 1)
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Dir1.Path = Form1.DLLocation
End Sub
