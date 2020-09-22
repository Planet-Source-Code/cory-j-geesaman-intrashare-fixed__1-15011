VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   Caption         =   "Using IntraShare"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form7"
   ScaleHeight     =   5895
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Features 
      Height          =   4575
      Left            =   0
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Label Label7 
         Caption         =   $"Form7.frx":0000
         Height          =   4215
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7575
      End
   End
   Begin VB.Frame Downloads 
      Height          =   4575
      Left            =   0
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Label Label6 
         Caption         =   $"Form7.frx":037E
         Height          =   4215
         Left            =   3600
         TabIndex        =   12
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image Image5 
         Height          =   2895
         Left            =   120
         Picture         =   "Form7.frx":045D
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame Sharing 
      Height          =   4575
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Image Image4 
         Height          =   2895
         Left            =   120
         Picture         =   "Form7.frx":B17F
         Top             =   240
         Width           =   3360
      End
      Begin VB.Label Label5 
         Caption         =   $"Form7.frx":15EA1
         Height          =   4215
         Left            =   3600
         TabIndex        =   10
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Search 
      Height          =   4575
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Label Label4 
         Caption         =   $"Form7.frx":1604D
         Height          =   4215
         Left            =   3600
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
      Begin VB.Image Image3 
         Height          =   2895
         Left            =   120
         Picture         =   "Form7.frx":1612E
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame Chat 
      Height          =   4575
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   7815
      Begin VB.Image Image2 
         Height          =   2895
         Left            =   120
         Picture         =   "Form7.frx":20E50
         Top             =   240
         Width           =   3360
      End
      Begin VB.Label Label3 
         Caption         =   $"Form7.frx":2BB72
         Height          =   4215
         Left            =   3600
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5640
      Width           =   7815
   End
   Begin VB.Frame Connect 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   7815
      Begin VB.Label Label2 
         Caption         =   $"Form7.frx":2BE47
         Height          =   4215
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   4200
         Left            =   120
         Picture         =   "Form7.frx":2C1F2
         Top             =   240
         Width           =   3030
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      Max             =   6
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Connecting"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Slider1_Click()
Select Case Slider1.Value
Case 1
Label1.Caption = "Connecting"
Connect.Visible = True
Chat.Visible = False
Search.Visible = False
Sharing.Visible = False
Downloads.Visible = False
Features.Visible = False
Case 2
Label1.Caption = "Chating"
Connect.Visible = False
Chat.Visible = True
Search.Visible = False
Sharing.Visible = False
Downloads.Visible = False
Features.Visible = False
Case 3
Label1.Caption = "Search"
Connect.Visible = False
Chat.Visible = False
Search.Visible = True
Sharing.Visible = False
Downloads.Visible = False
Features.Visible = False
Case 4
Label1.Caption = "Sharing files"
Connect.Visible = False
Chat.Visible = False
Search.Visible = False
Sharing.Visible = True
Downloads.Visible = False
Features.Visible = False
Case 5
Label1.Caption = "Downloading Files"
Connect.Visible = False
Chat.Visible = False
Search.Visible = False
Sharing.Visible = False
Downloads.Visible = True
Features.Visible = False
Case 6
Label1.Caption = "Features"
Connect.Visible = False
Chat.Visible = False
Search.Visible = False
Sharing.Visible = False
Downloads.Visible = False
Features.Visible = True
End Select
End Sub
