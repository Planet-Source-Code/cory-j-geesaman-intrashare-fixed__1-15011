VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6685E735-3BF6-11D1-A345-444553540000}#1.1#0"; "GRADIENTTITLE.OCX"
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About IntraShare"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   ControlBox      =   0   'False
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6180
   StartUpPosition =   3  'Windows Default
   Begin GradientTitle.graTitle graTitle1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      GradEndColor    =   0
      InGradStartColor=   0
      InGradEndColor  =   0
      SysColor        =   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   5975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3415
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   6033
      _Version        =   393216
      TabHeight       =   520
      MouseIcon       =   "Form5.frx":0442
      TabCaption(0)   =   "Credits"
      TabPicture(0)   =   "Form5.frx":075C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Legal BS"
      TabPicture(1)   =   "Form5.frx":0778
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "WebSite"
      TabPicture(2)   =   "Form5.frx":0794
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox Text2 
         Height          =   2985
         Left            =   -74918
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "Form5.frx":07B0
         Top             =   360
         Width           =   5815
      End
      Begin VB.TextBox Text1 
         Height          =   2985
         Left            =   82
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "Form5.frx":0AA3
         Top             =   360
         Width           =   5815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5975
   End
   Begin VB.Label VersionLbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Share"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   255
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Intra"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "Form5.frx":0AF1
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
VersionLbl.Caption = "v" & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 2 Then
SSTab1.Tab = PreviousTab
ShellExecute hwnd, "open", "http://www.intrasoft-inc.com", vbNullString, vbNullString, 0
End If
End Sub
