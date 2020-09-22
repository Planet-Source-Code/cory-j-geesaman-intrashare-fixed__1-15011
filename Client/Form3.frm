VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6685E735-3BF6-11D1-A345-444553540000}#1.1#0"; "GRADIENTTITLE.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   2940
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList Flags 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0442
            Key             =   "a"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0896
            Key             =   "b"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":0CEA
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":113E
            Key             =   "d"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":1592
            Key             =   "e"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":19E6
            Key             =   "f"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":1E3A
            Key             =   "g"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":228E
            Key             =   "h"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":26E2
            Key             =   "i"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2B36
            Key             =   "j"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":2F8A
            Key             =   "k"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":33DE
            Key             =   "l"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":3832
            Key             =   "m"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":3C86
            Key             =   "n"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":40DA
            Key             =   "o"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":452E
            Key             =   "p"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":4982
            Key             =   "q"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":4DD6
            Key             =   "r"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":522A
            Key             =   "s"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":567E
            Key             =   "t"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":5AD2
            Key             =   "u"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":5F26
            Key             =   "v"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":637A
            Key             =   "w"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":67CE
            Key             =   "x"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":6C22
            Key             =   "y"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form3.frx":7076
            Key             =   "z"
         EndProperty
      EndProperty
   End
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
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   80
      TabIndex        =   10
      Top             =   3060
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   80
      TabIndex        =   6
      Top             =   0
      Width           =   2775
      Begin MSComctlLib.ImageCombo FlagList 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   2340
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
         ImageList       =   "Flags"
      End
      Begin VB.TextBox Host 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   2535
      End
      Begin VB.TextBox Password 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox Username 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1140
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Flag:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2100
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Host:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   180
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1500
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   900
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Not (InStr(1, Username.Text, "_") Or InStr(1, Password.Text, "_") Or FlagList.Text = "") Then
Form1.Socket1.HostAddress = Host.Text
Form1.Socket1.HostName = Host.Text
Form1.Socket1.Connect
Form1.Text1.Text = ""
Form1.RichTextBox1.Text = ""
Me.Hide
Else
MsgBox "You Have Used An Invalid Charactor In Your Username Or Password Or You Did Not Choose A Flag"
End If
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
For i = 1 To Flags.ListImages.Count Step 1
FlagList.ComboItems.Add , , Chr(i + 96), i, i
Next i
End Sub
