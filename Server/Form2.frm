VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Special"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   6840
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Frame StatusBarSelector 
      Caption         =   "Status Bar Text"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   6120
      Width           =   8535
      Begin VB.TextBox StatusBarText 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   220
         Width           =   8295
      End
   End
   Begin VB.Frame MessageBoxSelector 
      Caption         =   "MessageBox Message"
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   2600
      Begin VB.TextBox MessageBoxText 
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   240
         Width           =   2375
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   8760
      _ExtentX        =   15452
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame ImageSelector 
      Caption         =   "Image Selector"
      Height          =   5895
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   5775
      Begin VB.TextBox CloseTime 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "1"
         Top             =   5505
         Width           =   5535
      End
      Begin VB.TextBox URL 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   4905
         Width           =   5535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   4305
         Width           =   1335
      End
      Begin VB.TextBox ImageLocation 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4305
         Width           =   4215
      End
      Begin VB.PictureBox ImageToSend 
         Height          =   3760
         Left            =   120
         ScaleHeight     =   3705
         ScaleMode       =   0  'User
         ScaleWidth      =   5475
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "Time(1000=1 Second) Before Letting User Close Window:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   5280
         Width           =   5535
      End
      Begin VB.Label Label2 
         Caption         =   "URL To Goto When Clicked:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4680
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Image:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   5535
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.bmp, *.jpg, *.jpeg, *.gif, *.ico, *.cur, *.tiff|*.bmp;*.jpg;*.jpeg;*.gif;*.ico;*.cur;*.tiff"
   End
   Begin VB.Frame Frame2 
      Caption         =   "Where Do You Want To Send It"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2600
      Begin VB.OptionButton StatusBar 
         Caption         =   "To The Status Bar"
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   600
         Width           =   2475
      End
      Begin VB.OptionButton MessageBox 
         Caption         =   "In A MessageBox"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "What Do You Want To Send"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2362
      Begin VB.OptionButton Text 
         Caption         =   "Text"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton ImageR 
         Caption         =   "Image"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   360
         Width           =   2235
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CD.ShowOpen
ImageLocation.Text = CD.FileName
ImageToSend.Picture = LoadPicture(ImageLocation.Text)
End Sub

Private Sub Command2_Click()
Dim MsgOut As String
If Text.Value = True Then
If MessageBox.Value = True Then
MsgOut = "DISPLAYDTMSGBOX" & MessageBoxText.Text
End If
If StatusBar.Value = True Then
MsgOut = "DISPLAYDTSTATUS" & StatusBarText.Text
End If
End If
If ImageR.Value = True Then
For i = Len(URL.Text) To 1 Step -1
If Mid(URL.Text, i, 1) = " " Then
URL.SelStart = i - 1
URL.SelLength = 1
URL.SelText = "%20"
End If
Next i
PicText = ""
For x = 1 To ImageToSend.Picture.Width Step 1
For y = 1 To ImageToSend.Picture.Height Step 1
If y = ImageToSend.Picture.Height Then
PicText = PicText & ImageToSend.Point(x, y) & "|"
Else
PicText = PicText & ImageToSend.Point(x, y) & "\"
End If
Next y
Next x
MsgOut = "DISPLAYDTAIMAGE" & CloseTime.Text & " " & URL.Text & " " & PicText
End If
Dim stra As String
Dim inta As Integer
stra = MsgOut
inta = Len(stra)
For i = 1 To Form1.AllPPL.ListItems.Count Step 1
Form1.Socket(i).Write stra, inta
Next i
End Sub

Private Sub Command3_Click()
If Text.Value = True Then
If MessageBox.Value = True Then
MsgBox MessageBoxText.Text, vbOKOnly, "Message From Server Administrator"
End If
If StatusBar.Value = True Then
StatusBar1.SimpleText = StatusBarText.Text
End If
End If
If ImageR.Value = True Then
If CloseTime.Text < 1 Then
CloseTime.Text = 1
End If
URL.Enabled = False
For i = Len(URL.Text) To 1 Step -1
If Mid(URL.Text, i, 1) = " " Then
URL.SelStart = i - 1
URL.SelLength = 1
URL.SelText = "%20"
End If
Next i
Dim PViewer As New Form3
Load PViewer
PViewer.Timer2.Interval = CloseTime.Text
PViewer.URL.Text = URL.Text
If PViewer.URL.Text <> "" Then
PViewer.Caption = PViewer.URL.Text
End If
PViewer.ImageR.Picture = ImageToSend.Picture
PViewer.ImageR.ScaleMode = PViewer.ScaleMode
PViewer.Width = (PViewer.ImageR.Picture.Width / 1.79) + 120
PViewer.Command1.Top = PViewer.ImageR.Height + 60
PViewer.Command1.Width = PViewer.ScaleWidth
PViewer.Height = (PViewer.ImageR.Height + 720)
PViewer.Show , Me
URL.Enabled = True
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub ImageR_Click()
ImageR.Enabled = False
MsgBox "This Feature Is Not Yet Available"
Exit Sub
Frame2.Enabled = False
ImageSelector.Enabled = True
MessageBoxSelector.Enabled = False
StatusBarSelector.Enabled = False
End Sub

Private Sub MessageBox_Click()
If Frame2.Enabled = True Then
MessageBoxSelector.Enabled = True
StatusBarSelector.Enabled = False
End If
End Sub

Private Sub StatusBar_Click()
If Frame2.Enabled = True Then
MessageBoxSelector.Enabled = False
StatusBarSelector.Enabled = True
End If
End Sub

Private Sub Text_Click()
Frame2.Enabled = True
ImageSelector.Enabled = False
End Sub
