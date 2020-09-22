VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IntraShare Server v1.0"
   ClientHeight    =   9975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9975
   ScaleWidth      =   12720
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Socket 
      Index           =   0
      Left            =   3240
      Top             =   3600
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   "localhost"
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   -1  'True
      Library         =   ""
      Linger          =   0
      LocalPort       =   2223
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   2222
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin MSComctlLib.ListView AllPPL 
      Height          =   3135
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "Flags"
      SmallIcons      =   "Flags"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   5210
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "IP Address"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Current Channel"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Socket #"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox ChannelBox 
      Height          =   375
      Left            =   8280
      TabIndex        =   15
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Create Channel"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Remove Channel(s)"
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Whisper User"
      Height          =   375
      Left            =   3270
      TabIndex        =   12
      Top             =   4680
      Width           =   1655
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   2880
      Top             =   2160
   End
   Begin MSComctlLib.ListView StorageList 
      Height          =   1815
      Left            =   0
      TabIndex        =   29
      Top             =   8160
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   3201
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "IconHoldLarge"
      SmallIcons      =   "IconHold"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Owner"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Last Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox PPLCount 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   1440
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00800000&
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00000080&
      Height          =   255
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00008000&
      Height          =   255
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H00FF00FF&
      Height          =   255
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00FF0000&
      Height          =   255
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000080FF&
      Height          =   255
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   255
   End
   Begin RichTextLib.RichTextBox TBox 
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0442
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Kick User"
      Height          =   375
      Left            =   1640
      TabIndex        =   11
      Top             =   4680
      Width           =   1655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ban User"
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4680
      Width           =   1655
   End
   Begin MSComctlLib.ListView AllChannels 
      Height          =   3495
      Left            =   9840
      TabIndex        =   8
      Top             =   4680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Channel"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Owner"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kick From Channel"
      Height          =   375
      Left            =   9840
      TabIndex        =   6
      Top             =   4320
      Width           =   2895
   End
   Begin MSComctlLib.ImageList Flags 
      Left            =   3840
      Top             =   1440
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
            Picture         =   "Form1.frx":04F0
            Key             =   "a"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0944
            Key             =   "b"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D98
            Key             =   "c"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11EC
            Key             =   "d"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1640
            Key             =   "e"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A94
            Key             =   "f"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EE8
            Key             =   "g"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":233C
            Key             =   "h"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2790
            Key             =   "i"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2BE4
            Key             =   "j"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3038
            Key             =   "k"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":348C
            Key             =   "l"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":38E0
            Key             =   "m"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D34
            Key             =   "n"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4188
            Key             =   "o"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45DC
            Key             =   "p"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A30
            Key             =   "q"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E84
            Key             =   "r"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":52D8
            Key             =   "s"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":572C
            Key             =   "t"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B80
            Key             =   "u"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5FD4
            Key             =   "v"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6428
            Key             =   "w"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":687C
            Key             =   "x"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6CD0
            Key             =   "y"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7124
            Key             =   "z"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView PPl 
      Height          =   4335
      Left            =   9840
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7646
      View            =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "Flags"
      SmallIcons      =   "Flags"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox SendBox 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"Form1.frx":7578
   End
   Begin RichTextLib.RichTextBox AllChat 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":7626
   End
   Begin RichTextLib.RichTextBox ChannelChat 
      Height          =   3735
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6588
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":76D4
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send To Channel"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   4320
      Width           =   4215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Send Special"
      Height          =   375
      Left            =   4200
      TabIndex        =   30
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send To All"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Label CurrentChannel 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Users As New UserArray
Dim Usr As New User
Dim LItem As ListItem
Dim FreeSocket As Integer
Dim i As Integer

Public Function FileExists(FileName As String) As Boolean
On Error GoTo Fls
Open FileName For Input As 1
Close 1
FileExists = True
Fls:
FileExists = False
End Function

Public Function GetFreeSocket() As Integer
On Error Resume Next
    Dim nIndex As Integer

    For nIndex = 1 To FreeSocket + 1
        If Not Socket(nIndex).Connected Then
            If nIndex = FreeSocket + 1 Then FreeSocket = FreeSocket + 1
            Exit For
        End If
    Next nIndex

        Load Socket(nIndex)

    Socket(nIndex).AddressFamily = AF_INET
    Socket(nIndex).Protocol = IPPROTO_IP
    Socket(nIndex).SocketType = SOCK_STREAM
    Socket(nIndex).Blocking = False
    Socket(nIndex).AutoResolve = False
    Socket(nIndex).Listen

GetFreeSocket = nIndex
End Function

Public Function GetChannel(Index As Integer) As String
GetChannel = AllChannels.ListItems(Index).Text
End Function

Public Function ChannelList(Channel As String)
Dim TStr As String
Dim TInt As Integer
Dim TStr1 As String
Dim TInt1 As Integer
Dim TStr2 As String
Dim TInt2 As Integer
Dim ChannelDex As Integer
TStr1 = "THECHANNL" & Channel
TInt1 = Len(TStr1)
For i = 1 To AllChannels.ListItems.Count Step 1
If AllChannels.ListItems(i).Text = Channel Then
ChannelDex = i
Exit For
End If
Next i
TStr2 = "CHANNLIST"
For i = 1 To AllChannels.ListItems.Count Step 1
If i <> 2 Then
TStr2 = TStr2 & AllChannels.ListItems(i).Text & "/"
End If
Next i
TInt2 = Len(TStr2)
TStr = "THEPPLLST"
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(3) = Channel Then
TStr = TStr & AllPPL.ListItems(i).SmallIcon & AllPPL.ListItems(i).SubItems(1) & "/"
End If
Next i
TInt = Len(TStr)
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(3) = Channel Then
Socket(AllPPL.ListItems(i).SubItems(4)).Write TStr2, TInt2
Socket(AllPPL.ListItems(i).SubItems(4)).Write TStr, TInt
Users.eUser(AllPPL.ListItems(i).SubItems(4)).CurrentChannel = ChannelDex
Timer1.Interval = 6
Do Until Timer1.Interval = 1
DoEvents
Loop
Socket(AllPPL.ListItems(i).SubItems(4)).Write TStr1, TInt1
If CurrentChannel.Caption = Channel Then
Set LItem = PPl.ListItems.Add(, , , AllPPL.ListItems(i).SmallIcon, AllPPL.ListItems(i).SmallIcon)
LItem.SubItems(1) = AllPPL.ListItems(i).SubItems(1)
End If
End If
Next i
End Function

Public Function SwitchChannel(Index As Integer, Channel As String)
Dim i As Integer
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(4) = Index Then
AllPPL.ListItems(i).SubItems(3) = Channel
Exit For
End If
Next i
End Function

Private Sub AllChannels_ItemClick(ByVal Item As MSComctlLib.ListItem)
CurrentChannel.Caption = Item.Text
PPl.ListItems.Clear
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(3) = CurrentChannel.Caption Then
Set LItem = PPl.ListItems.Add(, , AllPPL.ListItems(i).SubItems(1), AllPPL.ListItems(i).Icon, AllPPL.ListItems(i).SmallIcon)
End If
Next i
End Sub

Private Sub AllChat_Change()
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelLength = 0
End Sub

Private Sub ChannelChat_Change()
ChannelChat.SelStart = Len(ChannelChat.Text)
ChannelChat.SelLength = 0
End Sub

Private Sub Command1_Click()
Dim stra As String
Dim inta As Integer
Dim i As Integer
SendBox.SelStart = 0
SendBox.SelLength = Len(SendBox.Text)
stra = SendBox.SelRTF
inta = Len(stra)
For i = 1 To AllPPL.ListItems.Count Step 1
Socket(i).Write stra, inta
Next i
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelLength = 0
AllChat.SelRTF = stra
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelLength = 0
AllChat.SelText = vbNewLine
SendBox.Text = ""
End Sub

Private Sub Command10_Click()
Set LItem = AllChannels.ListItems.Add(, , ChannelBox.Text)
LItem.SubItems(1) = "Admin"
ChannelBox.Text = ""
End Sub

Private Sub Command11_Click()
Form2.Show , Me
End Sub

Private Sub Command12_Click()
SendBox.SelColor = &HFF&
End Sub

Private Sub Command13_Click()
SendBox.SelColor = &HFF00&
End Sub

Private Sub Command14_Click()
SendBox.SelColor = &HFF0000
End Sub

Private Sub Command15_Click()
SendBox.SelColor = &HFFFF&
End Sub

Private Sub Command16_Click()
SendBox.SelColor = &H8000&
End Sub

Private Sub Command17_Click()
SendBox.SelColor = &H80&
End Sub

Private Sub Command18_Click()
SendBox.SelColor = &HFFFFFF
End Sub

Private Sub Command19_Click()
SendBox.SelColor = &HFF00FF
End Sub

Private Sub Command2_Click()
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(3) = CurrentChannel.Caption Then
Dim b As String
SendBox.SelStart = 0
SendBox.SelLength = Len(SendBox.Text)
b = SendBox.SelRTF
Socket(AllPPL.ListItems(i).SubItems(4)).Write b, Len(b)
End If
Next i
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelLength = 0
AllChat.SelRTF = b
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelLength = 0
AllChat.SelText = vbNewLine
ChannelChat.SelStart = Len(ChannelChat.Text)
ChannelChat.SelLength = 0
ChannelChat.SelRTF = b
ChannelChat.SelStart = Len(ChannelChat.Text)
ChannelChat.SelLength = 0
ChannelChat.SelText = vbNewLine
SendBox.Text = ""
End Sub

Private Sub Command20_Click()
SendBox.SelColor = &H800000
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim TVar As Integer
Dim Chan As String
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(1) = PPl.SelectedItem.Text Then
TVar = AllPPL.ListItems(i).SubItems(4)
Exit For
End If
Next i
Chan = BAN_CHANNEL
SwitchChannel TVar, Chan
ChannelList CurrentChannel.Caption
Dim TStr As String
TStr = "THEPPLLST"
Dim TInt As Integer
TInt = Len(TStr)
Socket(AllPPL.ListItems(i).SubItems(4)).Write TStr, TInt
End Sub

Private Sub Command4_Click()
SendBox.SelColor = &H0&
End Sub

Private Sub Command5_Click()
Dim a As String
Dim b As Integer
a = "TOTBANNED"
b = 9
On Error GoTo exit_sub
Open "Ban.lst" For Append As 1
Print #1, Socket(AllPPL.SelectedItem.SubItems(4)).PeerAddress
Print #1, Socket(AllPPL.SelectedItem.SubItems(4)).PeerName
Close 1
Socket(AllPPL.SelectedItem.SubItems(4)).Write a, b
Socket(AllPPL.SelectedItem.SubItems(4)).Disconnect
Socket(AllPPL.SelectedItem.SubItems(4)).Cleanup
a = AllPPL.SelectedItem.SubItems(4)
AllPPL.ListItems.Remove AllPPL.SelectedItem.Index
Users.Remove a
exit_sub:
Exit Sub
End Sub

Private Sub Command6_Click()
Dim a As String
Dim b As Integer
a = "KICKEDOUT"
b = 9
On Error GoTo exit_sub
Socket(AllPPL.SelectedItem.SubItems(4)).Write a, b
Socket(AllPPL.SelectedItem.SubItems(4)).Disconnect
Socket(AllPPL.SelectedItem.SubItems(4)).Cleanup
a = AllPPL.SelectedItem.SubItems(4)
AllPPL.ListItems.Remove AllPPL.SelectedItem.Index
Users.Remove a
exit_sub:
Exit Sub
End Sub

Private Sub Command7_Click()
Dim TStr As String
Dim TInt As Integer
SendBox.SelStart = 0
SendBox.SelLength = Len(SendBox.Text)
TStr = SendBox.SelRTF
TInt = Len(TStr)
Socket(AllPPL.SelectedItem.SubItems(4)).Write TStr, TInt
End Sub

Private Sub Command8_Click()
SendBox.SelColor = &H80FF&
End Sub

Private Sub Command9_Click()
If AllChannels.SelectedItem.Text = MAIN_CHANNEL Or AllChannels.SelectedItem.Index = 1 Or AllChannels.SelectedItem.Text = BAN_CHANNEL Or AllChannels.SelectedItem.Index = 2 Then
MsgBox "You Cannot Remove The Main Channel Or The Ban Channel", vbExclamation, "IntraShare Server"
Else
ChannelBox.Text = AllChannels.SelectedItem.Text
AllChannels.ListItems.Remove (AllChannels.SelectedItem.Index)
End If
End Sub

Private Sub Form_Load()
Me.Caption = "IntraShare Server - v" & App.Major & "." & App.Minor & App.Revision
Set LItem = AllChannels.ListItems.Add(, , MAIN_CHANNEL)
LItem.SubItems(1) = "Admin"
Set LItem = AllChannels.ListItems.Add(, , BAN_CHANNEL)
LItem.SubItems(1) = "Admin"
    Socket(0).AddressFamily = AF_INET
    Socket(0).Protocol = IPPROTO_IP
    Socket(0).SocketType = SOCK_STREAM
    Socket(0).Blocking = False
    Socket(0).AutoResolve = False
    Socket(0).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
For i = 0 To FreeSocket Step 1
Socket(i).Disconnect
Socket(i).Cleanup
Next i
End Sub

Private Sub Socket_Accept(Index As Integer, SocketId As Integer)
Socket(GetFreeSocket()).Accept = SocketId
End Sub

Private Sub Socket_Connect(Index As Integer)
If Index <> 0 Then
Dim a As String
Dim b As Integer
a = "SEND_DATA"
b = 9
Socket(Index).Write a, b
End If
End Sub

Private Sub Socket_Disconnect(Index As Integer)
Dim a As String
a = Index
Dim ItBeOk As Boolean
ItBeOk = False
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(4) = a Then
ItBeOk = True
End If
Next i
If ItBeOk = True Then
On Error Resume Next
Dim OUsr As String
OUsr = Users.eUser(Index).UserName
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(1) = Users.eUser(Index).UserName Then
AllPPL.ListItems.Remove i
Exit For
End If
Next i
For i = 1 To PPl.ListItems.Count Step 1
If PPl.ListItems(i).SubItems(1) = Users.eUser(Index).UserName Then
PPl.ListItems.Remove (i)
End If
Next i
Socket(Index).Cleanup
Dim b As String
Dim c As Integer
For i = StorageList.ListItems.Count To 1 Step -1
If StorageList.ListItems(i).SubItems(1) = Users.eUser(Index).UserName Then
StorageList.ListItems.Remove (i)
End If
Next i
b = "DTSTORAGE" & Users.eUser(Index).UserName
Users.Remove a
c = Len(b)
For i = 1 To AllPPL.ListItems.Count Step 1
Socket(AllPPL.ListItems(i).SubItems(4)).Write b, c
Next i
End If
End Sub

Private Sub Socket_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
Dim a As Integer
Dim b As String
Dim i As Integer
Dim j As Integer
Dim FlagIcon As String
Dim ItBeOk As Boolean
Dim TVar As String
a = Socket(Index).Read(b, DataLength)
If Mid(b, 1, 9) = "SENT_DATA" Then
ItBeOk = True
On Error GoTo exit_sub
Open "Ban.lst" For Input As 1
Do Until EOF(1)
Line Input #1, TVar
If TVar <> "" And TVar <> " " And (TVar = Socket(Index).PeerAddress Or TVar = Socket(Index).PeerName) Then
ItBeOk = False
End If
Loop
Close 1
If ItBeOk = True Then
b = Right(b, Len(b) - 9)
Usr.IPAddress = Socket(Index).PeerAddress
Usr.Flag = Mid(b, 1, 1)
j = 0
b = Mid(b, 2, Len(b) - 1)
TStr = ""
Dim UsrNam As String
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
j = i
Exit For
End If
Next i
UsrNam = Mid(b, 1, j - 1)
b = Mid(b, j + 1)
Usr.UserName = UsrNam
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.Path & "/Users/" & UsrNam & ".usr") = True Then
Dim Pwd As String
Open App.Path & "/Users/" & UsrNam & ".usr" For Input As 1
Input #1, Pwd
Close 1
If Pwd <> b Then
b = "BADPASSWD"
a = Len(b)
Socket(Index).Write b, a
Timer1.Interval = 10
Do Until Timer1.Interval = 1
DoEvents
Loop
Socket(Index).Disconnect
Socket(Index).Cleanup
Exit Sub
End If
Else
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
fso.createfolder (App.Path & "/Users/")
Open App.Path & "/Users/" & UsrNam & ".usr" For Output As 1
Print #1, b
Close 1
End If
Set LItem = AllPPL.ListItems.Add(, , "", Usr.Flag, Usr.Flag)
LItem.SubItems(1) = Usr.UserName
LItem.SubItems(2) = Usr.IPAddress
LItem.SubItems(3) = MAIN_CHANNEL
LItem.SubItems(4) = Index
Dim e As String
e = Index
Users.Add Usr, e
ChannelList MAIN_CHANNEL
Else
b = "TOTBANNED"
a = 9
Socket(Index).Write b, a
Exit Sub
End If
ElseIf Mid(b, 1, 9) = "CHNCHANGE" Then
Dim Hit As Boolean
b = Right(b, Len(b) - 9)
ItBeOk = True
For i = 1 To AllChannels.ListItems.Count Step 1
If AllChannels.ListItems(i).Text = b Then
ItBeOk = False
End If
Next i
If ItBeOk = True Then
Set LItem = AllChannels.ListItems.Add(, , b)
LItem.SubItems(1) = Users.eUser(Index).UserName
End If
Dim OChan As Integer
OChan = Users.eUser(Index).CurrentChannel
SwitchChannel Index, b
ChannelList b
ChannelList GetChannel(OChan)
ElseIf Mid(b, 1, 9) = "ADSTORAGE" Then
On Error Resume Next
If Users.eUser(Index).UserName <> "" Then
Dim Ob As Integer, Nb As Integer
Ob = 0
Nb = 100
Do Until Ob = Nb
Ob = Ob + 1
Loop
b = Right(b, Len(b) - 9)
TStr = ""
j = 0
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
j = j + 1
If j = 1 Then
Set LItem = StorageList.ListItems.Add(, , TStr)
LItem.SubItems(1) = Users.eUser(Index).UserName
TStr = ""
ElseIf j = 2 Then
LItem.SubItems(2) = TStr
TStr = ""
ElseIf j = 3 Then
LItem.SubItems(3) = TStr
TStr = ""
j = 0
End If
Else
TStr = TStr & Mid(b, i, 1)
End If
Next i
End If
ElseIf Mid(b, 1, 9) = "SEARCHSTR" Then
b = Right(b, Len(b) - 9)
Dim TInt As Integer
TStr = ""
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
j = i
Exit For
End If
Next i
Dim SearchBy As String
Dim SearchStr As String
Dim ByFileName As Boolean
Dim ByOwner As Boolean
Dim BySize As Boolean
Dim ByModified As Boolean
ItBeOk = False

SearchBy = Mid(b, 1, j - 1)
SearchStr = Mid(b, j + 1)
ByFileName = False
ByOwner = False
BySize = False
ByModified = False

Select Case SearchBy
Case "By Filename"
ByFileName = True
Case "By Owner"
ByOwner = True
Case "By Filesize"
BySize = True
Case "By Date Modified"
ByModified = True
Case "All Of The Above"
ByFileName = True
ByOwner = True
BySize = True
ByModified = True
End Select

For i = 1 To StorageList.ListItems.Count Step 1
ItBeOk = False
If ByFileName = True Then
If InStr(1, StorageList.ListItems(i).Text, SearchStr, vbTextCompare) Then
ItBeOk = True
End If
End If
If ByOwner = True Then
If InStr(1, StorageList.ListItems(i).SubItems(1), SearchStr, vbTextCompare) Then
ItBeOk = True
End If
End If
If BySize = True Then
If InStr(1, StorageList.ListItems(i).SubItems(2), SearchStr, vbTextCompare) Then
ItBeOk = True
End If
End If
If ByModified = True Then
If InStr(1, StorageList.ListItems(i).SubItems(4), SearchStr, vbTextCompare) Then
ItBeOk = True
End If
End If
If ItBeOk = True Then
TStr = TStr & StorageList.ListItems(i).Text & "_" & _
StorageList.ListItems(i).SubItems(1) & "_" & _
StorageList.ListItems(i).SubItems(2) & "_" & _
StorageList.ListItems(i).SubItems(3) & "_"
j = j + 1
End If
Next i
Dim TS As String
TS = "SEARCHRES" & TStr
TInt = Len(TStr)
Socket(Index).Write TS, TInt
ElseIf Mid(b, 1, 9) = "GETREMOTE" Then
b = Right(b, Len(b) - 9)
j = 0
Dim IDex As String
Dim FileName As String
Dim Owner As String
Dim FileSize As String
Dim DateModified As String
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
j = j + 1
If j = 1 Then
IDex = TStr
TStr = ""
ElseIf j = 2 Then
FileName = TStr
TStr = ""
ElseIf j = 3 Then
Owner = TStr
TStr = ""
ElseIf j = 4 Then
FileSize = TStr
TStr = ""
ElseIf j = 5 Then
DateModified = TStr
TStr = ""
End If
Else
TStr = TStr & Mid(b, i, 1)
End If
Next i
Dim tIP As String
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(1) = Owner Then
tIP = Users.eUser(AllPPL.ListItems(i).SubItems(4)).IPAddress
Exit For
End If
Next i
b = "DLREMOTEF" & IDex & "_" & tIP
a = Len(b)
Socket(Index).Write b, a
Else
TBox.Text = ""
TBox.SelStart = 0
TBox.SelLength = 0
TBox.SelRTF = b
TBox.SelStart = 0
TBox.SelLength = 0
TBox.SelRTF = Mid(Users.eUser(Index).UserName, 1, Len(Users.eUser(Index).UserName)) & ": "
TBox.SelStart = 0
TBox.SelLength = Len(TBox.Text)
b = TBox.SelRTF
ItBeOk = False
On Error Resume Next
For i = 1 To PPl.ListItems.Count Step 1
For j = 1 To Form1.PPLCount.Text Step 1
If PPl.ListItems(i).Text = Users.eUser(j).UserName Then
ItBeOk = True
End If
Next j
Next i
If ItBeOk = True Then
ChannelChat.SelLength = 0
ChannelChat.SelStart = Len(ChannelChat.Text)
ChannelChat.SelRTF = b
ChannelChat.SelStart = Len(ChannelChat.Text)
ChannelChat.SelText = vbNewLine
End If
AllChat.SelLength = 0
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelRTF = b
AllChat.SelStart = Len(AllChat.Text)
AllChat.SelText = vbNewLine
For i = 1 To AllPPL.ListItems.Count Step 1
If AllPPL.ListItems(i).SubItems(3) = AllChannels.ListItems(Users.eUser(Index).CurrentChannel).Text Then
Socket(AllPPL.ListItems(i).SubItems(4)).Write b, Len(b)
End If
Next i
Exit Sub
End If
Exit Sub
exit_sub:
Open "Ban.lst" For Output As 1
Print #1, ""
Close 1
Exit Sub
End Sub

Private Sub Timer1_Timer()
Timer1.Interval = Timer1.Interval - 1
End Sub

Private Sub Timer2_Timer()
For i = StorageList.ListItems.Count To 1 Step -1
If StorageList.ListItems(i).SubItems(1) = "" Then
StorageList.ListItems.Remove (i)
End If
Next i
End Sub
