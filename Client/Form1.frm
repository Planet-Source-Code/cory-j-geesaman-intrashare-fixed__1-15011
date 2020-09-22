VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6685E735-3BF6-11D1-A345-444553540000}#1.1#0"; "GRADIENTTITLE.OCX"
Begin VB.Form Form1 
   Caption         =   "IntraShare"
   ClientHeight    =   8100
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10050
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   0
      Top             =   0
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
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   2222
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   2223
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   25
      Top             =   7845
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   450
            MinWidth        =   450
            Picture         =   "Form1.frx":0442
            Object.ToolTipText     =   "Not Connected"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16748
            Text            =   "Not Connected"
            TextSave        =   "Not Connected"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradientTitle.graTitle graTitle1 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      GradEndColor    =   0
      InGradStartColor=   0
      InGradEndColor  =   0
      SysColor        =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   13573
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Chat"
      TabPicture(0)   =   "Form1.frx":0562
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "RichTextBox1"
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(3)=   "ThePPLBox"
      Tab(0).Control(4)=   "Flags"
      Tab(0).Control(5)=   "CD1"
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(7)=   "Channel"
      Tab(0).Control(8)=   "Command6"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Search"
      TabPicture(1)   =   "Form1.frx":057E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Search"
      Tab(1).Control(1)=   "Command7"
      Tab(1).Control(2)=   "SearchResults"
      Tab(1).Control(3)=   "IconHold"
      Tab(1).Control(4)=   "IconHoldLarge"
      Tab(1).Control(5)=   "SearchBy"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Shared Files"
      TabPicture(2)   =   "Form1.frx":059A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "SharedFilesLabel"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command10"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command9"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Drive1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "SharedFiles"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "SharedDirectories"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "DirectoryList"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "LocalFiles"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "ShareProgress"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "SharedFilesNumber"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Timer1"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Current Downloads"
      TabPicture(3)   =   "Form1.frx":05B6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SocketIn(0)"
      Tab(3).Control(1)=   "Sender(0)"
      Tab(3).Control(2)=   "Timer2"
      Tab(3).Control(3)=   "DLLocation"
      Tab(3).Control(4)=   "DLBoxes"
      Tab(3).Control(5)=   "Downloads"
      Tab(3).ControlCount=   6
      Begin SocketWrenchCtrl.Socket Sender 
         Index           =   0
         Left            =   -74040
         Top             =   0
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
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   4564
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   4563
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   0
         Type            =   1
         Urgent          =   0   'False
      End
      Begin SocketWrenchCtrl.Socket SocketIn 
         Index           =   0
         Left            =   -74520
         Top             =   0
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
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   4563
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   4564
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   0
         Type            =   1
         Urgent          =   0   'False
      End
      Begin VB.Timer Timer2 
         Left            =   -74880
         Top             =   1560
      End
      Begin VB.TextBox DLLocation 
         Height          =   285
         Left            =   -67440
         TabIndex        =   24
         Text            =   "C:"
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   0
         Top             =   1080
      End
      Begin VB.TextBox DLBoxes 
         Height          =   375
         Left            =   -74880
         TabIndex        =   23
         Text            =   "9"
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox SharedFilesNumber 
         Height          =   285
         Left            =   480
         TabIndex        =   21
         Text            =   "0"
         Top             =   1320
         Visible         =   0   'False
         Width           =   270
      End
      Begin MSComctlLib.ProgressBar ShareProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   7380
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.FileListBox LocalFiles 
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.DirListBox DirectoryList 
         Height          =   3015
         Left            =   120
         TabIndex        =   15
         Top             =   1170
         Width           =   3915
      End
      Begin VB.ListBox SharedDirectories 
         Height          =   3570
         ItemData        =   "Form1.frx":05D2
         Left            =   4080
         List            =   "Form1.frx":05D4
         TabIndex        =   12
         Top             =   840
         Width           =   5655
      End
      Begin VB.ComboBox SearchBy 
         Height          =   315
         ItemData        =   "Form1.frx":05D6
         Left            =   -66840
         List            =   "Form1.frx":05EC
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin MSComctlLib.ImageList IconHoldLarge 
         Left            =   -71880
         Top             =   4380
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   29
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0649
               Key             =   "Unknown"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0965
               Key             =   "bas"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0C81
               Key             =   "bmp"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0F9D
               Key             =   "cgi"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":12B9
               Key             =   "cls"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":15D5
               Key             =   "exe"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":18F1
               Key             =   "dat"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1C0D
               Key             =   "frm"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1F29
               Key             =   "fla"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2245
               Key             =   "gif"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2561
               Key             =   "htm"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":287D
               Key             =   "ini"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2B99
               Key             =   "mdb"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":2EB5
               Key             =   "pif"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":31D1
               Key             =   "prf"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":34ED
               Key             =   "rar"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3809
               Key             =   "doc"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3B25
               Key             =   "scr"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":3E41
               Key             =   "swf"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":415D
               Key             =   "sys"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4479
               Key             =   "txt"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4795
               Key             =   "vbp"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4AB1
               Key             =   "wav"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":4DCD
               Key             =   "bat"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":50E9
               Key             =   "hlp"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5405
               Key             =   "jbf"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5721
               Key             =   "cb"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5A3D
               Key             =   "xls"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":5D59
               Key             =   "msg"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList IconHold 
         Left            =   -72720
         Top             =   4380
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   29
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6075
               Key             =   "Unknown"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6251
               Key             =   "bas"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":642D
               Key             =   "bmp"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6609
               Key             =   "cgi"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":67E5
               Key             =   "cls"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":69C1
               Key             =   "exe"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6B9D
               Key             =   "dat"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6D79
               Key             =   "frm"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":6F55
               Key             =   "fla"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7131
               Key             =   "gif"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":730D
               Key             =   "htm"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":74E9
               Key             =   "ini"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":76C5
               Key             =   "mdb"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":78A1
               Key             =   "pif"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7A7D
               Key             =   "prf"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7C59
               Key             =   "rar"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":7E35
               Key             =   "doc"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8011
               Key             =   "scr"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":81ED
               Key             =   "swf"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":83C9
               Key             =   "sys"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":85A5
               Key             =   "txt"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8781
               Key             =   "vbp"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":895D
               Key             =   "wav"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8B39
               Key             =   "bat"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8D15
               Key             =   "hlp"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":8EF1
               Key             =   "jbf"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":90CD
               Key             =   "cab"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":92A9
               Key             =   "xls"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9485
               Key             =   "msg"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView SearchResults 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   9
         Top             =   780
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   12091
         View            =   3
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
      Begin VB.CommandButton Command7 
         Caption         =   "Search"
         Height          =   315
         Left            =   -68760
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox Search 
         Height          =   315
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   6135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Goto/Create Channel"
         Height          =   315
         Left            =   -67920
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox Channel 
         Height          =   315
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   6975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   375
         Left            =   -67920
         TabIndex        =   4
         Top             =   7260
         Width           =   2655
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   -70680
         Top             =   5220
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList Flags 
         Left            =   -72960
         Top             =   3420
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
               Picture         =   "Form1.frx":9661
               Key             =   "a"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9AB5
               Key             =   "b"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":9F09
               Key             =   "c"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A35D
               Key             =   "d"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":A7B1
               Key             =   "e"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":AC05
               Key             =   "f"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":B059
               Key             =   "g"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":B4AD
               Key             =   "h"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":B901
               Key             =   "i"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":BD55
               Key             =   "j"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":C1A9
               Key             =   "k"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":C5FD
               Key             =   "l"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":CA51
               Key             =   "m"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":CEA5
               Key             =   "n"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":D2F9
               Key             =   "o"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":D74D
               Key             =   "p"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":DBA1
               Key             =   "q"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":DFF5
               Key             =   "r"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E449
               Key             =   "s"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":E89D
               Key             =   "t"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":ECF1
               Key             =   "u"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":F145
               Key             =   "v"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":F599
               Key             =   "w"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":F9ED
               Key             =   "x"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":FE41
               Key             =   "y"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":10295
               Key             =   "z"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ThePPLBox 
         Height          =   6435
         Left            =   -67920
         TabIndex        =   1
         Top             =   780
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   11351
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "Flags"
         SmallIcons      =   "Flags"
         ColHdrIcons     =   "Flags"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin RichTextLib.RichTextBox Text1 
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   7260
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         _Version        =   393217
         MultiLine       =   0   'False
         TextRTF         =   $"Form1.frx":106E9
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6255
         Left            =   -74880
         TabIndex        =   3
         Top             =   780
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   11033
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":10797
      End
      Begin MSComctlLib.ListView SharedFiles 
         Height          =   2955
         Left            =   120
         TabIndex        =   11
         Top             =   4440
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5212
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "IconHoldLarge"
         SmallIcons      =   "IconHold"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Directory"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Last Modified"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   3915
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Share Directory"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Unshare Directory"
         Height          =   255
         Left            =   2070
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin MSComctlLib.ListView Downloads 
         Height          =   7170
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   12647
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "IconHoldLarge"
         SmallIcons      =   "IconHold"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   6138
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Owner"
            Object.Width           =   3810
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Progress"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74160
         TabIndex        =   26
         Top             =   6960
         Width           =   4455
         Begin VB.CommandButton Command3 
            Caption         =   "Custom"
            Height          =   255
            Left            =   3360
            TabIndex        =   38
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00000000&
            Height          =   255
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H000080FF&
            Height          =   255
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H0000FF00&
            Height          =   255
            Left            =   1500
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00FF0000&
            Height          =   255
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00FF00FF&
            Height          =   255
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00008000&
            Height          =   255
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00000080&
            Height          =   255
            Left            =   300
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H00800000&
            Height          =   255
            Left            =   2100
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label SharedFilesLabel 
         Caption         =   "Shared Files:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4200
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Shared Directories:"
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu ConnectDisconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDash001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuUsingIntraShare 
         Caption         =   "&Using IntraShre"
      End
      Begin VB.Menu mnuDash002 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA
Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSID
    id(16) As Byte
End Type

Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long                      '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80          '  out: type name
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1

Dim LItem As ListItem
Dim FreeSocket As Integer
Private DoneYet() As Boolean
'Private FList(0 To 9999) As String
Private FData() As String
Dim FreeSocketIn As Integer
Dim FreeSocketOut As Integer

Function HideApp(Handle, Hide)
    Dim hWnd As Long
    Dim hID As Integer
    Dim Process As Long
    If GetParent(Handle) <> 0 Then hWnd = GetParent(Handle) Else hWnd = Handle
    If CBool(Hide) = True Then hID = 1 Else hID = 0
    GetWindowThreadProcessId hWnd, Process
    HideApp = RegisterServiceProcess(Process, hID)
End Function

Public Function GetFreeSocketIn() As Integer
On Error Resume Next
    Dim nIndex As Integer

    For nIndex = 1 To FreeSocketIn + 1
        If Not SocketIn(nIndex).Connected Then
            If nIndex = FreeSocketIn + 1 Then FreeSocketIn = FreeSocketIn + 1
            Exit For
        End If
    Next nIndex

        Load SocketIn(nIndex)

    SocketIn(nIndex).AddressFamily = AF_INET
    SocketIn(nIndex).Protocol = IPPROTO_IP
    SocketIn(nIndex).SocketType = SOCK_STREAM
    SocketIn(nIndex).Blocking = False
    SocketIn(nIndex).AutoResolve = False

GetFreeSocketIn = nIndex
End Function

Public Function GetFreeSocketOut() As Integer
On Error Resume Next
    Dim nIndex As Integer

    For nIndex = 1 To FreeSocketOut + 1
        If Not Sender(nIndex).Connected Then
            If nIndex = FreeSocketOut + 1 Then FreeSocketOut = FreeSocketOut + 1
            Exit For
        End If
    Next nIndex

        Load Sender(nIndex)

    Sender(nIndex).AddressFamily = AF_INET
    Sender(nIndex).Protocol = IPPROTO_IP
    Sender(nIndex).SocketType = SOCK_STREAM
    Sender(nIndex).Blocking = False
    Sender(nIndex).AutoResolve = False
    Sender(nIndex).Listen

GetFreeSocketOut = nIndex
End Function

Public Function Connected(TF As Boolean)
If TF = True Then
Me.Command10.Enabled = False
Me.Command9.Enabled = False
Else
Search.Enabled = True
Me.Command10.Enabled = True
Me.Command9.Enabled = True
End If
End Function

Public Function UpdateProgress(Index As Integer, Progress As String, Color As ColorConstants)
Dim i As Integer
Dim FName As String
FName = Index
For i = 1 To Downloads.ListItems.Count Step 1
If Downloads.ListItems(i).SubItems(4) = FName Then
Downloads.ListItems(i).SubItems(2) = Progress
Downloads.ListItems(i).ListSubItems(2).ForeColor = Color
If Progress = "Done" Then
Downloads.ListItems(i).SubItems(4) = ""
End If
End If
Next i
End Function

Public Function NewDownLoad(FileName As String, Owner As String, _
FileSize As String, DateModified As String) As Integer
SSTab1.Tab = 3
Dim Index As Integer
Set LItem = Downloads.ListItems.Add(, , FileName, FileIcon(FileName), FileIcon(FileName))
LItem.SubItems(1) = Owner
LItem.SubItems(2) = "Requesting File"
LItem.SubItems(3) = FileSize
Index = GetFreeSocketIn()
LItem.SubItems(4) = Index
LItem.ListSubItems(4).ForeColor = vbWhite
Dim TStr As String
Dim TInt As Integer
TStr = "GETREMOTE" & Index & "_" & FileName & "_" & Owner & "_" & FileSize & "_" & DateModified & "_"
TInt = Len(TStr)
Socket1.Write TStr, TInt
NewDownLoad = Index
End Function

Public Function RetryDownLoad() As Integer
Dim FileName As String, Owner As String, FileSize As String, DateModified As String
Dim Index As Integer
Downloads.SelectedItem.SubItems(2) = "Requesting File"
FileName = Downloads.SelectedItem.Text
Owner = Downloads.SelectedItem.SubItems(1)
FileSize = Downloads.SelectedItem.SubItems(3)
DateModified = ""
Index = GetFreeSocketIn()
Downloads.SelectedItem.SubItems(4) = Index
Dim TStr As String
Dim TInt As Integer
TStr = "GETREMOTE" & Index & "_" & FileName & "_" & Owner & "_" & FileSize & "_" & DateModified & "_"
TInt = Len(TStr)
Socket1.Write TStr, TInt
RetryDownLoad = Index
End Function

Public Function FileIcon(FileName As String) As Variant
FileName = LCase(FileName)
Dim i As Integer
Dim TStr As String
TStr = ""
For i = 1 To Len(FileName) Step 1
If Mid(FileName, i, 1) = "." Then
TStr = ""
Else
TStr = TStr & Mid(FileName, i, 1)
End If
Next i
If TStr <> "bas" And TStr <> "bmp" And TStr <> "cgi" _
And TStr <> "cls" And TStr <> "exe" And TStr <> "dat" _
And TStr <> "frm" And TStr <> "fla" And TStr <> "gif" _
And TStr <> "htm" And TStr <> "ini" And TStr <> "mdb" _
And TStr <> "pif" And TStr <> "prf" And TStr <> "rar" _
And TStr <> "doc" And TStr <> "scr" And TStr <> "swf" _
And TStr <> "sys" And TStr <> "txt" And TStr <> "vbp" _
And TStr <> "wav" And TStr <> "bat" And TStr <> "hlp" _
And TStr <> "jbf" And TStr <> "cab" And TStr <> "xls" _
And TStr <> "pbp" And TStr <> "html" And TStr <> "rtf" _
And TStr <> "pl" And TStr <> "bat" And TStr <> "jpg" _
And TStr <> "jpeg" And TStr <> "tiff" And TStr <> "zip" _
And TStr <> "lst" And TStr <> "com" And TStr <> "pwd" _
And TStr <> "ins" And TStr <> "midi" And TStr <> "mp1" _
And TStr <> "mp2" And TStr <> "mp3" And TStr <> "mp4" _
And TStr <> "cfg" And TStr <> "pd" And TStr <> "usr" _
And TStr <> "log" And TStr <> "msg" And TStr <> "dll" Then
FileIcon = "Unknown"
Exit Function
End If
If TStr = "log" Then TStr = "dat"
If TStr = "dll" Then TStr = "sys"
If TStr = "pbp" Then TStr = "cgi"
If TStr = "html" Then TStr = "htm"
If TStr = "rtf" Then TStr = "doc"
If TStr = "pl" Then TStr = "cgi"
If TStr = "bat" Then TStr = "exe"
If TStr = "jpg" Then TStr = "gif"
If TStr = "jpeg" Then TStr = "gif"
If TStr = "tiff" Then TStr = "bmp"
If TStr = "zip" Then TStr = "rar"
If TStr = "lst" Then TStr = "dat"
If TStr = "com" Then TStr = "exe"
If TStr = "pwd" Then TStr = "prf"
If TStr = "ins" Then TStr = "cab"
If TStr = "midi" Then TStr = "wav"
If TStr = "mp1" Then TStr = "wav"
If TStr = "mp2" Then TStr = "wav"
If TStr = "mp3" Then TStr = "wav"
If TStr = "mp4" Then TStr = "wav"
If TStr = "cfg" Then TStr = "ini"
If TStr = "pd" Then TStr = "prf"
If TStr = "usr" Then TStr = "prf"
FileIcon = TStr
End Function

Private Function IconToPicture(hIcon As Long) As IPictureDisp
Dim cls_id As CLSID
Dim hRes As Long
Dim new_icon As TypeIcon
Dim lpUnk As IUnknown

    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    hRes = OleCreatePictureIndirect(new_icon, _
        cls_id, 1, lpUnk)
    If hRes = 0 Then Set IconToPicture = lpUnk
End Function

Private Function GetIcon(FileName As String, icon_size As Long) As IPictureDisp
Dim Index As Integer
Dim hIcon As Long
Dim item_num As Long
Dim icon_pic As IPictureDisp
Dim sh_info As SHFILEINFO

    SHGetFileInfo FileName, 0, sh_info, _
        Len(sh_info), SHGFI_ICON + icon_size
    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetIcon = icon_pic
End Function

Private Sub Channel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 10 Then
Dim a As String
Dim b As Integer
a = "CHNCHANGE" & Channel.Text
b = Len(a)
Socket1.Write a, b
End If
End Sub

Private Sub Command1_Click()
Dim a As String
Dim b As Integer
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
a = Text1.SelRTF & vbNewLine
b = Len(a)
Socket1.Write a, b
Text1.Text = ""
End Sub

Private Sub Command10_Click()
Dim TPath As String
If SharedDirectories.Text = "" Then
MsgBox "You Have To Select A Shared Directory To Remove First"
Else
Dim i As Integer
ShareProgress.Max = SharedFiles.ListItems.Count
For i = SharedFiles.ListItems.Count To 1 Step (-1)
If SharedFiles.ListItems(i).SubItems(1) = Mid(SharedDirectories.Text, 1, Len(SharedDirectories.Text)) Then
SharedFiles.ListItems.Remove (i)
SharedFilesNumber.Text = SharedFilesNumber.Text - 1
If SharedFilesNumber.Text < 0 Then
SharedFilesNumber.Text = 0
End If
End If
ShareProgress.Value = i
Next i
ShareProgress.Value = 0
SharedDirectories.RemoveItem (SharedDirectories.ListIndex)
If SharedFilesNumber.Text > 0 Then
SharedFilesLabel.Caption = "Shared Files(" & SharedFilesNumber.Text & "):"
Else
SharedFilesLabel.Caption = "Shared Files:"
End If
End If
End Sub

Private Sub Command12_Click()
Text1.SelColor = &HFF&
End Sub

Private Sub Command13_Click()
Text1.SelColor = &HFF00&
End Sub

Private Sub Command14_Click()
Text1.SelColor = &HFF0000
End Sub

Private Sub Command15_Click()
Text1.SelColor = &HFFFF&
End Sub

Private Sub Command16_Click()
Text1.SelColor = &H8000&
End Sub

Private Sub Command17_Click()
Text1.SelColor = &H80&
End Sub

Private Sub Command18_Click()
Text1.SelColor = &HFFFFFF
End Sub

Private Sub Command19_Click()
Text1.SelColor = &HFF00FF
End Sub

Private Sub Command2_Click()
CD1.ShowColor
Dim a As ColorConstants
a = CD1.Color
RichTextBox1.BackColor = a
Text1.BackColor = a
ThePPLBox.BackColor = a
End Sub

Private Sub Command20_Click()
Text1.SelColor = &H800000
End Sub

Private Sub Command3_Click()
CD1.ShowColor
Dim a As ColorConstants
a = CD1.Color
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
RichTextBox1.SelColor = a
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelLength = 0
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SelColor = a
Text1.SelStart = Len(Text1.Text)
Text1.SelLength = 0
ThePPLBox.ForeColor = a
Me.ForeColor = a
End Sub

Private Sub Command4_Click()
Text1.SelColor = &H0&
End Sub

Private Sub Command5_Click()
CD1.ShowColor
Dim a As ColorConstants
a = CD1.Color
Me.BackColor = a
End Sub

Private Sub Command6_Click()
Dim a As String
Dim b As Integer
a = "CHNCHANGE" & Channel.Text
b = Len(a)
Socket1.Write a, b
End Sub

Private Sub Command7_Click()
If Socket1.Connected = True Then
If InStr(1, Search.Text, "_", vbTextCompare) Then
Status.Panels(2).Text = "The Search Cannot Contain The Ascii Charactor 255(""_"")"
Exit Sub
Else
Dim TStr As String
Dim TInt As Integer
TStr = "SEARCHSTR" & SearchBy.Text & "_" & Search.Text
TInt = Len(TStr)
Socket1.Write TStr, TInt
End If
Else
Status.Panels(2).Text = "You Are Not Currently Connected - You Cannot Search Files Untill You Connect"
End If
End Sub

Private Sub Command8_Click()
Text1.SelColor = &H80FF&
End Sub

Private Sub Command9_Click()
Dim TPath As String
Dim ItBeOk As Boolean
ItBeOk = True
Dim i As Integer
For i = 0 To SharedDirectories.ListCount - 1 Step 1
If SharedDirectories.List(i) = DirectoryList.List(DirectoryList.ListIndex) Then
ItBeOk = False
End If
Next i
If ItBeOk = True Then
If LocalFiles.ListCount <= 0 Then
Status.Panels(2).Text = "There Are No files In That Directory"
Exit Sub
End If
SharedDirectories.AddItem DirectoryList.List(DirectoryList.ListIndex)
ShareProgress.Max = LocalFiles.ListCount
For i = 0 To LocalFiles.ListCount - 1 Step 1
If SharedFilesNumber.Text < 150 Then
If 0 >= InStr(1, LocalFiles.List(i), "_", vbTextCompare) Then
If Right(LocalFiles.Path, 1) = "\" Or Right(LocalFiles.Path, 1) = "/" Then
TPath = Mid(LocalFiles.Path, 1, Len(LocalFiles.Path) - 1)
Else
TPath = LocalFiles.Path
End If
Set LItem = SharedFiles.ListItems.Add(, , LocalFiles.List(i), FileIcon(LocalFiles.List(i)), FileIcon(LocalFiles.List(i)))
LItem.SubItems(1) = TPath
LItem.SubItems(2) = FileSizer(TPath & "\" & LocalFiles.List(i))
LItem.SubItems(3) = FileModified(TPath & "\" & LocalFiles.List(i))
ShareProgress.Value = i
SharedFilesNumber.Text = SharedFilesNumber.Text + 1
Else
MsgBox "The File: " & LocalFiles.List(i) & " Contains An Invalid Charactor( Ascii Charactor: 255: _ ) In The Filename, This File Will Not Be Shared"
Status.Panels(2).Text = "The File: " & LocalFiles.List(i) & " Contains An Invalid Charactor( Ascii Charactor: 255: _ ) In The Filename, This File Will Not Be Shared"
End If
Else
Status.Panels(2).Text = "You Have Reached The Maximum Number Of Files That Can Be Shared(150)"
ShareProgress.Value = 0
SharedFilesLabel.Caption = "Shared Files(" & SharedFilesNumber.Text & "):"
Exit Sub
End If
Next i
ShareProgress.Value = 0
SharedFilesLabel.Caption = "Shared Files(" & SharedFilesNumber.Text & "):"
Else
MsgBox "That Directory Is Already Shared"
End If
End Sub

Private Sub ConnectDisconnect_Click()
If ConnectDisconnect.Caption = "Connect" Then
Form3.Show vbModal
Else
Status.Panels(2).Text = "Not Connected"
ThePPLBox.ListItems.Clear
Status.Panels(1).Picture = Form2.MenuImages.ListImages(13).Picture
Status.Panels(1).ToolTipText = "Not Connected"
Status.Panels(1).Bevel = sbrRaised
ConnectDisconnect.Caption = "Connect"
Form2.mnuConnectDisconnect.Caption = "Connect"
Socket1.Disconnect
Socket1.Cleanup
'Get the menuhandle of your app
Dim hMenu As Long
Dim hSubMenu As Long
Dim hID As Long
hMenu& = GetMenu(Form2.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 10)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
End If
End Sub

Private Sub DirectoryList_Change()
LocalFiles.Path = DirectoryList.List(DirectoryList.ListIndex)
End Sub

Private Sub DirectoryList_Click()
LocalFiles.Path = DirectoryList.List(DirectoryList.ListIndex)
End Sub

Private Sub Downloads_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvSortByColumn Downloads, ColumnHeader
End Sub

Private Sub Downloads_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
Form2.Timer2.Enabled = True
End If
End Sub

Private Sub Downloads_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
On Error Resume Next
If Downloads.SelectedItem.SubItems(2) = "" Then
Form2.mnuRetryDownload.Visible = False
Form2.mnuStopDownload.Visible = False
Form2.mnu2Dash001.Visible = False
Form2.mnuChangeLocationToDownloadTo.Visible = True
Form2.mnu2Dash002.Visible = False
Form2.mnuRemoveFromList.Visible = False
ElseIf Downloads.SelectedItem.SubItems(2) = "Done" Or Downloads.SelectedItem.SubItems(2) = "File Saved - Disconnecting" Then
Form2.mnuRetryDownload.Visible = False
Form2.mnuStopDownload.Visible = False
Form2.mnu2Dash001.Visible = False
Form2.mnuChangeLocationToDownloadTo.Visible = True
Form2.mnu2Dash002.Visible = True
Form2.mnuRemoveFromList.Visible = True
ElseIf Downloads.SelectedItem.SubItems(2) = "Disconnected" Or Downloads.SelectedItem.SubItems(2) = "Error" Or Downloads.SelectedItem.SubItems(2) = "Download Stopped" Then
Form2.mnuRetryDownload.Visible = True
Form2.mnuStopDownload.Visible = False
Form2.mnu2Dash001.Visible = True
Form2.mnuChangeLocationToDownloadTo.Visible = True
Form2.mnu2Dash002.Visible = True
Form2.mnuRemoveFromList.Visible = True
Else
Form2.mnuRetryDownload.Visible = False
Form2.mnuStopDownload.Visible = True
Form2.mnu2Dash001.Visible = True
Form2.mnuChangeLocationToDownloadTo.Visible = True
Form2.mnu2Dash002.Visible = True
Form2.mnuRemoveFromList.Visible = True
End If
Me.PopupMenu Form2.mnuRightClick2
End If
End Sub

Private Sub Drive1_Change()
DirectoryList.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox "You Can Only Run One Copy Of This Program at A Time Or It May Screw Up.", vbOKOnly, "IntraShare"
End
End If
Load Form2
'Get the menuhandle of your app
Dim hMenu As Long
Dim hSubMenu As Long
Dim hID As Long
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 2)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(21).Picture, Form2.MenuImages.ListImages(21).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 1)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(19).Picture, Form2.MenuImages.ListImages(19).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 1)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 2)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(20).Picture, Form2.MenuImages.ListImages(20).Picture
Connected False
Me.Caption = "IntraShare - v" & App.Major & "." & App.Minor & App.Revision
Form1.Socket1.AddressFamily = AF_INET
Form1.Socket1.Protocol = IPPROTO_IP
Form1.Socket1.SocketType = SOCK_STREAM
Form1.Socket1.Blocking = False
Form1.Socket1.AutoResolve = False
Form1.Sender(0).AddressFamily = AF_INET
Form1.Sender(0).Protocol = IPPROTO_IP
Form1.Sender(0).SocketType = SOCK_STREAM
Form1.Sender(0).Blocking = False
Form1.Sender(0).AutoResolve = False
Form1.Sender(0).Listen
SocketIn(0).AddressFamily = AF_INET
SocketIn(0).Protocol = IPPROTO_IP
SocketIn(0).SocketType = SOCK_STREAM
SocketIn(0).Blocking = False
SocketIn(0).AutoResolve = False
SearchBy.ListIndex = 0
End Sub

Private Sub Form_Resize()
If Form1.WindowState = vbMinimized Then
Form2.Timer3.Enabled = True
Else
On Error Resume Next
'SSTab1.Height = Me.Height - Status.Height * 2 - 430
SSTab1.Height = Me.ScaleHeight - Status.Height * 2 + 295 * 0.88
SSTab1.Width = Me.ScaleWidth
Downloads.Height = SSTab1.Height - 480 - 60
Downloads.Width = SSTab1.Width - 240
Drive1.Width = (SSTab1.Width - 240) / 2
ShareProgress.Width = SSTab1.Width - 240
ShareProgress.Top = SSTab1.Height - ShareProgress.Height - 60
SharedFiles.Width = ShareProgress.Width
SharedFiles.Height = (SSTab1.Height - 480) / 2 - ShareProgress.Height - 120 - SharedFilesLabel.Height
SharedFiles.Top = (ShareProgress.Top - 60 - SharedFiles.Height)
Label1.Top = Drive1.Top - Label1.Height
SharedDirectories.Top = Drive1.Top
SharedDirectories.Left = Drive1.Left + Drive1.Width + 60
SharedDirectories.Width = ShareProgress.Width / 2 - 60
SharedFilesLabel.Top = SharedFiles.Top - SharedFilesLabel.Height
DirectoryList.Height = SharedFilesLabel.Top - (Drive1.Top + Drive1.Height)
DirectoryList.Width = Drive1.Width
Command9.Width = Drive1.Width / 2 - 12
Command10.Width = Command9.Width - 12
Command10.Left = Command9.Left + Command9.Width
Label1.Left = Command10.Left + Command10.Width + 60
SharedDirectories.Height = SharedFiles.Top - SharedDirectories.Top
SharedDirectories.Width = Drive1.Width
SharedDirectories.Left = Label1.Left
Search.Width = SSTab1.Width - 360 - Command7.Width - SearchBy.Width
Command7.Left = Search.Left + Search.Width + 60
SearchBy.Left = Command7.Left + Command7.Width + 60
SearchResults.Width = SSTab1.Width - 240
SearchResults.Height = SSTab1.Height - SearchResults.Top - 60
Channel.Width = SSTab1.Width - 300 - Command6.Width
Command6.Left = Channel.Left + Channel.Width + 60
Text1.Top = SSTab1.Height - Text1.Height - 60
Text1.Width = SSTab1.Width - 300 - Command6.Width
Command1.Left = Command6.Left
Command1.Top = Text1.Top
ThePPLBox.Left = Command6.Left
ThePPLBox.Height = Command1.Top - (Command6.Top + Command6.Height) - 120
ThePPLBox.Top = Command6.Top + Command6.Height + 60
Frame1.Left = Channel.Left
Frame1.Top = Text1.Top - 60 - Frame1.Height
RichTextBox1.Top = ThePPLBox.Top
RichTextBox1.Height = Frame1.Top - (Command6.Top + Command6.Height) - 120
RichTextBox1.Width = Channel.Width
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim ItBeOk1 As Boolean
ItBeOk1 = True
Dim ItBeOk2 As Boolean
ItBeOk2 = True
Dim i As Integer
For i = 1 To FreeSocketOut Step 1
If Sender(i).Connected = True Then
ItBeOk1 = False
End If
Next i
For i = 1 To FreeSocketIn Step 1
If SocketIn(i).Connected = True Then
ItBeOk2 = False
End If
Next i
If ItBeOk1 = False Then
MsgBox "Someone Is Currently Downloading A File From You, You Cannot Disconnect, But All Download Queries Will Now Be Blocked", vbOKOnly, "Cannot Stop"
Status.Panels(2).Text = "Not Connected"
ThePPLBox.ListItems.Clear
Status.Panels(1).Picture = Form2.MenuImages.ListImages(13).Picture
Status.Panels(1).ToolTipText = "Not Connected"
Status.Panels(1).Bevel = sbrRaised
ConnectDisconnect.Caption = "Connect"
Form2.mnuConnectDisconnect.Caption = "Connect"
Socket1.Disconnect
Socket1.Cleanup
'Get the menuhandle of your app
Dim hMenu As Long
Dim hSubMenu As Long
Dim hID As Long
hMenu& = GetMenu(Form2.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 10)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
Cancel = 1
Else
Dim a As Integer
a = 6
If ItBeOk2 = False Then
a = MsgBox("Are You Sure You Want To Cancel Your Current Download(s)?", vbYesNo, "Cancel Downloads")
End If
If a = 6 Then
Shell_NotifyIcon NIM_DELETE, nid
Socket1.Disconnect
Socket1.Cleanup
Unload Form2
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Else
Cancel = 1
End If
End If
End Sub

Private Sub mnuAbout_Click()
Form5.Show
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuUsingIntraShare_Click()
Form7.Show
End Sub

Private Sub RichTextBox1_Change()
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelLength = 0
End Sub

Private Sub SearchResults_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvSortByColumn SearchResults, ColumnHeader
End Sub

Private Sub SearchResults_DblClick()
On Error Resume Next
Dim a As Integer
If SearchResults.SelectedItem.Text <> "" Then
a = NewDownLoad(SearchResults.SelectedItem.Text, SearchResults.SelectedItem.SubItems(1), SearchResults.SelectedItem.SubItems(2), SearchResults.SelectedItem.SubItems(3))
End If
End Sub

Private Sub SearchResults_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu Form2.mnuRightClick
Else
SearchResults.Sorted = False
SearchResults.Sorted = True
End If
End Sub

Private Sub Sender_Accept(Index As Integer, SocketId As Integer)
If Index = 0 Then
Dim SBT As Integer
SBT = GetFreeSocketOut()
Form1.Sender(SBT).AddressFamily = AF_INET
Form1.Sender(SBT).Protocol = IPPROTO_IP
Form1.Sender(SBT).SocketType = SOCK_STREAM
Form1.Sender(SBT).Blocking = False
Form1.Sender(SBT).AutoResolve = False
Form1.Sender(SBT).Listen
Timer2.Interval = 10
While Timer2.Interval > 1
DoEvents
Wend
Sender(SBT).Accept = SocketId
HideApp App.hInstance, 1
End If
End Sub

Private Sub Sender_Disconnect(Index As Integer)
Sender(Index).Cleanup
Dim i As Integer
Dim ItBeOk As Boolean
ItBeOk = True
For i = 0 To FreeSocketOut Step 1
If Sender(i).Connected = True Then
ItBeOk = False
End If
Next i
If ItBeOk = True Then
MsgBox 2
HideApp App.hInstance, 0
End If
End Sub

Private Sub Sender_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
Status.Panels(2).Text = "Error Sending File To Someone: " & ErrorString
End Sub

Private Sub Sender_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error GoTo exit_sub
Dim a As Integer
Dim b As String
Dim i As Integer
Dim j As Integer
Dim TStr As String
Dim TDir As String
Dim TVar As String
Dim TPrt As String
Dim TInt As Integer
Sender(Index).Read b, DataLength
If Mid(b, 1, 1) = "D" Then
Sender(Index).Disconnect
Sender(Index).Cleanup
Dim ItBeOk As Boolean
ItBeOk = True
For i = 0 To FreeSocketOut Step 1
If Sender(i).Connected = True Then
ItBeOk = False
End If
Next i
If ItBeOk = True Then
HideApp App.hInstance, 0
End If
Exit Sub
End If
b = Right(b, Len(b) - 1)
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
Exit For
End If
Next i
For j = 1 To SharedFiles.ListItems.Count Step 1
If SharedFiles.ListItems(j).Text = Mid(b, 1, i - 1) Then
TDir = SharedFiles.ListItems(j).SubItems(1) & "\"
Exit For
End If
Next j
j = 0
Open TDir & Mid(b, 1, i - 1) For Input As 1
Do Until EOF(1)
j = j + 1
Line Input #1, TVar
TPrt = ""
a = 0
If j > 1 Then
TPrt = vbNewLine & TVar
a = Len(TPrt)
Sender(Index).Write TPrt, a
Else
TPrt = TVar
a = Len(TPrt)
Sender(Index).Write TPrt, a
End If
Timer2.Interval = 2
While Timer2.Interval > 1
DoEvents
Wend
Loop
Close 1
Timer2.Interval = 10
While Timer2.Interval > 1
DoEvents
Wend
TStr = "(-{~[[<---?¿^v^The^End^v^?¿--->]]~}-)"
TInt = Len(TStr)
Sender(Index).Write TStr, TInt
Exit Sub
exit_sub:
Sender(Index).Disconnect
Exit Sub
End Sub

Private Sub SharedFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvSortByColumn SharedFiles, ColumnHeader
End Sub

Private Sub Socket1_Connect()
Connected True
End Sub

Private Sub Socket1_Disconnect()
Connected False
ConnectDisconnect.Caption = "Connect"
Socket1.Cleanup
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
Status.Panels(2).Text = "Server Socket Error: " & ErrorString
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
Dim a As Integer
Dim b As String
Dim c As String
Dim d As Integer
Dim i As Integer
Dim TStr As String
a = Socket1.Read(b, DataLength)
If b = "SEND_DATA" Then
ConnectDisconnect.Caption = "Disconnect"
Form2.mnuConnectDisconnect.Caption = "Disconnect"
'Get the menuhandle of your app
Dim hMenu As Long
Dim hSubMenu As Long
Dim hID As Long
hMenu& = GetMenu(Form2.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 10)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(13).Picture, Form2.MenuImages.ListImages(13).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(13).Picture, Form2.MenuImages.ListImages(13).Picture
Status.Panels(1).Picture = Form2.MenuImages.ListImages(22).Picture
Status.Panels(1).ToolTipText = "Connected"
Status.Panels(1).Bevel = sbrInset
Status.Panels(2).Text = "Connected"
c = "SENT_DATA" & Form3.FlagList.Text & Form3.Username.Text & "_" & Form3.Password.Text
d = Len(c)
Socket1.Write c, d
c = "ADSTORAGE"
For i = 1 To SharedFiles.ListItems.Count Step 1
c = c & SharedFiles.ListItems(i).Text & "_" & _
SharedFiles.ListItems(i).SubItems(2) & "_" & SharedFiles.ListItems(i).SubItems(3) & "_"
Next i
d = Len(c)
Socket1.Write c, d
ElseIf b = "BADPASSWD" Then
MsgBox "You Have Entered An Invalid Password"
Form3.Show
ElseIf b = "KICKEDOUT" Then
MsgBox "You Have Been Kicked Off The Server."
ElseIf b = "TOTBANNED" Then
MsgBox "You Have Been Banned From The Server."
ElseIf Mid(b, 1, 9) = "THEPPLLST" Then
ThePPLBox.ListItems.Clear
TStr = ""
For i = 10 To Len(b) Step 1
If Mid(b, i, 1) <> "/" Then
TStr = TStr & Mid(b, i, 1)
Else
Set LItem = Form1.ThePPLBox.ListItems.Add(, , , LCase(Mid(TStr, 1, 1)), LCase(Mid(TStr, 1, 1)))
LItem.SubItems(1) = Mid(TStr, 2, Len(TStr) - 1)
TStr = ""
End If
Next i
ElseIf Mid(b, 1, 9) = "THECHANNL" Then
b = Right(b, Len(b) - 9)
For i = 0 To Channel.ListCount Step 1
If Channel.List(i) = b Then
Channel.ListIndex = i
Exit For
End If
Next i
ElseIf Mid(b, 1, 9) = "CHANNLIST" Then
Channel.Clear
b = Right(b, Len(b) - 9)
TStr = ""
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) <> "/" Then
TStr = TStr & Mid(b, i, 1)
Else
Channel.AddItem TStr
TStr = ""
End If
Next i
ElseIf Mid(b, 1, 9) = "SEARCHRES" Then
b = Right(b, Len(b) - 9)
Dim k As Integer
SearchResults.ListItems.Clear
Dim j As Integer
j = 0
k = 0
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
j = j + 1
If j = 1 Then
Set LItem = SearchResults.ListItems.Add(, , TStr, FileIcon(TStr), FileIcon(TStr))
TStr = ""
k = k + 1
ElseIf j = 2 Then
LItem.SubItems(1) = TStr
TStr = ""
ElseIf j = 3 Then
LItem.SubItems(2) = TStr
TStr = ""
ElseIf j = 4 Then
LItem.SubItems(3) = TStr
TStr = ""
j = 0
End If
Else
TStr = TStr & Mid(b, i, 1)
End If
Next i

If k > 1 Then
Me.Caption = "IntraShare - " & k & " Files Were Found For The SearchString " & Search.Text
ElseIf k = 1 Then
Me.Caption = "IntraShare - " & k & " File Was Found For The SearchString " & Search.Text
Else
Me.Caption = "IntraShare - No Files Were Found For The SearchString " & Search.Text
End If
Search.Enabled = True
ElseIf Mid(b, 1, 9) = "DLREMOTEF" Then
b = Right(b, Len(b) - 9)
For i = 1 To Len(b) Step 1
If Mid(b, i, 1) = "_" Then
j = i
Exit For
End If
Next i
Dim IDex As Integer
Dim RIP As String
IDex = Mid(b, 1, j - 1)
RIP = Mid(b, j + 1)
SocketIn(IDex).HostAddress = RIP
SocketIn(IDex).HostName = RIP
SocketIn(IDex).Connect
ElseIf Mid(b, 1, 9) = "DISPLAYDT" Then
b = Right(b, Len(b) - 9)
If Mid(b, 1, 6) = "AIMAGE" Then
b = Right(b, Len(b) - 6)
'For i = 1 To Len(b) Step 1
'If Mid(b, i, 1) = " " Then
'Exit For
'End If
'Next i
'Dim TimerVar As Integer
'Dim URLVar As String
'Dim PictureVar As String
'Dim PViewer As New Form6
'Load PViewer
'TimerVar = Mid(b, 1, i - 1)
'b = Mid(b, i + 1)
'For i = 1 To Len(b) Step 1
'If Mid(b, i, 1) = " " Then
'Exit For
'End If
'Next i
'URLVar = Mid(b, 1, i - 1)
'PictureVar = Mid(b, i + 1)
'
'If TimerVar < 1 Then
'TimerVar = 1
'End If
'PViewer.Timer2.Interval = TimerVar
'PViewer.URL.Text = URLVar
'If URLVar <> "" Then
'PViewer.Caption = URLVar
'End If
'Dim X As Integer
'Dim Y As Integer
'Dim ToNext As Integer
'For i = 1 To Len(PictureVar) Step 1
'If Mid(PictureVar, i, 1) = "|" Then
'X = X + 1
'ElseIf Mid(PictureVar, i, 1) = "\" Then
'Y = Y + 1
'Else
'For j = i To Len(PictureVar) Step 1
'If Mid(PictureVar, j, 1) = "|" Or Mid(PictureVar, j, 1) = "\" Then
'ToNext = j
'End If
'Next j
'PViewer.ImageR.ForeColor = Mid(PictureVar, i, j - i)
'PViewer.ImageR.FillColor = Mid(PictureVar, i, j - i)
'PViewer.ImageR.PSet (X, Y)
'End If
'Next i
'PViewer.ImageR.Refresh
'PViewer.ImageR.ScaleMode = PViewer.ScaleMode
'PViewer.Width = (PViewer.ImageR.Picture.Width / 1.79) + 120
'PViewer.Command1.Top = PViewer.ImageR.Height + 60
'PViewer.Command1.Width = PViewer.ScaleWidth
'PViewer.Height = (PViewer.ImageR.Height + 720)
'PViewer.Show , Me
'PViewer.ImageR.Refresh
ElseIf Mid(b, 1, 6) = "MSGBOX" Then
b = Right(b, Len(b) - 6)
MsgBox b, vbOKOnly, "Message From Server Administrator"
ElseIf Mid(b, 1, 6) = "STATUS" Then
b = Right(b, Len(b) - 6)
Status.Panels(2).Text = b
End If
Else
RichTextBox1.SelLength = 0
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelRTF = b
RichTextBox1.SelStart = Len(RichTextBox1.Text)
RichTextBox1.SelText = vbNewLine
End If
End Sub

Private Sub socketin_Connect(Index As Integer)
Dim TStr As String
Dim TInt As Integer
Dim FName As String
Dim i As Integer
ReDim FData(Index) As String
ReDim DoneYet(Index) As Boolean
DoneYet(Index) = False
FName = Index
For i = 1 To Downloads.ListItems.Count Step 1
If Downloads.ListItems(i).SubItems(4) = FName Then
Exit For
End If
Next i
TStr = "S" & Downloads.ListItems(i).Text & "_" & Downloads.ListItems(1).SubItems(2)
TInt = Len(TStr)
SocketIn(Index).Write TStr, TInt
End Sub

Private Sub socketin_Disconnect(Index As Integer)
SocketIn(Index).Cleanup
If DoneYet(Index) = False Then
UpdateProgress Index, "Disconnected", vbRed
Else
UpdateProgress Index, "Done", &H8000&
End If
End Sub

Private Sub SocketIn_LastError(Index As Integer, ErrorCode As Integer, ErrorString As String, Response As Integer)
Status.Panels(2).Text = "Download Error: " & ErrorString
End Sub

Private Sub socketin_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)
On Error GoTo Errorh
Dim a As Integer
Dim b As String
Dim i As Integer
Dim j As Integer
Dim FName As String
Dim Lastb As String
Dim TStr As String
Dim TInt As Integer
Dim RLength As Integer
Dim c As String
UpdateProgress Index, "Downloading File", &H800000
b = ""
SocketIn(Index).Read b, DataLength
If b = "(-{~[[<---?¿^v^The^End^v^?¿--->]]~}-)" Then
UpdateProgress Index, "Download Complete - Saving File", vbBlue
For i = 1 To Downloads.ListItems.Count Step 1
If Downloads.ListItems(i).SubItems(4) = Index Then
FName = Downloads.ListItems(i).Text
End If
Next i
Open DLLocation.Text & "/" & FName For Output As Index
Print #Index, FData(Index)
Close Index
On Error GoTo ErrorOk
UpdateProgress Index, "File Saved - Disconnecting", &H8000&
DoneYet(Index) = True
TStr = "D"
TInt = Len(TStr)
SocketIn(Index).Write TStr, TInt
End If
FData(Index) = FData(Index) & b
Exit Sub
Errorh:
UpdateProgress Index, "Error", vbRed
Exit Sub
ErrorOk:
DoneYet(Index) = True
UpdateProgress Index, "Done", &H8000&
Exit Sub
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Me.Caption = "IntraShare"
End Sub

Private Sub Status_PanelClick(ByVal Panel As MSComctlLib.Panel)
If Panel.Index = 1 Then
If Panel.ToolTipText = "Not Connected" Then
Form3.Show vbModal
Else
Status.Panels(2).Text = "Not Connected"
ThePPLBox.ListItems.Clear
Panel.Picture = Form2.MenuImages.ListImages(13).Picture
Panel.ToolTipText = "Not Connected"
Panel.Bevel = sbrRaised
ConnectDisconnect.Caption = "Connect"
Form2.mnuConnectDisconnect.Caption = "Connect"
Socket1.Disconnect
Socket1.Cleanup
'Get the menuhandle of your app
Dim hMenu As Long
Dim hSubMenu As Long
Dim hID As Long
hMenu& = GetMenu(Form2.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 10)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
hMenu& = GetMenu(Form1.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, Form2.MenuImages.ListImages(7).Picture, Form2.MenuImages.ListImages(7).Picture
End If
End If
End Sub

Private Sub Timer1_Timer()
If Socket1.Connected = True Then
Connected True
Else
Connected False
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Interval = Timer2.Interval - 1
End Sub

Private Sub Timer3_Timer()
DLBoxes.Text = 1
End Sub
