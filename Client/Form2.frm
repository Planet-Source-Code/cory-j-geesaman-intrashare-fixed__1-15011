VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6685E735-3BF6-11D1-A345-444553540000}#1.1#0"; "GRADIENTTITLE.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7305
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9225
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7305
   ScaleWidth      =   9225
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4200
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   1440
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
   Begin MSComctlLib.ImageList MenuImages 
      Left            =   2520
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":066E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0786
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":089E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":0F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":114A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1266
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1382
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":149E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":15BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":16D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":17F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":190E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1B3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuRightClick 
      Caption         =   "mnuRightClick"
      Begin VB.Menu mnuView 
         Caption         =   "View"
         Begin VB.Menu mnuLargeIcons 
            Caption         =   "Large Icons"
         End
         Begin VB.Menu mnuSmallIcons 
            Caption         =   "Small Icons"
         End
         Begin VB.Menu mnuList 
            Caption         =   "List"
         End
         Begin VB.Menu mnuDetails 
            Caption         =   "Details"
         End
      End
      Begin VB.Menu d001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoArrange 
         Caption         =   "Auto Arrange"
         Checked         =   -1  'True
      End
      Begin VB.Menu d002 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "Download"
      End
   End
   Begin VB.Menu mnuRightClick2 
      Caption         =   "mnuRightClick2"
      Begin VB.Menu mnuStopDownload 
         Caption         =   "Stop Download"
      End
      Begin VB.Menu mnuRetryDownload 
         Caption         =   "Retry Download"
      End
      Begin VB.Menu mnu2Dash001 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeLocationToDownloadTo 
         Caption         =   "Change Location To Download To"
      End
      Begin VB.Menu mnu2Dash002 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemoveFromList 
         Caption         =   "Remove From List"
      End
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "mnuSysTray"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuUsingIntraShare 
         Caption         =   "Using IntraShare"
      End
      Begin VB.Menu mnuDash005 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurrentDownloads 
         Caption         =   "Current Downloads"
      End
      Begin VB.Menu mnuSharedFiles 
         Caption         =   "Shared Files"
      End
      Begin VB.Menu mnuSearchFiles 
         Caption         =   "Search Files"
      End
      Begin VB.Menu mnuchat 
         Caption         =   "Chat"
      End
      Begin VB.Menu mnuDash006 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuDash004 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnectDisconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDash007 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()
SetBold Form2, 2, 8
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 12)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(21).Picture, MenuImages.ListImages(21).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 10)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(7).Picture, MenuImages.ListImages(7).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 8)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(14).Picture, MenuImages.ListImages(14).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 6)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(15).Picture, MenuImages.ListImages(15).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 5)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(16).Picture, MenuImages.ListImages(16).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 4)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(17).Picture, MenuImages.ListImages(17).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 3)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(18).Picture, MenuImages.ListImages(18).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 1)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(19).Picture, MenuImages.ListImages(19).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 2)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(20).Picture, MenuImages.ListImages(20).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuAbout_Click()
Form5.Show
End Sub

Private Sub mnuAutoArrange_Click()
If mnuAutoArrange.Checked = True Then
mnuAutoArrange.Checked = False
Form1.SearchResults.Sorted = False
Else
mnuAutoArrange.Checked = True
Form1.SearchResults.Sorted = True
End If
End Sub

Private Sub mnuChangeLocationToDownloadTo_Click()
Form4.Show vbModal
End Sub

Private Sub mnuchat_Click()
Form1.SSTab1.Tab = 0
Form1.WindowState = vbNormal
Form1.Show
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuConnectDisconnect_Click()
If Form1.ConnectDisconnect.Caption = "Connect" Then
Form3.Show vbModal
Else
Form1.Status.Panels(2).Text = "Not Connected"
Form1.ThePPLBox.ListItems.Clear
Form1.Status.Panels(1).Picture = Form2.MenuImages.ListImages(13).Picture
Form1.Status.Panels(1).ToolTipText = "Not Connected"
Form1.Status.Panels(1).Bevel = sbrRaised
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
Form1.ConnectDisconnect.Caption = "Connect"
mnuConnectDisconnect.Caption = "Connect"
Form1.Socket1.Disconnect
Form1.Socket1.Cleanup
End If
End Sub

Private Sub mnuCurrentDownloads_Click()
Form1.SSTab1.Tab = 3
Form1.WindowState = vbNormal
Form1.Show
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuDetails_Click()
Form1.SearchResults.View = lvwReport
End Sub

Private Sub mnuDownload_Click()
On Error Resume Next
Dim a As Integer
If Form1.SearchResults.SelectedItem.Text <> "" Then
a = Form1.NewDownLoad(Form1.SearchResults.SelectedItem.Text, Form1.SearchResults.SelectedItem.SubItems(1), Form1.SearchResults.SelectedItem.SubItems(2), Form1.SearchResults.SelectedItem.SubItems(3))
End If
End Sub

Private Sub mnuExit_Click()
Unload Form1
End Sub

Private Sub mnuLargeIcons_Click()
Form1.SearchResults.View = lvwIcon
End Sub

Private Sub mnuList_Click()
Form1.SearchResults.View = lvwList
End Sub

Private Sub mnuRemoveFromList_Click()
Timer2.Enabled = True
End Sub

Private Sub mnuRestore_Click()
Form1.WindowState = vbNormal
Form1.Show
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuRetryDownload_Click()
Form1.RetryDownLoad
End Sub

Private Sub mnuSearchFiles_Click()
Form1.SSTab1.Tab = 1
Form1.WindowState = vbNormal
Form1.Show
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuSharedFiles_Click()
Form1.SSTab1.Tab = 2
Form1.WindowState = vbNormal
Form1.Show
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuSmallIcons_Click()
Form1.SearchResults.View = lvwSmallIcon
End Sub

Private Sub mnuStopDownload_Click()
Form1.UpdateProgress Form1.Downloads.SelectedItem.SubItems(4), "Download Stopped", &H80&
Form1.SocketIn(Form1.Downloads.SelectedItem.SubItems(4)).Disconnect
End Sub

Private Sub mnuUsingIntraShare_Click()
Form7.Show
End Sub

Private Sub Timer1_Timer()
Form2.mnuRetryDownload.Visible = True
Form2.mnuStopDownload.Visible = True
Form2.mnu2Dash001.Visible = True
Form2.mnuChangeLocationToDownloadTo.Visible = True
Form2.mnu2Dash002.Visible = True
Form2.mnuRemoveFromList.Visible = True
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(1).Picture, MenuImages.ListImages(1).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 1)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(2).Picture, MenuImages.ListImages(2).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 2)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(3).Picture, MenuImages.ListImages(3).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 3)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(4).Picture, MenuImages.ListImages(4).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 2)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(5).Picture, MenuImages.ListImages(6).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 0)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 4)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(7).Picture, MenuImages.ListImages(7).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 1)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 0)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(9).Picture, MenuImages.ListImages(9).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 1)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 1)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(10).Picture, MenuImages.ListImages(10).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 1)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 3)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(11).Picture, MenuImages.ListImages(11).Picture
'Get the menuhandle of your app
hMenu& = GetMenu(Me.hWnd)

'Get the handle of the first submenu (Hello)
hSubMenu& = GetSubMenu(hMenu&, 1)

'Get the menuId of the first entry (Bitmap)
hID& = GetMenuItemID(hSubMenu&, 5)

'Add the bitmap
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, MenuImages.ListImages(12).Picture, MenuImages.ListImages(12).Picture
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
Form1.Downloads.ListItems.Remove Form1.Downloads.SelectedItem.Index
End Sub

Private Sub Timer3_Timer()
Timer3.Enabled = False
   nid.cbSize = Len(nid)
   nid.hWnd = Me.hWnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = Me.Icon
   nid.szTip = "IntraShare v" & App.Major & "." & App.Minor & App.Revision & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
   Form1.Hide
End Sub

Public Sub SetBold(frmBold As Form, iMenuIndex As Long, iItemIndex As Long)
    Dim hMnu As Long, hSubMnu As Long
    hMnu = GetMenu(frmBold.hWnd)
    hSubMnu = GetSubMenu(hMnu, iMenuIndex)
    SetMenuDefaultItem hSubMnu, iItemIndex, 1&
End Sub

Private Sub Form_MouseMove _
   (Button As Integer, _
    Shift As Integer, _
    X As Single, _
    Y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       Case WM_LBUTTONDBLCLK
       Form1.WindowState = vbNormal
       Form1.Show
       Shell_NotifyIcon NIM_DELETE, nid
       Case WM_RBUTTONDOWN
       Case WM_RBUTTONUP
       Me.PopupMenu Me.mnuSysTray
       Case WM_RBUTTONDBLCLK
    End Select
End Sub

