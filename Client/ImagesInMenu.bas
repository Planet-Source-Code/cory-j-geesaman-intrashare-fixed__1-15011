Attribute VB_Name = "Images"
Option Explicit

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long

Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1

Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long

Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Public Const MF_BITMAP = &H4&

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean

Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&

Public Function FileSizer(FileName As String) As String
Dim FileSize1 As String
Dim FileSize2 As String
    ' convert the current file size to different format
    Static xx
    FileSize1 = FileLen(FileName)
    xx = FileSize1 / 1024
    If Len(FileSize1) >= 7 Then
        FileSize2 = Format((xx / 1024), "0.00")
        FileSize2 = FileSize2 & " MB"
    ElseIf Len(FileSize1) >= 4 Then
        xx = Format((FileSize1 / 1024), "0.00")
        FileSize2 = xx & " K"
    Else
        FileSize2 = FileSize1 & " Bytes"
    End If
FileSizer = FileSize2
End Function

Public Function ByteSizer(FileSize1 As String) As String
Dim FileSize2 As String
    ' convert the current file size to different format
    Static xx
    xx = FileSize1 / 1024
    If Len(FileSize1) >= 7 Then
        FileSize2 = Format((xx / 1024), "0.00")
        FileSize2 = FileSize2 & " MB"
    ElseIf Len(FileSize1) >= 4 Then
        xx = Format((FileSize1 / 1024), "0.00")
        FileSize2 = xx & " K"
    Else
        FileSize2 = FileSize1 & " Bytes"
    End If
ByteSizer = FileSize2
End Function

Function FileModified(FileName As String) As String
    Dim nLength As Long
    Dim sSpaces As Long
    Dim NewEnt As String
    Dim DateFix As String
    Dim NewDate As String
    DateFix = Str(FileDateTime(FileName))
    If Mid(DateFix, 2, 1) = "/" Then
        'm1
        NewDate = "0" + Mid(DateFix, 1, 2)
        If Mid(DateFix, 4, 1) = "/" Then
            'm1 d1
            NewDate = NewDate + "0" + Mid(DateFix, 3, 4) + Space(2)
            If Len(DateFix) < 9 Then
                'm1 d1 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 9, 1) = ":" Then
                'm1 d1 h1
                NewDate = NewDate + "0" + Mid(DateFix, 8, 17)
            Else
                'm1 d1 h2
                NewDate = NewDate + Mid(DateFix, 8, 18)
            End If
        Else
            'm1 d2
            NewDate = NewDate + Mid(DateFix, 3, 5) + Space(2)
            If Len(DateFix) < 9 Then
                'm1 d2 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 10, 1) = ":" Then
                'm1 d2 h1
                NewDate = NewDate + "0" + Mid(DateFix, 9, 17)
            Else
                'm1 d2 h2
                NewDate = NewDate + Mid(DateFix, 9, 18)
            End If
        End If
    Else
        'm2
        NewDate = Mid(DateFix, 1, 3)
        If Mid(DateFix, 5, 1) = "/" Then
            'm2 d1
            NewDate = NewDate + "0" + Mid(DateFix, 4, 4) + Space(2)
            If Len(DateFix) < 9 Then
                'm2 d1 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 10, 1) = ":" Then
                'm2 d1 h1
                NewDate = NewDate + "0" + Mid(DateFix, 9, 17)
            Else
                'm2 d1 h2
                NewDate = NewDate + Mid(DateFix, 9, 18)
            End If
        Else
            'm2 d2
            NewDate = NewDate + Mid(DateFix, 4, 5) + Space(2)
            If Len(DateFix) < 9 Then
                'm2 d2 Midnight
                NewDate = NewDate + "12:00:00 AM"
            End If
            If Mid(DateFix, 11, 1) = ":" Then
                'm2 d2 h1
                NewDate = NewDate + "0" + Mid(DateFix, 10, 17)
            Else
                'm2 d2 h2
                NewDate = NewDate + Mid(DateFix, 10, 18)
            End If
        End If
    End If
FileModified = NewDate
End Function
