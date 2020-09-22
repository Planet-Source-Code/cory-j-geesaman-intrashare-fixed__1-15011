VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private tUsr As Collection

Public Function Add(nUsr As User, Index As String)
tUsr.Add nUsr, Index
Dim i As Integer
i = Form1.PPLCount.Text
Form1.PPLCount.Text = i + 1
End Function

Public Function Remove(Index As String)
On Error Resume Next
tUsr.Remove Index
End Function

Public Function eUser(Index As Integer) As User
On Error Resume Next
Set eUser = tUsr.Item(Index)
End Function

Public Function iCount() As String
Set iCount = Form1.PPLCount.Text
End Function

Private Sub Class_Initialize()
Set tUsr = New Collection
End Sub
