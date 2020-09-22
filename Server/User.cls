VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim srNm As String
Dim CrrntChnnl As Integer
Dim IPaddr As String
Dim Flg As String

Property Let UserName(UN As String)
srNm = UN
End Property
Property Get UserName() As String
UserName = srNm
End Property

Property Let CurrentChannel(Index As Integer)
CrrntChnnl = Index
End Property
Property Get CurrentChannel() As Integer
CurrentChannel = CrrntChnnl
End Property

Property Let IPAddress(IPAddress As String)
IPaddr = IPAddress
End Property
Property Get IPAddress() As String
IPAddress = IPaddr
End Property

Property Let Flag(Key As String)
Flg = Key
End Property
Property Get Flag() As String
Flag = Flg
End Property

Private Sub Class_Initialize()
CrrntChnnl = 0
End Sub
