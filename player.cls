VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private programNumber As Integer
Private playerAName As String
Private playerATeam As String
Private playerBName As String
Private playerBTeam As String

Public Property Get programNum() As Integer
    programNum = programNumber
End Property
Public Property Get AName() As String
    AName = playerAName
End Property
Public Property Get BName() As String
    BName = playerBName
End Property
Public Property Get ATeam() As String
    ATeam = playerATeam
End Property
Public Property Get BTeam() As String
    BTeam = playerBTeam
End Property

Public Property Let programNum(n As Integer)
    programNumber = n
End Property
Public Property Let AName(a As String)
    playerAName = a
End Property
Public Property Let BName(b As String)
    playerBName = b
End Property
Public Property Let ATeam(a As String)
    playerATeam = a
End Property
Public Property Let BTeam(b As String)
    playerBTeam = b
End Property


