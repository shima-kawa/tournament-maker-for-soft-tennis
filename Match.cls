VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_matchID As Integer
Private p_playerLeftNumber As Integer
Private p_playerRightnumber As Integer
Private p_matcheGames As Integer

Public Property Get matchID() As Integer
    matchID = p_matchID
End Property

Public Property Get leftNum() As Integer
    leftNum = p_playerLeftNumber
End Property

Public Property Get rightNum() As Integer
    rightNum = p_playerRightnumber
End Property
Public Property Get matchGames() As Integer
    matchGames = p_matcheGames
End Property

Public Property Let matchID(id As Integer)
    p_matchID = id
End Property

Public Property Let leftNum(LEFT As Integer)
    p_playerLeftNumber = LEFT
End Property

Public Property Let rightNum(RIGHT As Integer)
    p_playerRightnumber = RIGHT
End Property

Public Property Let matchGames(n As Integer)
    p_matcheGames = n
End Property
