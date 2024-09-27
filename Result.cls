VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_matchID As Integer
Private p_leftScore As Integer
Private p_rightScore As Integer
Private p_winner As Integer

Public Property Get matchID() As Integer
    matchID = p_matchID
End Property

Public Property Get leftScore() As Integer
    leftScore = p_leftScore
End Property

Public Property Get rightScore() As Integer
    rightScore = p_rightScore
End Property

Public Property Get winner() As Integer
    winner = p_winner
End Property


Public Property Let matchID(id As Integer)
    p_matchID = id
End Property

Public Property Let leftScore(s As Integer)
    p_leftScore = s
End Property

Public Property Let rightScore(s As Integer)
    p_rightScore = s
End Property

Public Property Let winner(w As Integer)
    p_winner = w
End Property
