VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "seedArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private seedarray() As Integer
Private teams As Integer
Private baseTeams As Integer
Private needRound As Integer
Private playerNumber As Integer

Public Sub make(t As Integer)
    teams = t
    baseTeams = culNumberOfBaseTeams(teams)
    needRound = culNumberOfNeedRounds(teams)
    playerNumber = 0
    
    ReDim seedarray(1 To baseTeams)
    
    Call makeSeedArray(1, 1, True)
    
    Debug.Print ("OBJ: base = " & baseTeams)
End Sub

Public Property Get seed(number As Integer)
    seed = seedarray(number)
End Property

Private Function culNumberOfNeedRounds(teams As Integer) As Integer
    culNumberOfNeedRounds = WorksheetFunction.RoundUp(Log(teams) / Log(2), 0)
End Function
Private Function culNumberOfBaseTeams(teams As Integer) As Integer
    culNumberOfBaseTeams = 2 ^ culNumberOfNeedRounds(teams)
End Function

Function makeSeedArray(parent As Integer, round As Integer, state As Boolean)
    Dim baseNumThisRound As Integer
    Dim youngNum As Integer
    Dim oldNum As Integer
    
    baseNumThisRound = 1
    
    If (round > needRound) Then
        playerNumber = playerNumber + 1
        'Debug.Print "OBJ: No" & playerNumber & " ‘æ " & parent & " ƒV[ƒh"
        seedarray(playerNumber) = parent
        Exit Function
    End If
    
    baseNumThisRound = 2 ^ round
    
    If (state = True) Then
        youngNum = parent
        oldNum = baseNumThisRound + 1 - parent
    Else
        youngNum = baseNumThisRound + 1 - parent
        oldNum = parent
    End If
    
    Call makeSeedArray(youngNum, round + 1, True)
    Call makeSeedArray(oldNum, round + 1, False)
    
End Function

