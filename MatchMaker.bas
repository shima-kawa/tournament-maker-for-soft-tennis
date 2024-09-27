Attribute VB_Name = "MatchMaker"
Sub makeMatches()

    setUp
    
    '------------------------------------------------
    Dim teams As Integer
    Dim baseTeams As Integer
    Dim seeds As New seedarray
    
    teams = teamsRange
    baseTeams = culNumberOfBaseTeams(teams)
    seeds.make (teams)
    
    Debug.Print "teams= " & teams
    Debug.Print "base teams= " & baseTeams
End Sub
Function culNumberOfNeedRounds(teams As Integer) As Integer
    culNumberOfNeedRounds = WorksheetFunction.RoundUp(Log(teams) / Log(2), 0)
End Function
Function culNumberOfBaseTeams(teams As Integer) As Integer
    culNumberOfBaseTeams = 2 ^ culNumberOfNeedRounds(teams)
End Function


Sub test()

setUp
'------------------------------------------------

Dim i As Integer
Dim teams As Integer
Dim basePlayerId As Integer
Dim playerId As Integer
Dim seeds As New seedarray

testWorksheet.Cells.Clear

basePlayerId = 0
playerId = 0
teams = teamsRange
baseTeams = culNumberOfBaseTeams(teams)

seeds.make (baseTeams)

' ������
For i = 1 To baseTeams - 1
    testWorksheet.Cells(i, 1) = i ' match id
    testWorksheet.Cells(i, 2) = "UNDECIDED" ' A
    testWorksheet.Cells(i, 3) = "UNDECIDED" ' B
    testWorksheet.Cells(i, 4) = "UNDECIDED" ' winner
Next i


' 1���
For baseMatchID = baseTeams / 2 To baseTeams - 1
    basePlayerId = basePlayerId + 1
    If (seeds.seed(basePlayerId) <= teams) Then
        playerId = playerId + 1
        testWorksheet.Cells(baseMatchID, 2) = playerId
    Else
        testWorksheet.Cells(baseMatchID, 2) = 0
    End If
    
    basePlayerId = basePlayerId + 1
    If (seeds.seed(basePlayerId) <= teams) Then
        playerId = playerId + 1
        testWorksheet.Cells(baseMatchID, 3) = playerId
    Else
        testWorksheet.Cells(baseMatchID, 3) = 0
    End If
    
    If (testWorksheet.Cells(baseMatchID, 2) = 0) Then
        testWorksheet.Cells(baseMatchID, 4) = testWorksheet.Cells(baseMatchID, 3)
    ElseIf (testWorksheet.Cells(baseMatchID, 3) = 0) Then
        testWorksheet.Cells(baseMatchID, 4) = testWorksheet.Cells(baseMatchID, 2)
    End If
        
Next baseMatchID

updateBaseMatches (baseTeams) ' �s�폟�̏������܂�

makeMaches (teams)


End Sub
Function makeMaches(teams As Integer)


Dim baseTeams As Integer
Dim needRounds As Integer
Dim round As Integer
Dim match As Integer
Dim i As Integer
Dim row As Integer

baseTeams = culNumberOfBaseTeams(teams)
needRounds = culNumberOfNeedRounds(teams)
row = 1

matchesWS.Cells.Clear
matchesWS.Cells(row, G_idCol) = "����ID"
matchesWS.Cells(row, G_roundCol) = "���"
matchesWS.Cells(row, G_fromCol) = "�n��"
matchesWS.Cells(row, G_toCol) = "�I��"
matchesWS.Cells(row, G_baseMatchIdCol) = "�x�[�X����Id"
matchesWS.Cells(row, G_statusCol) = "���"
matchesWS.Cells(row, G_matchGamesCol) = "�}�b�`��"
matchesWS.Cells(row, G_leftCol) = "��No"
matchesWS.Cells(row, G_rightCol) = "�ENo"
matchesWS.Cells(row, G_winnerCol) = "����"
matchesWS.Cells(row, G_scoreLeftCol) = "���X�R�A"
matchesWS.Cells(row, G_scoreRightCol) = "�E�X�R�A"
matchesWS.Cells(row, G_addressLeftRowCol) = "�g�[�i�����g" & vbLf & "�����WRow"
matchesWS.Cells(row, G_addressLeftColCol) = "�g�[�i�����g" & vbLf & "�����WCol"
matchesWS.Cells(row, G_addressRightRowCol) = "�g�[�i�����g" & vbLf & "�E���WRow"
matchesWS.Cells(row, G_addressRightColCol) = "�g�[�i�����g" & vbLf & "�E���WCol"
matchesWS.Cells(row, G_nextMatchRowCol) = "���ΐ�" & vbLf & "�s"
matchesWS.Cells(row, G_nextMatchColCol) = "���ΐ�" & vbLf & "��"
matchesWS.Cells(row, G_LRCol) = "LR"
row = row + 1

For round = 1 To needRounds
    For match = baseTeams / (2 ^ round) To baseTeams / (2 ^ (round - 1)) - 1
        If (baseMatchesWS.Cells(match, 4) = "UNDECIDED") Then
            If (baseMatchesWS.Cells(match, 2) = "UNDECIDED" Or baseMatchesWS.Cells(match, 3) = "UNDECIDED") Then
                matchesWS.Cells(row, G_idCol) = row - 1
                matchesWS.Cells(row, G_fromCol) = getLeftLimit(match)
                matchesWS.Cells(row, G_toCol) = getRightLimit(match)
                matchesWS.Cells(row, G_baseMatchIdCol) = match
                matchesWS.Cells(row, G_roundCol) = round
                matchesWS.Cells(row, G_statusCol) = MATCH_NOT_ALLOWED
                matchesWS.Cells(row, G_matchGamesCol) = 7 ' temp
                If (baseMatchesWS.Cells(match, 2) <> "UNDECIDED" And baseMatchesWS.Cells(match, 3) = "UNDECIDED") Then ' LEFT���s�폟
                    matchesWS.Cells(row, G_leftCol) = baseMatchesWS.Cells(match, 2)
                ElseIf (baseMatchesWS.Cells(match, 2) = "UNDECIDED" And baseMatchesWS.Cells(match, 3) <> "UNDECIDED") Then
                    matchesWS.Cells(row, G_rightCol) = baseMatchesWS.Cells(match, 3)
                End If
                row = row + 1
            Else
                matchesWS.Cells(row, G_idCol) = row - 1
                matchesWS.Cells(row, G_fromCol) = getLeftLimit(match)
                matchesWS.Cells(row, G_toCol) = getRightLimit(match)
                matchesWS.Cells(row, G_baseMatchIdCol) = match
                matchesWS.Cells(row, G_roundCol) = round
                matchesWS.Cells(row, G_statusCol) = MATCH_ALLOWED_NOPRINT
                matchesWS.Cells(row, G_leftCol) = getLeftLimit(match)
                matchesWS.Cells(row, G_rightCol) = getRightLimit(match)
                matchesWS.Cells(row, G_matchGamesCol) = 7 ' temp
                row = row + 1
            End If
        End If
    Next match
Next round

Call determineNextAddress

End Function
Function updateBaseMatches(baseMatches As Integer)
    Dim i As Integer
    
    For i = 1 To baseMatches / 2 - 1
        If (baseMatchesWS.Cells(i * 2, 4) <> "UNDECIDED") Then
            baseMatchesWS.Cells(i, 2) = baseMatchesWS.Cells(i * 2, 4)
        End If
        If (baseMatchesWS.Cells(i * 2 + 1, 4) <> "UNDECIDED") Then
            baseMatchesWS.Cells(i, 3) = baseMatchesWS.Cells(i * 2 + 1, 4)
        End If
    Next i
End Function

Function getLeftLimit(ByVal match As Integer) As Integer
    
    While baseMatchesWS.Cells(match, 2) = "UNDECIDED"
        match = match * 2
    Wend
    
    getLeftLimit = baseMatchesWS.Cells(match, 2)
End Function

Function getRightLimit(ByVal match As Integer) As Integer
    
    While baseMatchesWS.Cells(match, 3) = "UNDECIDED"
        match = match * 2 + 1
    Wend
    
    getRightLimit = baseMatchesWS.Cells(match, 3)
End Function

Function getRequiredGames(game As Integer) As Integer
    getRequiredGames = WorksheetFunction.RoundUp(game / 2, 0)
End Function

' �����V�[�g�ɁA���ʂ��o�͂���g�[�i�����g�V�[�g��̍��W��ۑ�����
' �����FbaseMatchId, �e���W
Function setAddress(baseMatchID As Integer, selectedSide As Integer, row As Integer, col As Integer)
    Dim lastRow As Integer
    Dim i As Integer
    Dim side As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.count, 1).End(xlUp).row
    
    If (col < (G_endTournamentArea - G_startTournamentArea + 1) / 2 + G_startTournamentArea) Then
        side = LEFT
    Else
        side = RIGHT
    End If

    For i = 1 To lastRow
        If (matchesWS.Cells(i, G_baseMatchIdCol) = baseMatchID) Then
            If (selectedSide = LEFT) Then
                With matchesWS
                    .Cells(i, G_addressLeftRowCol) = row
                    .Cells(i, G_addressLeftColCol) = col
                    .Cells(i, G_LRCol) = side
                End With
            Else
                With matchesWS
                    .Cells(i, G_addressRightRowCol) = row
                    .Cells(i, G_addressRightColCol) = col
                    .Cells(i, G_LRCol) = side
                End With
            End If
            Exit Function
        End If
    Next i
    MsgBox "�G���[�B�����V�[�g�ɑΏۂ̎�����������܂���ł����B"
End Function

Function determineNextAddress()
    Dim i As Integer
    Dim j As Integer
    Dim lastRow As Integer
    Dim nextRow As Integer
    Dim baseId As Integer
    Dim nextBaseId As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow - 1 '1�s�ڂ̓^�C�g���A�ŏI�s�͌����Ȃ̂ŁA���O
        baseId = matchesWS.Cells(i, G_baseMatchIdCol)
        nextBaseId = WorksheetFunction.RoundDown(baseId / 2, 0)
        
        ' ���̎����̍s��T��
        For j = i To lastRow
            If (matchesWS.Cells(j, G_baseMatchIdCol) = nextBaseId) Then
                nextRow = j
                Exit For
            End If
        Next j
        
        If (baseId Mod 2 = 0) Then
            matchesWS.Cells(i, G_nextMatchRowCol) = nextRow
            matchesWS.Cells(i, G_nextMatchColCol) = G_leftCol
        Else
            matchesWS.Cells(i, G_nextMatchRowCol) = nextRow
            matchesWS.Cells(i, G_nextMatchColCol) = G_rightCol
        End If
    Next i
    
    
End Function
