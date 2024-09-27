Attribute VB_Name = "ResultRegister"
Function registerResult(r As result)

    setUp
    
    ' ---------------------------------------------------
    Dim lastRow As Integer
    Dim row As Integer
    Dim addressRow As Integer
    Dim addressCol As Integer
    Dim m As match
    Dim whichWinner As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.count, 1).End(xlUp).row
    'r.id��T��
    For row = 1 To lastRow
        If (r.matchID = matchesWS.Cells(row, G_idCol)) Then
            Exit For
        End If
    Next row
    
    '�X�R�A��o�^
    matchesWS.Cells(row, G_scoreLeftCol) = r.leftScore
    matchesWS.Cells(row, G_scoreRightCol) = r.rightScore
    
    '���ғo�^
    matchesWS.Cells(row, G_winnerCol) = r.winner
    If (matchesWS.Cells(row, G_leftCol) = r.winner) Then
        whichWinner = LEFT
    Else
        whichWinner = RIGHT
    End If
    
    '�g�[�i�����g�ɔ��f
    addressRow = matchesWS.Cells(row, G_addressLeftRowCol)
    addressCol = matchesWS.Cells(row, G_addressLeftColCol)
    If (whichWinner = LEFT) Then
        tournamentWS.Cells(addressRow, addressCol) = Chr(Asc("�@") + r.leftScore - 1)
    Else
        tournamentWS.Cells(addressRow, addressCol) = r.leftScore
    End If
    
    addressRow = matchesWS.Cells(row, G_addressRightRowCol)
    addressCol = matchesWS.Cells(row, G_addressRightColCol)
    'tournamentWS.Cells(addressRow, addressCol) = r.rightScore
    If (whichWinner = RIGHT) Then
        tournamentWS.Cells(addressRow, addressCol) = Chr(Asc("�@") + r.rightScore - 1)
    Else
        tournamentWS.Cells(addressRow, addressCol) = r.rightScore
    End If
    
    '�g���̍X�V
    If (matchesWS.Cells(row, G_LRCol) = LEFT) Then
        Call drawResultLine(matchesWS.Cells(row, G_baseMatchIdCol), matchesWS.Cells(row, G_addressLeftRowCol) + 1, matchesWS.Cells(row, G_addressRightRowCol) - 1, matchesWS.Cells(row, G_addressLeftColCol) - 1, whichWinner, LEFT)
    Else
        Call drawResultLine(matchesWS.Cells(row, G_baseMatchIdCol), matchesWS.Cells(row, G_addressLeftRowCol) + 1, matchesWS.Cells(row, G_addressRightRowCol) - 1, matchesWS.Cells(row, G_addressLeftColCol) + 1, whichWinner, RIGHT)
    End If
    
    '���ΐ�̓o�^
    addressRow = matchesWS.Cells(row, G_nextMatchRowCol)
    addressCol = matchesWS.Cells(row, G_nextMatchColCol)
    matchesWS.Cells(addressRow, addressCol) = r.winner
    
    '�X�e�[�^�X�̍X�V
    matchesWS.Cells(row, G_statusCol) = MATCH_FINISHED
    If (matchesWS.Cells(addressRow, G_leftCol) <> "" And matchesWS.Cells(addressRow, G_rightCol) <> "") Then
        matchesWS.Cells(addressRow, G_statusCol) = MATCH_ALLOWED_NOPRINT
    End If
    
End Function

' �����̌���
' �����̃v���O����No���L�[�ɁA������T���B����������A�����I�u�W�F�N�g��Ԃ��B
' �����F�Ⴂ���̃v���O�����ԍ�
' �߂�l�F�������ʂ̎����I�u�W�F�N�g
Function findMatch(key As Integer) As match
    
    ' ---------------------------------------------------
    Dim lastRow As Integer
    Dim row As Integer
    Dim matchObj As match
    
    lastRow = matchesWS.Cells(matchesWS.Rows.count, 1).End(xlUp).row
    
    For row = 2 To lastRow
        If (matchesWS.Cells(row, G_statusCol) = MATCH_ALLOWED_PRINTED And matchesWS.Cells(row, G_leftCol) = key) Then
            Set findMatch = New match
            findMatch.matchID = matchesWS.Cells(row, G_idCol)
            findMatch.leftNum = matchesWS.Cells(row, G_leftCol)
            findMatch.rightNum = matchesWS.Cells(row, G_rightCol)
            findMatch.matchGames = matchesWS.Cells(row, G_matchGamesCol)
            Exit Function
        End If
    Next row
    
    Set findMatch = Nothing
End Function

