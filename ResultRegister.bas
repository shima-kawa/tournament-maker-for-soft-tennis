Attribute VB_Name = "ResultRegister"
Function registerResult(r As Result)

    setUp
    
    ' ---------------------------------------------------
    Dim lastRow As Integer
    Dim row As Integer
    Dim addressRow As Integer
    Dim addressCol As Integer
    Dim m As Match
    Dim whichWinner As Integer
    
    Debug.Print ("Regist: matchID: " & r.matchID & ", score: " & r.leftScore & "-" & r.rightScore & ", winner: " & r.winner)
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    'r.id��T��
    For row = 1 To lastRow
        If (r.matchID = matchesWS.Cells(row, G_idCol)) Then
            Exit For
        End If
    Next row
    
    '�����V�[�g�ɃX�R�A��o�^
    matchesWS.Cells(row, G_scoreLeftCol) = r.leftScore
    matchesWS.Cells(row, G_scoreRightCol) = r.rightScore
    
    '�����V�[�g�ɏ��҂�o�^
    matchesWS.Cells(row, G_winnerCol) = r.winner
    If (matchesWS.Cells(row, G_leftCol) = r.winner) Then
        whichWinner = LEFT
    Else
        whichWinner = RIGHT
    End If
    
    '�g�[�i�����g�V�[�g�ɔ��f
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
    If culNumberOfNeedRounds(teamsRange.Value) <> matchesWS.Cells(row, G_roundCol) Then ' �����ȊO
        addressRow = matchesWS.Cells(row, G_nextMatchRowCol)
        addressCol = matchesWS.Cells(row, G_nextMatchColCol)
        matchesWS.Cells(addressRow, addressCol) = r.winner
    End If
    
    '�X�e�[�^�X�̍X�V
    Debug.Print ("Before Status of current: " & matchesWS.Cells(row, G_statusCol).Address(False, False) & " " & matchesWS.Cells(row, G_statusCol))
    Debug.Print ("Before Status of next   : " & matchesWS.Cells(addressRow, G_statusCol).Address(False, False) & " " & matchesWS.Cells(addressRow, G_statusCol))
    matchesWS.Cells(row, G_statusCol) = MATCH_FINISHED
    If ( _
        matchesWS.Cells(addressRow, G_leftCol) <> "" And _
        matchesWS.Cells(addressRow, G_rightCol) <> "" And _
        matchesWS.Cells(addressRow, G_statusCol) <> MATCH_FINISHED _
    ) Then
        matchesWS.Cells(addressRow, G_statusCol) = MATCH_ALLOWED_NOPRINT
    End If
    Debug.Print ("After Status of current : " & matchesWS.Cells(row, G_statusCol).Address(False, False) & " " & matchesWS.Cells(row, G_statusCol))
    Debug.Print ("After Status of next    : " & matchesWS.Cells(addressRow, G_statusCol).Address(False, False) & " " & matchesWS.Cells(addressRow, G_statusCol))
    
End Function

' �����̌���
' �����̃v���O����No���L�[�ɁA������T���B����������A�����I�u�W�F�N�g��Ԃ��B
' �����F�Ⴂ���̃v���O�����ԍ�
' �߂�l�F�������ʂ̎����I�u�W�F�N�g
Function findMatch(key As Integer) As Match
    
    ' ---------------------------------------------------
    Dim lastRow As Integer
    Dim row As Integer
    Dim matchObj As Match
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    For row = 2 To lastRow
        If (matchesWS.Cells(row, G_statusCol) = MATCH_ALLOWED_PRINTED And matchesWS.Cells(row, G_leftCol) = key) Then
            Set findMatch = New Match
            findMatch.matchID = matchesWS.Cells(row, G_idCol)
            findMatch.leftNum = matchesWS.Cells(row, G_leftCol)
            findMatch.rightNum = matchesWS.Cells(row, G_rightCol)
            findMatch.matchGames = matchesWS.Cells(row, G_matchGamesCol)
            Exit Function
        End If
    Next row
    
    Set findMatch = Nothing
End Function

' �w�肵���X�e�[�^�X�̎����̌���
Function findAllMatchesWithStatus(key As Integer, status As Integer) As Match()

    Dim lastRow As Integer
    Dim row As Integer
    Dim matchObj As Match
    Dim matches() As Match
    Dim index As Integer

    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    ReDim matches(0)
    
    index = UBound(matches)
    
    For row = 2 To lastRow
        If (matchesWS.Cells(row, G_statusCol) = status And (matchesWS.Cells(row, G_leftCol) = key Or matchesWS.Cells(row, G_rightCol) = key)) Then
            index = index + 1
            ReDim Preserve matches(index)
            Set matches(index) = New Match
            matches(index).matchID = matchesWS.Cells(row, G_idCol)
            matches(index).leftNum = matchesWS.Cells(row, G_leftCol)
            matches(index).rightNum = matchesWS.Cells(row, G_rightCol)
            matches(index).matchGames = matchesWS.Cells(row, G_matchGamesCol)
        End If
    Next row
    
    findAllMatchesWithStatus = matches

End Function

Function findResult(matchID As Integer) As Result
    Dim lastRow As Integer
    Dim row As Integer
    Dim resultObj As Result
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    For row = 2 To lastRow
        If (matchesWS.Cells(row, G_idCol) = matchID) Then
            If (matchesWS.Cells(row, G_statusCol) = MATCH_FINISHED) Then
                Set resultObj = New Result
                resultObj.matchID = matchID
                resultObj.leftScore = matchesWS.Cells(row, G_scoreLeftCol)
                resultObj.rightScore = matchesWS.Cells(row, G_scoreRightCol)
                resultObj.winner = matchesWS.Cells(row, G_winnerCol)
            Else
                Set findResult = Nothing
            End If
            Set findResult = resultObj
            Exit Function
        End If
    Next row
End Function
' ���͂��ꂽ�I��̎����ꗗ���擾����
' �ꗗ���C����ʏ�ɕ\��
' �ύX���ꂽ�X�R�A��T��
' ���s���ς�邩�ǂ����̃`�F�b�N
' ���s���ς��Ȃ��ꍇ�A�g�[�i�����g�V�[�g�ɃX�R�A���L�ځA�����������A�Ԑ��������A�őΉ��I��
' ���s���ς��ꍇ�A���̐�̎������I���ς݂��m�F
' ��̎������n�܂��Ă��Ȃ��ꍇ�́A��L���l�����A�̓_�[�̏�Ԃ��m�F�A����ς݂�������A�Ĉ���������邩�����B�Ĉ������ꍇ�́A�̓_�\�{�^���������悤�Ɏw��
' ��̎������I���ς݂̏ꍇ�́A�I���ς݂̎������ꗗ�Ŏ擾�B���̏ꍇ�A�ΏۑI�肪�ւ���������݂̂�����̂��A�g�[�i�����g�̍Ō�܂Ō���̂�...
' ���̂悤�ɕύX����܂����A��낵���ł����B����Ƃ��A���̎����̌��ʂ�j��
' N ���� �X�R�A �� �X�R�A ���� N�@���@N ���� �X�R�A �� �X�R�A ���� M
Sub aiu()
    setUp
    MsgBox findResult(1).winner
End Sub
