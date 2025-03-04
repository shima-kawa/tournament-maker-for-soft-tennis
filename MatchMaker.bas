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

' baseMatchIDの試合が何回戦目かを返す
' 引数: baseMatchID
' 戻り値: round
' ex) teams:65~128 -> baseMatchID=16 -> round=3
Function culRound(baseMatchID As Integer) As Integer
    Dim teams As Integer
    Dim requiredRounds As Integer
    
    teams = teamsRange
    requiredRounds = culNumberOfNeedRounds(teams)
    
    culRound = requiredRounds - WorksheetFunction.RoundDown(Log(baseMatchID) / Log(2), 0)
End Function

' そのラウンド(N回戦)の総試合数を返す
' 引数: round
' 戻り値: そのroundの試合数
' ex) teams:65~128 -> round=4 -> returns 8
Function culMatchesPerRound(round As Integer)
    Dim teams As Integer
    Dim requiredRounds As Integer
    
    teams = teamsRange
    requiredRounds = culNumberOfNeedRounds(teams)
    
    culMatchesPerRound = 2 ^ (requiredRounds - round)
    
End Function

Sub test()

setUp
'------------------------------------------------

Dim i As Integer
Dim teams As Integer
Dim basePlayerId As Integer
Dim playerID As Integer
Dim seeds As New seedarray
Dim baseTeams As Integer
Dim startPlgNo As Integer


testWorksheet.Cells.Clear

startPlgNo = plgStartNoRange.Value
basePlayerId = 0
playerID = 0
teams = teamsRange
baseTeams = culNumberOfBaseTeams(teams)

seeds.make (baseTeams)

' 初期化
For i = 1 To baseTeams - 1
    testWorksheet.Cells(i, 1) = i ' match id
    testWorksheet.Cells(i, 2) = "UNDECIDED" ' A
    testWorksheet.Cells(i, 3) = "UNDECIDED" ' B
    testWorksheet.Cells(i, 4) = "UNDECIDED" ' winner
Next i


' 1回戦
For baseMatchID = baseTeams / 2 To baseTeams - 1
    basePlayerId = basePlayerId + 1
    If (seeds.seed(basePlayerId) <= teams) Then
        playerID = playerID + 1
        testWorksheet.Cells(baseMatchID, 2) = playerID + startPlgNo - 1
    Else
        testWorksheet.Cells(baseMatchID, 2) = 0
    End If
    
    basePlayerId = basePlayerId + 1
    If (seeds.seed(basePlayerId) <= teams) Then
        playerID = playerID + 1
        testWorksheet.Cells(baseMatchID, 3) = playerID + startPlgNo - 1
    Else
        testWorksheet.Cells(baseMatchID, 3) = 0
    End If
    
    If (testWorksheet.Cells(baseMatchID, 2) = 0) Then
        testWorksheet.Cells(baseMatchID, 4) = testWorksheet.Cells(baseMatchID, 3)
    ElseIf (testWorksheet.Cells(baseMatchID, 3) = 0) Then
        testWorksheet.Cells(baseMatchID, 4) = testWorksheet.Cells(baseMatchID, 2)
    End If
        
Next baseMatchID

updateBaseMatches (baseTeams) ' 不戦勝の処理を含む

makeMaches (teams)


End Sub
Function makeMaches(teams As Integer)


Dim baseTeams As Integer
Dim needRounds As Integer
Dim round As Integer
Dim Match As Integer
Dim i As Integer
Dim row As Integer

baseTeams = culNumberOfBaseTeams(teams)
needRounds = culNumberOfNeedRounds(teams)
row = 1

matchesWS.Cells.Clear
matchesWS.Cells(row, G_idCol) = "試合ID"
matchesWS.Cells(row, G_roundCol) = "回戦"
matchesWS.Cells(row, G_fromCol) = "始番"
matchesWS.Cells(row, G_toCol) = "終番"
matchesWS.Cells(row, G_baseMatchIdCol) = "ベース試合Id"
matchesWS.Cells(row, G_statusCol) = "状態"
matchesWS.Cells(row, G_matchGamesCol) = "マッチ数"
matchesWS.Cells(row, G_leftCol) = "左No"
matchesWS.Cells(row, G_rightCol) = "右No"
matchesWS.Cells(row, G_winnerCol) = "勝者"
matchesWS.Cells(row, G_scoreLeftCol) = "左スコア"
matchesWS.Cells(row, G_scoreRightCol) = "右スコア"
matchesWS.Cells(row, G_addressLeftRowCol) = "トーナメント" & vbLf & "左座標Row"
matchesWS.Cells(row, G_addressLeftColCol) = "トーナメント" & vbLf & "左座標Col"
matchesWS.Cells(row, G_addressRightRowCol) = "トーナメント" & vbLf & "右座標Row"
matchesWS.Cells(row, G_addressRightColCol) = "トーナメント" & vbLf & "右座標Col"
matchesWS.Cells(row, G_nextMatchRowCol) = "次対戦" & vbLf & "行"
matchesWS.Cells(row, G_nextMatchColCol) = "次対戦" & vbLf & "列"
matchesWS.Cells(row, G_LRCol) = "LR"
row = row + 1

For round = 1 To needRounds
    For Match = baseTeams / (2 ^ round) To baseTeams / (2 ^ (round - 1)) - 1
        If (baseMatchesWS.Cells(Match, 4) = "UNDECIDED") Then
            If (baseMatchesWS.Cells(Match, 2) = "UNDECIDED" Or baseMatchesWS.Cells(Match, 3) = "UNDECIDED") Then
                matchesWS.Cells(row, G_idCol) = row - 1
                matchesWS.Cells(row, G_fromCol) = getLeftLimit(Match)
                matchesWS.Cells(row, G_toCol) = getRightLimit(Match)
                matchesWS.Cells(row, G_baseMatchIdCol) = Match
                matchesWS.Cells(row, G_roundCol) = round
                matchesWS.Cells(row, G_statusCol) = MATCH_NOT_ALLOWED
                matchesWS.Cells(row, G_matchGamesCol) = 7 ' temp
                If (baseMatchesWS.Cells(Match, 2) <> "UNDECIDED" And baseMatchesWS.Cells(Match, 3) = "UNDECIDED") Then ' LEFTが不戦勝
                    matchesWS.Cells(row, G_leftCol) = baseMatchesWS.Cells(Match, 2)
                ElseIf (baseMatchesWS.Cells(Match, 2) = "UNDECIDED" And baseMatchesWS.Cells(Match, 3) <> "UNDECIDED") Then
                    matchesWS.Cells(row, G_rightCol) = baseMatchesWS.Cells(Match, 3)
                End If
                row = row + 1
            Else
                matchesWS.Cells(row, G_idCol) = row - 1
                matchesWS.Cells(row, G_fromCol) = getLeftLimit(Match)
                matchesWS.Cells(row, G_toCol) = getRightLimit(Match)
                matchesWS.Cells(row, G_baseMatchIdCol) = Match
                matchesWS.Cells(row, G_roundCol) = round
                matchesWS.Cells(row, G_statusCol) = MATCH_ALLOWED_NOPRINT
                matchesWS.Cells(row, G_leftCol) = getLeftLimit(Match)
                matchesWS.Cells(row, G_rightCol) = getRightLimit(Match)
                matchesWS.Cells(row, G_matchGamesCol) = 7 ' temp
                row = row + 1
            End If
        End If
    Next Match
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

Function getLeftLimit(ByVal Match As Integer) As Integer
    
    While baseMatchesWS.Cells(Match, 2) = "UNDECIDED"
        Match = Match * 2
    Wend
    
    getLeftLimit = baseMatchesWS.Cells(Match, 2)
End Function

Function getRightLimit(ByVal Match As Integer) As Integer
    
    While baseMatchesWS.Cells(Match, 3) = "UNDECIDED"
        Match = Match * 2 + 1
    Wend
    
    getRightLimit = baseMatchesWS.Cells(Match, 3)
End Function

Function getRequiredGames(game As Integer) As Integer
    getRequiredGames = WorksheetFunction.RoundUp(game / 2, 0)
End Function

' 試合シートに、結果を出力するトーナメントシート上の座標を保存する
' 引数：baseMatchId, selectedSide(プレイヤーのサイド), 各座標
Function setAddress(baseMatchID As Integer, selectedSide As Integer, row As Integer, col As Integer)
    Dim lastRow As Integer
    Dim i As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    

    For i = 1 To lastRow
        If (matchesWS.Cells(i, G_baseMatchIdCol) = baseMatchID) Then
            If (selectedSide = LEFT) Then
                With matchesWS
                    .Cells(i, G_addressLeftRowCol) = row
                    .Cells(i, G_addressLeftColCol) = col
                End With
            Else
                With matchesWS
                    .Cells(i, G_addressRightRowCol) = row
                    .Cells(i, G_addressRightColCol) = col
                End With
            End If
            Exit Function
        End If
    Next i
    MsgBox "エラー。試合シートに対象の試合が見つかりませんでした。"
End Function

Function determineNextAddress()
    Dim i As Integer
    Dim j As Integer
    Dim lastRow As Integer
    Dim nextRow As Integer
    Dim baseId As Integer
    Dim nextBaseId As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    For i = 2 To lastRow - 1 '1行目はタイトル、最終行は決勝なので、除外
        baseId = matchesWS.Cells(i, G_baseMatchIdCol)
        nextBaseId = WorksheetFunction.RoundDown(baseId / 2, 0)
        
        ' 次の試合の行を探す
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

Function insertProgramNumber(startNum As Integer, teamNum As Integer)
    Dim i As Integer
    
    For i = 1 To teamNum
        playerListWS.Cells(i + 1, 1) = startNum + i - 1
    Next i
End Function

Sub a()
    setUp
    'Call insertProgramNumber(plgStartNoRange.Value, teamsRange.Value)

    'Call changeNumOfGames(1, 5)
    Debug.Print culMatchesPerRound(culRound(7))
End Sub
' 指定したラウンド(round 回戦)のゲーム数を変更する
Function changeNumOfGames(round As Integer, numGames As Integer)

    Dim i As Integer
    Dim lastRow As Integer
    Dim countOfChanges As Integer
    Dim msgRes As VbMsgBoxResult
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    countOfChanges = 0
    
    ' 指定したラウンドの試合に、終了済みの試合があるかチェック
    For i = 2 To lastRow
        If (matchesWS.Cells(i, G_roundCol).Value = round) Then
            If (matchesWS.Cells(i, G_statusCol).Value = MATCH_FINISHED) Then
                MsgBox "ゲーム数を変更できません。指定したラウンドの試合の一部がすでに" & matchesWS.Cells(i, G_matchGamesCol) & "ゲームで終了しています。強制的に変更する場合は、「試合」シートを直接編集してください。", _
                    Buttons:=vbCritical, _
                    Title:="エラー"
                Exit Function
            End If
            countOfChanges = countOfChanges + 1
        End If
    Next i
    
    If (countOfChanges = 0) Then
        MsgBox "エラー。対象の試合が見つかりませんでした。", _
                Buttons:=vbExclamation, _
                Title:="エラー"
        Exit Function
    End If
    
    ' ゲーム数変更確認
    msgRes = MsgBox(countOfChanges & "件の試合を" & numGames & "ゲームに変更します。", _
        Buttons:=vbOKCancel, _
        Title:="確認" _
    )
    
    If (msgRes <> vbOK) Then
        Exit Function
    End If
    
    ' ゲーム数変更
    For i = 2 To lastRow
        If (matchesWS.Cells(i, G_roundCol).Value = round) Then
            matchesWS.Cells(i, G_matchGamesCol) = numGames
        End If
    Next i
    
    MsgBox "変更が完了しました"

End Function

Function getNextMatchStatus(matchID As Integer) As Integer
    Dim row As Integer
    Dim lastRow As Integer
    Dim nextRow As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    
    For row = 2 To lastRow
        If (matchesWS.Cells(row, G_idCol).Value = matchID) Then
            nextRow = matchesWS.Cells(row, G_nextMatchRowCol).Value
            Exit For
        End If
    Next row
    
    getNextMatchStatus = matchesWS.Cells(nextRow, G_statusCol)
    
End Function

Function getFirstMatchID(playerID As Integer) As Integer
    Dim row As Integer
    Dim lastRow As Integer
    Dim startPlgNum As Integer
    Dim endPlgNum As Integer
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    For row = 2 To lastRow
        startPlgNum = matchesWS.Cells(row, G_fromCol).Value
        endPlgNum = matchesWS.Cells(row, G_toCol).Value
        If (playerID >= startPlgNum And playerID <= endPlgNum) Then
            getFirstMatchID = matchesWS.Cells(row, G_idCol).Value
            Exit Function
        End If
    Next row
End Function
' 試合シートのLRC列を決定する
Function setLRC()
setUp
    Dim row As Integer
    Dim lastRow As Integer
    Dim baseMatcheID As Integer
    Dim round As Integer
    Dim necessaryRounds As Integer
    Dim baseTeams As Integer
    Dim numPage As Integer
    Dim numDivisions As Integer
    Dim LRC As Integer
    
    necessaryRounds = culNumberOfNeedRounds(teamsRange.Value)
    baseTeams = culNumberOfBaseTeams(teamsRange.Value)
    numPage = getPageNumber(baseTeams, maxNumPerPageRange.Value)
    numDivisions = numPage * 2 ' ページごとに左右があるため、numPage*2
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    
    For row = 2 To lastRow
        baseMatcheID = matchesWS.Cells(row, G_baseMatchIdCol).Value
        round = culRound(baseMatcheID)
        LRC = WorksheetFunction.RoundDown(baseMatcheID / (baseTeams / WorksheetFunction.Power(2, round) / numDivisions), 0) Mod 2
        If baseMatcheID < numPage Then
            matchesWS.Cells(row, G_LRCol) = "-"
        ElseIf baseMatcheID < numPage * 2 Then
            matchesWS.Cells(row, G_LRCol) = CENTER
        Else
            If LRC = 0 Then
                matchesWS.Cells(row, G_LRCol) = LEFT
            Else
                matchesWS.Cells(row, G_LRCol) = RIGHT
            End If
        End If
    Next row
End Function
