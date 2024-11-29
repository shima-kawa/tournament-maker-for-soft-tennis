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
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    'r.idを探す
    For row = 1 To lastRow
        If (r.matchID = matchesWS.Cells(row, G_idCol)) Then
            Exit For
        End If
    Next row
    
    'スコアを登録
    matchesWS.Cells(row, G_scoreLeftCol) = r.leftScore
    matchesWS.Cells(row, G_scoreRightCol) = r.rightScore
    
    '勝者登録
    matchesWS.Cells(row, G_winnerCol) = r.winner
    If (matchesWS.Cells(row, G_leftCol) = r.winner) Then
        whichWinner = LEFT
    Else
        whichWinner = RIGHT
    End If
    
    'トーナメントに反映
    addressRow = matchesWS.Cells(row, G_addressLeftRowCol)
    addressCol = matchesWS.Cells(row, G_addressLeftColCol)
    If (whichWinner = LEFT) Then
        tournamentWS.Cells(addressRow, addressCol) = Chr(Asc("①") + r.leftScore - 1)
    Else
        tournamentWS.Cells(addressRow, addressCol) = r.leftScore
    End If
    
    addressRow = matchesWS.Cells(row, G_addressRightRowCol)
    addressCol = matchesWS.Cells(row, G_addressRightColCol)
    'tournamentWS.Cells(addressRow, addressCol) = r.rightScore
    If (whichWinner = RIGHT) Then
        tournamentWS.Cells(addressRow, addressCol) = Chr(Asc("①") + r.rightScore - 1)
    Else
        tournamentWS.Cells(addressRow, addressCol) = r.rightScore
    End If
    
    '枠線の更新
    If (matchesWS.Cells(row, G_LRCol) = LEFT) Then
        Call drawResultLine(matchesWS.Cells(row, G_baseMatchIdCol), matchesWS.Cells(row, G_addressLeftRowCol) + 1, matchesWS.Cells(row, G_addressRightRowCol) - 1, matchesWS.Cells(row, G_addressLeftColCol) - 1, whichWinner, LEFT)
    Else
        Call drawResultLine(matchesWS.Cells(row, G_baseMatchIdCol), matchesWS.Cells(row, G_addressLeftRowCol) + 1, matchesWS.Cells(row, G_addressRightRowCol) - 1, matchesWS.Cells(row, G_addressLeftColCol) + 1, whichWinner, RIGHT)
    End If
    
    '次対戦の登録
    addressRow = matchesWS.Cells(row, G_nextMatchRowCol)
    addressCol = matchesWS.Cells(row, G_nextMatchColCol)
    matchesWS.Cells(addressRow, addressCol) = r.winner
    
    'ステータスの更新
    matchesWS.Cells(row, G_statusCol) = MATCH_FINISHED
    If (matchesWS.Cells(addressRow, G_leftCol) <> "" And matchesWS.Cells(addressRow, G_rightCol) <> "") Then
        matchesWS.Cells(addressRow, G_statusCol) = MATCH_ALLOWED_NOPRINT
    End If
    
End Function

' 試合の検索
' 引数のプログラムNoをキーに、試合を探す。見つかったら、試合オブジェクトを返す。
' 引数：若い方のプログラム番号
' 戻り値：検索結果の試合オブジェクト
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

' 指定したステータスの試合の検索
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
' 入力された選手の試合一覧を取得する
' 一覧を修正画面上に表示
' 変更されたスコアを探す
' 勝敗が変わるかどうかのチェック
' 勝敗が変わらない場合、トーナメントシートにスコアを記載、黒線を引く、赤線を引く、で対応終了
' 勝敗が変わる場合、その先の試合が終了済みか確認
' 先の試合が始まっていない場合は、上記同様処理、採点票の状態を確認、印刷済みだったら、再印刷をかけるか聞く。再印刷する場合は、採点表ボタンを押すように指示
' 先の試合が終了済みの場合は、終了済みの試合を一覧で取得。この場合、対象選手が関わった試合のみを見るのか、トーナメントの最後まで見るのか...
' 次のように変更されますが、よろしいですか。それとも、この試合の結果を破棄
' N ○○ スコア 対 スコア ○○ N　→　N ○○ スコア 対 スコア △△ M
Sub aiu()
    setUp
    MsgBox findResult(1).winner
End Sub
