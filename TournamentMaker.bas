Attribute VB_Name = "TournamentMaker"
Sub makeTournament()

    setUp

    '--------------------------------------------------------------------------------
    Dim teams As Integer
    Dim baseTeams As Integer
    Dim maxNumPerPage As Integer
    Dim pageNum As Integer
    Dim roundEachPage As Integer
    Dim firstTeamNumberEachPage() As Integer
    Dim teamNumEachPage() As Integer
    Dim page As Integer
    Dim row As Integer
    Dim i As Integer
    Dim start As Integer
    Dim fin As Integer
    Dim match As Integer
    Dim round As Integer
    Dim tournaments() As Integer
    Dim maxRowperPage As Integer ' 各ページの実際の最大組数を保存。配列生成時に使用 TODO 似た変数があるので、考慮
    Dim betweenTwoLinesFlg As Boolean
    Dim betweenLinesStart As Integer
    Dim position(2, 2) As Integer
    Dim leftRow As Integer
    Dim rightRow As Integer
    Dim centerRow As Integer
    Dim startPlgNum As Integer
    
    

'おまじない
With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .Cursor = xlWait
End With
    
    teams = teamsRange
    maxTeamsNum = 0
    baseTeams = culNumberOfBaseTeams(teams)
    maxNumPerPage = maxNumPerPageRange
    roundEachPage = culNumberOfNeedRounds(teams)
    startPlgNum = plgStartNoRange.Value
    betweenTwoLinesFlg = False
    
    pageNum = getPageNumber(teams, maxNumPerPage)
    roundEachPage = roundEachPage - Log(pageNum) / Log(2)
    
    
    ' 各ページの最初の番号を取得
    ReDim firstTeamNumberEachPage(0 To pageNum, LEFT To RIGHT) As Integer
    Call getFirstTeamNumberEachPage(firstTeamNumberEachPage, teams, pageNum)
    
    ' 各ページのペア数を計算
    Debug.Print "各ページのペア数"
    ReDim teamNumEachPage(1 To pageNum, LEFT To RIGHT) As Integer
    For page = 1 To pageNum - 1
        ' 左側
        teamNumEachPage(page, LEFT) = firstTeamNumberEachPage(page, RIGHT) - firstTeamNumberEachPage(page, LEFT)
        If (maxRowperPage < teamNumEachPage(page, LEFT)) Then
            maxRowperPage = teamNumEachPage(page, LEFT)
        End If
        ' Debug.Print page & "ページ, 左側 " & teamNumEachPage(page, LEFT) & "組"
        
        ' 右側
        teamNumEachPage(page, RIGHT) = firstTeamNumberEachPage(page + 1, LEFT) - firstTeamNumberEachPage(page, RIGHT)
        If (maxRowperPage < teamNumEachPage(page, RIGHT)) Then
            maxRowperPage = teamNumEachPage(page, RIGHT)
        End If
        ' Debug.Print page & "ページ, 右側 " & teamNumEachPage(page, RIGHT) & "組"
    Next page
    ' 左側
    teamNumEachPage(page, LEFT) = firstTeamNumberEachPage(page, RIGHT) - firstTeamNumberEachPage(page, LEFT)
    If (maxRowperPage < teamNumEachPage(page, LEFT)) Then
        maxRowperPage = teamNumEachPage(page, LEFT)
    End If
    ' Debug.Print page & "ページ, 左側 " & teamNumEachPage(page, LEFT) & "組"
    ' 右側
    teamNumEachPage(page, RIGHT) = teams + startPlgNum - firstTeamNumberEachPage(page, RIGHT)
    If (maxRowperPage < teamNumEachPage(page, RIGHT)) Then
        maxRowperPage = teamNumEachPage(page, RIGHT)
    End If
    ' Debug.Print page & "ページ, 右側 " & teamNumEachPage(page, RIGHT) & "組"
    
    
    ' トーナメント作成----------------------------------------------------------------------
    
    ' クリア
    tournamentWS.Cells.Clear
    tournamentWS.ResetAllPageBreaks

    '概形作成
    row = 1
    For page = 1 To pageNum
        'tournamentWS.Cells(row, 1) = "種別（" & page & "）"
        For i = 1 To maxRowperPage
            ' 左側
            tournamentWS.Range(tournamentWS.Cells(row, G_numLeftCol), tournamentWS.Cells(row + 1, G_numLeftCol)).Merge
            If (i <= teamNumEachPage(page, LEFT)) Then
                With tournamentWS
                    .Cells(row, G_numLeftCol) = i + firstTeamNumberEachPage(page, LEFT) - 1
                    .Cells(row, G_teamLeftCol - 1) = "("
                    .Cells(row, G_teamLeftCol + 1) = ")"
                    .Cells(row + 1, G_teamLeftCol - 1) = "("
                    .Cells(row + 1, G_teamLeftCol + 1) = ")"
                End With
            End If
            
            ' 右側
            tournamentWS.Range(tournamentWS.Cells(row, G_numRightCol), tournamentWS.Cells(row + 1, G_numRightCol)).Merge
            If (i <= teamNumEachPage(page, RIGHT)) Then
                With tournamentWS
                    .Cells(row, G_numRightCol) = i + firstTeamNumberEachPage(page, RIGHT) - 1
                    .Cells(row, G_teamRightCol - 1) = "("
                    .Cells(row, G_teamRightCol + 1) = ")"
                    .Cells(row + 1, G_teamRightCol - 1) = "("
                    .Cells(row + 1, G_teamRightCol + 1) = ")"
                End With
            End If
            row = row + 2
        Next i
        tournamentWS.HPageBreaks.Add Range("A" & row) ' 改ページの挿入
    Next page
    
    ' 罫線の作成
    Debug.Print "罫線の作成"
    For page = 1 To pageNum
        ' 1回戦
        ' 左側
        start = baseTeams / 2 + (baseTeams / 4 / pageNum) * (page - 1) * 2
        fin = start + (baseTeams / 4 / pageNum) - 1
        Debug.Print page & "ページ, 左側S=" & start & ", F=" & fin
        
        index = 1
        For match = start To fin
            If (baseMatchesWS.Cells(match, 4) = "UNDECIDED") Then ' 1回戦あり
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    With .Range(.Cells(row, G_startTournamentArea), .Cells(row + 1, G_startTournamentArea))
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    End With
                    .Range(.Cells(row + 1, G_startTournamentArea + 1), .Cells(row + 1, G_startTournamentArea + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                Call setAddress(match, LEFT, row - 1, G_startTournamentArea + 1)
                Call setAddress(match, RIGHT, row + 2, G_startTournamentArea + 1)
                With tournamentWS
                    .Cells(row - 1, G_startTournamentArea + 1).HorizontalAlignment = xlLeft
                    .Cells(row - 1, G_startTournamentArea + 1).VerticalAlignment = xlBottom
                    .Cells(row + 2, G_startTournamentArea + 1).HorizontalAlignment = xlLeft
                    .Cells(row + 2, G_startTournamentArea + 1).VerticalAlignment = xlTop
                End With
                index = index + 2 ' 2組分次へ
            Else ' 不戦勝(注意：1回戦であるため、横線はwinnerを使用)
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    .Range(.Cells(row, G_startTournamentArea), .Cells(row, G_startTournamentArea + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                index = index + 1 ' 1組分次へ
            End If
        Next match
        
        ' 右側
        start = baseTeams / 2 + (baseTeams / 4 / pageNum) * ((page - 1) * 2 + 1)
        fin = start + (baseTeams / 4 / pageNum) - 1
        Debug.Print page & "ページ, 右側S=" & start & ", F=" & fin
        
        index = 1
        For match = start To fin
            If (baseMatchesWS.Cells(match, 4) = "UNDECIDED") Then ' 1回戦あり
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    With .Range(tournamentWS.Cells(row, G_endTournamentArea), tournamentWS.Cells(row + 1, G_endTournamentArea))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    End With
                    .Range(.Cells(row + 1, G_endTournamentArea - 1), .Cells(row + 1, G_endTournamentArea - 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                Call setAddress(match, LEFT, row - 1, G_endTournamentArea - 1)
                Call setAddress(match, RIGHT, row + 2, G_endTournamentArea - 1)
                With tournamentWS
                    .Cells(row - 1, G_endTournamentArea - 1).HorizontalAlignment = xlRight
                    .Cells(row - 1, G_endTournamentArea - 1).VerticalAlignment = xlBottom
                    .Cells(row + 2, G_endTournamentArea - 1).HorizontalAlignment = xlRight
                    .Cells(row + 2, G_endTournamentArea - 1).VerticalAlignment = xlTop
                End With
                index = index + 2 ' 2組分次へ
            Else ' 不戦勝(注意：1回戦であるため、横線はwinnerを使用)
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    .Range(.Cells(row, G_endTournamentArea), .Cells(row, G_endTournamentArea - 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                index = index + 1 ' 1組分次へ
            End If
        Next match
    Next page
    
    ' 2回戦以降
    match = 64
    For round = 2 To roundEachPage - 1
        match = baseTeams / (2 ^ round)
        For page = 1 To pageNum
            ' 左側
            For row = 1 + (maxRowperPage * (page - 1) * 2) To maxRowperPage * page * 2
                If (tournamentWS.Range(tournamentWS.Cells(row, G_startTournamentArea + round - 1), tournamentWS.Cells(row, G_startTournamentArea + round - 1)).Borders.Value = -4142) Then ' 枠線なし
                Else
                    If (betweenTwoLinesFlg = False) Then
                        betweenTwoLinesFlg = True
                        betweenLinesStart = row
                        Call setAddress(match, LEFT, row - 1, G_startTournamentArea + round)
                        With tournamentWS
                            .Cells(row - 1, G_startTournamentArea + round).HorizontalAlignment = xlLeft
                            .Cells(row - 1, G_startTournamentArea + round).VerticalAlignment = xlBottom
                        End With
                    Else
                        betweenTwoLinesFlg = False
                        If (match Mod 2 = 0) Then
                            centerRow = culCenter(betweenLinesStart, row - 1, True)
                        Else
                            centerRow = culCenter(betweenLinesStart, row - 1, False)
                        End If
                        tournamentWS.Range(tournamentWS.Cells(centerRow, G_startTournamentArea + round), tournamentWS.Cells(centerRow, G_startTournamentArea + round)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Call setAddress(match, RIGHT, row, G_startTournamentArea + round)
                        With tournamentWS
                            .Cells(row, G_startTournamentArea + round).HorizontalAlignment = xlLeft
                            .Cells(row, G_startTournamentArea + round).VerticalAlignment = xlTop
                        End With
                        match = match + 1
                    End If
                End If
                If (betweenTwoLinesFlg = True) Then
                    tournamentWS.Range(tournamentWS.Cells(row, G_startTournamentArea + round - 1), tournamentWS.Cells(row, G_startTournamentArea + round - 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
                End If
            Next row
            
            ' 右側
            For row = 1 + (maxRowperPage * (page - 1) * 2) To maxRowperPage * page * 2
                If (tournamentWS.Range(tournamentWS.Cells(row, G_endTournamentArea - round + 1), tournamentWS.Cells(row, G_endTournamentArea - round + 1)).Borders.Value = -4142) Then ' 枠線なし
                Else
                    If (betweenTwoLinesFlg = False) Then
                        betweenTwoLinesFlg = True
                        betweenLinesStart = row
                        Call setAddress(match, LEFT, row - 1, G_endTournamentArea - round)
                        With tournamentWS
                            .Cells(row - 1, G_endTournamentArea - round).HorizontalAlignment = xlRight
                            .Cells(row - 1, G_endTournamentArea - round).VerticalAlignment = xlBottom
                        End With
                    Else
                        betweenTwoLinesFlg = False
                        If (match Mod 2 = 0) Then
                            centerRow = culCenter(betweenLinesStart, row - 1, True)
                        Else
                            centerRow = culCenter(betweenLinesStart, row - 1, False)
                        End If
                        tournamentWS.Range(tournamentWS.Cells(centerRow, G_endTournamentArea - round), tournamentWS.Cells(centerRow, G_endTournamentArea - round)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Call setAddress(match, RIGHT, row, G_endTournamentArea - round)
                        With tournamentWS
                            .Cells(row, G_endTournamentArea - round).HorizontalAlignment = xlRight
                            .Cells(row, G_endTournamentArea - round).VerticalAlignment = xlTop
                        End With

                        match = match + 1
                    End If
                End If
                If (betweenTwoLinesFlg = True) Then
                    tournamentWS.Range(tournamentWS.Cells(row, G_endTournamentArea - round + 1), tournamentWS.Cells(row, G_endTournamentArea - round + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
                End If
            Next row
        Next page
    Next round
    
    ' 仕上げ
    Debug.Print "仕上げ"
    match = baseTeams / (2 ^ roundEachPage)
    For page = 1 To pageNum
        Debug.Print "page=" & page
        ' 各ページの最終ラウンドの罫線(横線)を合わせる-----------------------
        For row = 1 + (maxRowperPage * (page - 1) * 2) To maxRowperPage * page * 2 ' 左右の最後の所を探す
            If (tournamentWS.Range(tournamentWS.Cells(row, G_startTournamentArea + round - 1), tournamentWS.Cells(row, G_startTournamentArea + round - 1)).Borders.Value = -4142) Then
            Else
                Debug.Print "左：" & row
                leftRow = row
            End If
            If (tournamentWS.Range(tournamentWS.Cells(row, G_endTournamentArea - round + 1), tournamentWS.Cells(row, G_endTournamentArea - round + 1)).Borders.Value = -4142) Then
            Else
                Debug.Print "右：" & row
                rightRow = row
            End If
        Next row
        If (leftRow <> rightRow) Then
            ' 左側に合わせるので、右側の罫線を消去
            With tournamentWS
                .Range(.Cells(rightRow, G_endTournamentArea - round + 1), .Cells(rightRow, G_endTournamentArea - round + 1)).Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            End With
        End If
        With tournamentWS
            .Range(.Cells(leftRow, G_startTournamentArea + round - 1), .Cells(leftRow, G_endTournamentArea - round + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
        Call setAddress(match, LEFT, leftRow, G_startTournamentArea + round - 2)
        Call setAddress(match, RIGHT, leftRow, G_endTournamentArea - round + 2)
        match = match + 1
        
        
        '各ページ外の対戦の罫線（決勝等）------------------------------------
        Call drawBorders(pageNum, page, leftRow, maxRowperPage)
        
    Next page
    
    ' フォントの設定
    For i = G_startTournamentArea To G_endTournamentArea
        With tournamentWS.Columns(i).Font
            .Name = "ＭＳ Ｐゴシック"
            .Size = 8
            .Color = RGB(255, 0, 0)
        End With
    Next i
    
    ' 体裁の調整
    tournamentWS.Columns(G_numLeftCol).HorizontalAlignment = xlCenter
    tournamentWS.Columns(G_numRightCol).HorizontalAlignment = xlCenter
        
'おまじない解除
With Application
        .Cursor = xlDefault
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
End With

End Sub

Function getFirstTeamNumberEachPage(ByRef firstTeamNumberEachPage() As Integer, teams As Integer, pageNum As Integer)
    
    Dim step As Integer
    Dim match As Integer
    Dim page As Integer
    Dim baseTeams As Integer
    
    baseTeams = culNumberOfBaseTeams(teams)
    step = baseTeams / 2 / pageNum
    page = 1
    
    For match = baseTeams / 2 To baseTeams - 1 Step step
        firstTeamNumberEachPage(page, LEFT) = getLeftLimit(match) ' ページ左側
        firstTeamNumberEachPage(page, RIGHT) = getLeftLimit(match + step / 2) ' ページ右側
        page = page + 1
    Next match

End Function

Function getPageNumber(ByVal teams As Integer, maxNumPerPage As Integer) As Integer
    Dim pageNumber As Integer
    pageNumber = 1
    
    
    If (teams > maxNumPerPage) Then
        pageNumber = WorksheetFunction.RoundUp(Log(teams / maxNumPerPage) / Log(2), 0)
        pageNumber = 2 ^ pageNumber
        Debug.Print pageNumber
    End If
    
    getPageNumber = pageNumber
End Function

' ページをまたぐトーナメントの場合の、各ページの真ん中の罫線を作成
Function drawBorders(pageNum As Integer, page As Integer, leftRow As Integer, maxRowperPage As Integer)
    Dim col As Integer
    Dim topBorder As Boolean
    Dim position(1, 1) As Integer
    Dim middleLeftCol As Integer
    Dim middleRightCol As Integer
    
    middleLeftCol = WorksheetFunction.RoundUp((G_endTournamentArea - G_startTournamentArea) / 2, 0) + G_startTournamentArea - 1
    middleRightCol = WorksheetFunction.RoundUp((G_endTournamentArea - G_startTournamentArea) / 2, 0) + G_startTournamentArea

    ' ページ数が1ページの場合、特別処理
    If (pageNum = 1) Then
        With tournamentWS
            .Range(.Cells(leftRow - 1, middleRightCol), .Cells(leftRow - maxRowperPage / 2, middleRightCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        Exit Function
    End If
    
    ' 左右の列とトップ・ボトムの設定
    If (page <= pageNum / 2) Then
        col = middleRightCol
    Else
        col = middleLeftCol
    End If
    
    If (page Mod 2 = 0) Then
        topBorder = True ' ページ番号が偶数のときは、上の罫線
        position(0, 0) = leftRow - 1
        position(1, 0) = maxRowperPage * (page - 1) * 2 + 1
    Else
        topBorder = False ' ページ番号が奇数のときは、下の罫線
        position(0, 0) = leftRow
        position(1, 0) = maxRowperPage * page * 2
    End If
    
    ' 罫線を描画
    With tournamentWS
        With .Range(.Cells(position(0, 0), col), .Cells(position(1, 0), col))
            If topBorder Then
                .Borders(xlEdgeTop).LineStyle = xlContinuous
            Else
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
            End If
            If col = middleRightCol Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
            Else
                .Borders(xlEdgeRight).LineStyle = xlContinuous
            End If
        End With
    End With
End Function
' 結果に合わせて罫線を引く
' 引数
' 　baseMatchID : ベースマッチID(罫線の座標指定に使用)
' 　startRow : 2線間の内側の上側の行
' 　endRow : 2線間の内側の下側の行
' 　col : 内側の列
' 　winningSide : 勝者のサイド
' 　whichSide : 各ページトーナメントの左右
Function drawResultLine(baseMatchID As Integer, startRow As Integer, endRow As Integer, col As Integer, winningSide As Integer, whichSide As Integer)

    Dim startCol As Integer
    Dim center As Integer
    
    startCol = col
    
    center = culCenter(startRow, endRow, (baseMatchID Mod 2 = 0))
    
    ' 横方向の罫線の幅の調整
    Call adjustStartColForSide(startCol, startRow, endRow, col, xlEdgeTop, xlEdgeBottom, winningSide, whichSide)
    MsgBox "startCol=" & startCol & ", col=" & col
    
    If (whichSide = LEFT) Then
        If (winningSide = LEFT) Then
            Call drawRedBorders(startRow, center - 1, startCol, col, xlEdgeRight)
            Call drawRedBorders(startRow, center - 1, startCol, col, xlEdgeTop)
            Call drawRedBorders(center, center, col + 1, col + 1, xlEdgeTop)
        Else
            Call drawRedBorders(center, endRow, startCol, col, xlEdgeRight)
            Call drawRedBorders(center, endRow, startCol, col, xlEdgeBottom)
            Call drawRedBorders(center, center, col + 1, col + 1, xlEdgeTop)
        End If
    Else
        If (winningSide = LEFT) Then
            Call drawRedBorders(startRow, center - 1, col, startCol, xlEdgeLeft)
            Call drawRedBorders(startRow, center - 1, col, startCol, xlEdgeTop)
            Call drawRedBorders(center, center, col - 1, col - 1, xlEdgeTop)
        Else
            Call drawRedBorders(center, endRow, col, startCol, xlEdgeLeft)
            Call drawRedBorders(center, endRow, col, startCol, xlEdgeBottom)
            Call drawRedBorders(center, center, col - 1, col - 1, xlEdgeTop)
        End If
    End If
End Function

Sub abcde()
    setUp
    Call drawResultLine(2, 2, 4, 7, LEFT, LEFT)
    Call drawResultLine(33, 4, 5, 6, LEFT, LEFT)
    Call drawResultLine(38, 22, 23, 6, RIGHT, LEFT)
    Call drawResultLine(9, 17, 24, 8, RIGHT, LEFT)
    Call drawResultLine(4, 7, 20, 9, RIGHT, LEFT)
    Call drawResultLine(19, 23, 25, 7, LEFT, LEFT)
'    Call drawResultLine(18, 19, 6, RIGHT, LEFT)
'    Call drawResultLine(16, 18, 7, LEFT, LEFT)
'    Call drawResultLine(51, 53, 7, RIGHT, LEFT)
'    Call drawResultLine(4, 5, 19, LEFT, RIGHT)
'    Call drawResultLine(2, 4, 18, LEFT, RIGHT)
'    Call drawResultLine(8, 9, 19, LEFT, RIGHT)
'    Call drawResultLine(12, 13, 19, RIGHT, RIGHT)
'    Call drawResultLine(9, 12, 18, LEFT, RIGHT)
'    Call drawResultLine(3, 10, 17, LEFT, RIGHT)
End Sub
' 2線間の中央を算出する。
' 戻り値行のBorder(xlTop)に罫線を引くこと。
' 引数：2線間の内側2セルの行番号
' 　　　切上げ切捨てフラグ
Function culCenter(topRow As Integer, bottomRow As Integer, isFloor As Boolean) As Integer
    If (isFloor = True) Then
        culCenter = WorksheetFunction.RoundDown((bottomRow + 1 - topRow) / 2 + topRow, 0)
    Else
        culCenter = WorksheetFunction.RoundUp((bottomRow + 1 - topRow) / 2 + topRow, 0)
    End If
End Function
' 罫線の横幅を調整する
Function adjustStartColForSide(ByRef startCol As Integer, startRow As Integer, endRow As Integer, col As Integer, edgeTop As XlBordersIndex, edgeBottom As XlBordersIndex, winningSide As Integer, whichSide As Integer)
    If whichSide = LEFT Then
        If winningSide = LEFT Then
            While tournamentWS.Range(tournamentWS.Cells(startRow, startCol - 1), tournamentWS.Cells(startRow, startCol - 1)).Borders(edgeTop).LineStyle = xlContinuous
                startCol = startCol - 1
            Wend
        Else
            While tournamentWS.Range(tournamentWS.Cells(endRow, startCol - 1), tournamentWS.Cells(endRow, startCol - 1)).Borders(edgeBottom).LineStyle = xlContinuous
                startCol = startCol - 1
            Wend
        End If
    Else
        If winningSide = LEFT Then
            While tournamentWS.Range(tournamentWS.Cells(startRow, startCol + 1), tournamentWS.Cells(startRow, startCol + 1)).Borders(edgeTop).LineStyle = xlContinuous
                startCol = startCol + 1
            Wend
        Else
            While tournamentWS.Range(tournamentWS.Cells(endRow, startCol + 1), tournamentWS.Cells(endRow, startCol + 1)).Borders(edgeBottom).LineStyle = xlContinuous
                startCol = startCol + 1
            Wend
        End If
    End If
End Function
' 罫線を描画する。
' 4つ角と罫線の位置を指定する、
Function drawRedBorders(startRow As Integer, endRow As Integer, startCol As Integer, endCol As Integer, selectedLine As XlBordersIndex)
    With tournamentWS.Range(tournamentWS.Cells(startRow, startCol), tournamentWS.Cells(endRow, endCol)).Borders(selectedLine)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(255, 0, 0)
    End With
End Function
' トーナメントのヘッダーを設定
Function setHeader()
    tournamentWS.PageSetup.CenterHeader = categoryRange.Value & " (&P)"
End Function
