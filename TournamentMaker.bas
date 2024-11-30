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
    Dim Match As Integer
    Dim round As Integer
    Dim tournaments() As Integer
    Dim maxRowperPage As Integer ' �e�y�[�W�̎��ۂ̍ő�g����ۑ��B�z�񐶐����Ɏg�p TODO �����ϐ�������̂ŁA�l��
    Dim betweenTwoLinesFlg As Boolean
    Dim betweenLinesStart As Integer
    Dim position(2, 2) As Integer
    Dim leftRow As Integer
    Dim rightRow As Integer
    Dim centerRow As Integer
    Dim startPlgNum As Integer
    
    

'���܂��Ȃ�
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
    
    
    ' �e�y�[�W�̍ŏ��̔ԍ����擾
    ReDim firstTeamNumberEachPage(0 To pageNum, LEFT To RIGHT) As Integer
    Call getFirstTeamNumberEachPage(firstTeamNumberEachPage, teams, pageNum)
    
    ' �e�y�[�W�̃y�A�����v�Z
    Debug.Print "�e�y�[�W�̃y�A��"
    ReDim teamNumEachPage(1 To pageNum, LEFT To RIGHT) As Integer
    For page = 1 To pageNum - 1
        ' ����
        teamNumEachPage(page, LEFT) = firstTeamNumberEachPage(page, RIGHT) - firstTeamNumberEachPage(page, LEFT)
        If (maxRowperPage < teamNumEachPage(page, LEFT)) Then
            maxRowperPage = teamNumEachPage(page, LEFT)
        End If
        ' Debug.Print page & "�y�[�W, ���� " & teamNumEachPage(page, LEFT) & "�g"
        
        ' �E��
        teamNumEachPage(page, RIGHT) = firstTeamNumberEachPage(page + 1, LEFT) - firstTeamNumberEachPage(page, RIGHT)
        If (maxRowperPage < teamNumEachPage(page, RIGHT)) Then
            maxRowperPage = teamNumEachPage(page, RIGHT)
        End If
        ' Debug.Print page & "�y�[�W, �E�� " & teamNumEachPage(page, RIGHT) & "�g"
    Next page
    ' ����
    teamNumEachPage(page, LEFT) = firstTeamNumberEachPage(page, RIGHT) - firstTeamNumberEachPage(page, LEFT)
    If (maxRowperPage < teamNumEachPage(page, LEFT)) Then
        maxRowperPage = teamNumEachPage(page, LEFT)
    End If
    ' Debug.Print page & "�y�[�W, ���� " & teamNumEachPage(page, LEFT) & "�g"
    ' �E��
    teamNumEachPage(page, RIGHT) = teams + startPlgNum - firstTeamNumberEachPage(page, RIGHT)
    If (maxRowperPage < teamNumEachPage(page, RIGHT)) Then
        maxRowperPage = teamNumEachPage(page, RIGHT)
    End If
    ' Debug.Print page & "�y�[�W, �E�� " & teamNumEachPage(page, RIGHT) & "�g"
    
    
    ' �g�[�i�����g�쐬----------------------------------------------------------------------
    
    ' �N���A
    tournamentWS.Cells.Clear
    tournamentWS.ResetAllPageBreaks

    '�T�`�쐬
    row = 1
    For page = 1 To pageNum
        'tournamentWS.Cells(row, 1) = "��ʁi" & page & "�j"
        For i = 1 To maxRowperPage
            ' ����
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
            
            ' �E��
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
        tournamentWS.HPageBreaks.Add Range("A" & row) ' ���y�[�W�̑}��
    Next page
    
    ' �r���̍쐬
    Debug.Print "�r���̍쐬"
    For page = 1 To pageNum
        ' 1���
        ' ����
        start = baseTeams / 2 + (baseTeams / 4 / pageNum) * (page - 1) * 2
        fin = start + (baseTeams / 4 / pageNum) - 1
        Debug.Print page & "�y�[�W, ����S=" & start & ", F=" & fin
        
        index = 1
        For Match = start To fin
            If (baseMatchesWS.Cells(Match, 4) = "UNDECIDED") Then ' 1��킠��
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    With .Range(.Cells(row, G_startTournamentArea), .Cells(row + 1, G_startTournamentArea))
                        .Borders(xlEdgeRight).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    End With
                    .Range(.Cells(row + 1, G_startTournamentArea + 1), .Cells(row + 1, G_startTournamentArea + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                Call setAddress(Match, LEFT, row - 1, G_startTournamentArea + 1)
                Call setAddress(Match, RIGHT, row + 2, G_startTournamentArea + 1)
                With tournamentWS
                    .Cells(row - 1, G_startTournamentArea + 1).HorizontalAlignment = xlLeft
                    .Cells(row - 1, G_startTournamentArea + 1).VerticalAlignment = xlBottom
                    .Cells(row + 2, G_startTournamentArea + 1).HorizontalAlignment = xlLeft
                    .Cells(row + 2, G_startTournamentArea + 1).VerticalAlignment = xlTop
                End With
                index = index + 2 ' 2�g������
            Else ' �s�폟(���ӁF1���ł��邽�߁A������winner���g�p)
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    .Range(.Cells(row, G_startTournamentArea), .Cells(row, G_startTournamentArea + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                index = index + 1 ' 1�g������
            End If
        Next Match
        
        ' �E��
        start = baseTeams / 2 + (baseTeams / 4 / pageNum) * ((page - 1) * 2 + 1)
        fin = start + (baseTeams / 4 / pageNum) - 1
        Debug.Print page & "�y�[�W, �E��S=" & start & ", F=" & fin
        
        index = 1
        For Match = start To fin
            If (baseMatchesWS.Cells(Match, 4) = "UNDECIDED") Then ' 1��킠��
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    With .Range(tournamentWS.Cells(row, G_endTournamentArea), tournamentWS.Cells(row + 1, G_endTournamentArea))
                        .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Borders(xlEdgeTop).LineStyle = xlContinuous
                        .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    End With
                    .Range(.Cells(row + 1, G_endTournamentArea - 1), .Cells(row + 1, G_endTournamentArea - 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                Call setAddress(Match, LEFT, row - 1, G_endTournamentArea - 1)
                Call setAddress(Match, RIGHT, row + 2, G_endTournamentArea - 1)
                With tournamentWS
                    .Cells(row - 1, G_endTournamentArea - 1).HorizontalAlignment = xlRight
                    .Cells(row - 1, G_endTournamentArea - 1).VerticalAlignment = xlBottom
                    .Cells(row + 2, G_endTournamentArea - 1).HorizontalAlignment = xlRight
                    .Cells(row + 2, G_endTournamentArea - 1).VerticalAlignment = xlTop
                End With
                index = index + 2 ' 2�g������
            Else ' �s�폟(���ӁF1���ł��邽�߁A������winner���g�p)
                row = (maxRowperPage * (page - 1) + index) * 2
                With tournamentWS
                    .Range(.Cells(row, G_endTournamentArea), .Cells(row, G_endTournamentArea - 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                End With
                index = index + 1 ' 1�g������
            End If
        Next Match
    Next page
    
    ' 2���ȍ~
    Match = 64
    For round = 2 To roundEachPage - 1
        Match = baseTeams / (2 ^ round)
        For page = 1 To pageNum
            ' ����
            For row = 1 + (maxRowperPage * (page - 1) * 2) To maxRowperPage * page * 2
                If (tournamentWS.Range(tournamentWS.Cells(row, G_startTournamentArea + round - 1), tournamentWS.Cells(row, G_startTournamentArea + round - 1)).Borders.Value = -4142) Then ' �g���Ȃ�
                Else
                    If (betweenTwoLinesFlg = False) Then
                        betweenTwoLinesFlg = True
                        betweenLinesStart = row
                        Call setAddress(Match, LEFT, row - 1, G_startTournamentArea + round)
                        With tournamentWS
                            .Cells(row - 1, G_startTournamentArea + round).HorizontalAlignment = xlLeft
                            .Cells(row - 1, G_startTournamentArea + round).VerticalAlignment = xlBottom
                        End With
                    Else
                        betweenTwoLinesFlg = False
                        If (Match Mod 2 = 0) Then
                            centerRow = culCenter(betweenLinesStart, row - 1, True)
                        Else
                            centerRow = culCenter(betweenLinesStart, row - 1, False)
                        End If
                        tournamentWS.Range(tournamentWS.Cells(centerRow, G_startTournamentArea + round), tournamentWS.Cells(centerRow, G_startTournamentArea + round)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Call setAddress(Match, RIGHT, row, G_startTournamentArea + round)
                        With tournamentWS
                            .Cells(row, G_startTournamentArea + round).HorizontalAlignment = xlLeft
                            .Cells(row, G_startTournamentArea + round).VerticalAlignment = xlTop
                        End With
                        Match = Match + 1
                    End If
                End If
                If (betweenTwoLinesFlg = True) Then
                    tournamentWS.Range(tournamentWS.Cells(row, G_startTournamentArea + round - 1), tournamentWS.Cells(row, G_startTournamentArea + round - 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
                End If
            Next row
            
            ' �E��
            For row = 1 + (maxRowperPage * (page - 1) * 2) To maxRowperPage * page * 2
                If (tournamentWS.Range(tournamentWS.Cells(row, G_endTournamentArea - round + 1), tournamentWS.Cells(row, G_endTournamentArea - round + 1)).Borders.Value = -4142) Then ' �g���Ȃ�
                Else
                    If (betweenTwoLinesFlg = False) Then
                        betweenTwoLinesFlg = True
                        betweenLinesStart = row
                        Call setAddress(Match, LEFT, row - 1, G_endTournamentArea - round)
                        With tournamentWS
                            .Cells(row - 1, G_endTournamentArea - round).HorizontalAlignment = xlRight
                            .Cells(row - 1, G_endTournamentArea - round).VerticalAlignment = xlBottom
                        End With
                    Else
                        betweenTwoLinesFlg = False
                        If (Match Mod 2 = 0) Then
                            centerRow = culCenter(betweenLinesStart, row - 1, True)
                        Else
                            centerRow = culCenter(betweenLinesStart, row - 1, False)
                        End If
                        tournamentWS.Range(tournamentWS.Cells(centerRow, G_endTournamentArea - round), tournamentWS.Cells(centerRow, G_endTournamentArea - round)).Borders(xlEdgeTop).LineStyle = xlContinuous
                        Call setAddress(Match, RIGHT, row, G_endTournamentArea - round)
                        With tournamentWS
                            .Cells(row, G_endTournamentArea - round).HorizontalAlignment = xlRight
                            .Cells(row, G_endTournamentArea - round).VerticalAlignment = xlTop
                        End With

                        Match = Match + 1
                    End If
                End If
                If (betweenTwoLinesFlg = True) Then
                    tournamentWS.Range(tournamentWS.Cells(row, G_endTournamentArea - round + 1), tournamentWS.Cells(row, G_endTournamentArea - round + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
                End If
            Next row
        Next page
    Next round
    
    ' �d�グ
    Debug.Print "�d�グ"
    Match = baseTeams / (2 ^ roundEachPage)
    For page = 1 To pageNum
        Debug.Print "page=" & page
        ' �e�y�[�W�̍ŏI���E���h�̌r��(����)�����킹��-----------------------
        For row = 1 + (maxRowperPage * (page - 1) * 2) To maxRowperPage * page * 2 ' ���E�̍Ō�̏���T��
            If (tournamentWS.Range(tournamentWS.Cells(row, G_startTournamentArea + round - 1), tournamentWS.Cells(row, G_startTournamentArea + round - 1)).Borders.Value = -4142) Then
            Else
                Debug.Print "���F" & row
                leftRow = row
            End If
            If (tournamentWS.Range(tournamentWS.Cells(row, G_endTournamentArea - round + 1), tournamentWS.Cells(row, G_endTournamentArea - round + 1)).Borders.Value = -4142) Then
            Else
                Debug.Print "�E�F" & row
                rightRow = row
            End If
        Next row
        If (leftRow <> rightRow) Then
            ' �����ɍ��킹��̂ŁA�E���̌r��������
            With tournamentWS
                .Range(.Cells(rightRow, G_endTournamentArea - round + 1), .Cells(rightRow, G_endTournamentArea - round + 1)).Borders(xlEdgeTop).LineStyle = xlLineStyleNone
            End With
        End If
        With tournamentWS
            .Range(.Cells(leftRow, G_startTournamentArea + round - 1), .Cells(leftRow, G_endTournamentArea - round + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
        End With
        Call setAddress(Match, LEFT, leftRow, G_startTournamentArea + round - 2)
        Call setAddress(Match, RIGHT, leftRow, G_endTournamentArea - round + 2)
        Match = Match + 1
        
        
        '�e�y�[�W�O�̑ΐ�̌r���i�������j------------------------------------
        Call drawBorders(pageNum, page, leftRow, maxRowperPage)
        
    Next page
    
    ' �t�H���g�̐ݒ�
    For i = G_startTournamentArea To G_endTournamentArea
        With tournamentWS.Columns(i).Font
            .Name = "HG�ۺ޼��M-PRO"
            .Size = 8
            .Color = RGB(255, 0, 0)
        End With
    Next i
    
    ' �̍ق̒���
    With tournamentWS
        .Cells.Font.Size = 9
        .Columns(G_numLeftCol).HorizontalAlignment = xlCenter
        .Columns(G_numRightCol).HorizontalAlignment = xlCenter
        .Columns(G_teamLeftCol - 1).HorizontalAlignment = xlRight
        .Columns(G_teamLeftCol).HorizontalAlignment = xlCenter
        .Columns(G_teamLeftCol + 1).HorizontalAlignment = xlLeft
        .Columns(G_teamRightCol - 1).HorizontalAlignment = xlRight
        .Columns(G_teamRightCol).HorizontalAlignment = xlCenter
        .Columns(G_teamRightCol + 1).HorizontalAlignment = xlLeft
        
        .Columns(G_numLeftCol).Font.Name = "HG�ۺ޼��M-PRO"
        .Columns(G_numRightCol).Font.Name = "HG�ۺ޼��M-PRO"
        .Columns(G_nameLeftCol).Font.Name = "HG�ۺ޼��M-PRO"
        .Columns(G_nameRightCol).Font.Name = "HG�ۺ޼��M-PRO"
        .Columns(G_teamLeftCol).Font.Name = "HG�ۺ޼��M-PRO"
        .Columns(G_teamRightCol).Font.Name = "HG�ۺ޼��M-PRO"
    End With
    
    isTournamentGeneratedRange.Value = "��"
    isEditModeRange.Value = ""
    isInsertedPlayerInfo.Value = ""
        
'���܂��Ȃ�����
With Application
        .Cursor = xlDefault
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
End With

End Sub

Function getFirstTeamNumberEachPage(ByRef firstTeamNumberEachPage() As Integer, teams As Integer, pageNum As Integer)
    
    Dim step As Integer
    Dim Match As Integer
    Dim page As Integer
    Dim baseTeams As Integer
    
    baseTeams = culNumberOfBaseTeams(teams)
    step = baseTeams / 2 / pageNum
    page = 1
    
    For Match = baseTeams / 2 To baseTeams - 1 Step step
        firstTeamNumberEachPage(page, LEFT) = getLeftLimit(Match) ' �y�[�W����
        firstTeamNumberEachPage(page, RIGHT) = getLeftLimit(Match + step / 2) ' �y�[�W�E��
        page = page + 1
    Next Match

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

' �y�[�W���܂����g�[�i�����g�̏ꍇ�́A�e�y�[�W�̐^�񒆂̌r�����쐬
Function drawBorders(pageNum As Integer, page As Integer, leftRow As Integer, maxRowperPage As Integer)
    Dim col As Integer
    Dim topBorder As Boolean
    Dim position(1, 1) As Integer
    Dim middleLeftCol As Integer
    Dim middleRightCol As Integer
    
    middleLeftCol = WorksheetFunction.RoundUp((G_endTournamentArea - G_startTournamentArea) / 2, 0) + G_startTournamentArea - 1
    middleRightCol = WorksheetFunction.RoundUp((G_endTournamentArea - G_startTournamentArea) / 2, 0) + G_startTournamentArea

    ' �y�[�W����1�y�[�W�̏ꍇ�A���ʏ���
    If (pageNum = 1) Then
        With tournamentWS
            .Range(.Cells(leftRow - 1, middleRightCol), .Cells(leftRow - maxRowperPage / 2, middleRightCol)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End With
        Exit Function
    End If
    
    ' ���E�̗�ƃg�b�v�E�{�g���̐ݒ�
    If (page <= pageNum / 2) Then
        col = middleRightCol
    Else
        col = middleLeftCol
    End If
    
    If (page Mod 2 = 0) Then
        topBorder = True ' �y�[�W�ԍ��������̂Ƃ��́A��̌r��
        position(0, 0) = leftRow - 1
        position(1, 0) = maxRowperPage * (page - 1) * 2 + 1
    Else
        topBorder = False ' �y�[�W�ԍ�����̂Ƃ��́A���̌r��
        position(0, 0) = leftRow
        position(1, 0) = maxRowperPage * page * 2
    End If
    
    ' �r����`��
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
' ���ʂɍ��킹�Čr��������
' ����
' �@baseMatchID : �x�[�X�}�b�`ID(�r���̍��W�w��Ɏg�p)
' �@startRow : 2���Ԃ̓����̏㑤�̍s
' �@endRow : 2���Ԃ̓����̉����̍s
' �@col : �����̗�
' �@winningSide : ���҂̃T�C�h
' �@whichSide : �e�y�[�W�g�[�i�����g�̍��E
Function drawResultLine(baseMatchID As Integer, startRow As Integer, endRow As Integer, col As Integer, winningSide As Integer, whichSide As Integer)

    Dim startCol As Integer
    Dim center As Integer
    Dim round As Integer
        
    center = culCenter(startRow, endRow, (baseMatchID Mod 2 = 0))
    
    ' ����(������) ####################################################################################################################
    
    ' �����`����֐��Ŕ����o���āA�r�����Z�b�g�p�֐��Ƃ���Ƃ�����������Ȃ��BMatchID�������ɂ��āB
    '1���F��E���E��or�E�E��������
    '2���F2�Z�������Ă���㉺�E��or�E�E��������
    '3���ȍ~�F��or�E�E����������
    
    round = culNumberOfNeedRounds(teamsRange.Value) - WorksheetFunction.RoundUp(Log(baseMatchID + 1) / Log(2), 0) + 1
    Select Case round
        Case 1
            Debug.Print ("1���̏���")
            ' �����v���C���[�̌r�� ��
            startCol = col
            Call adjustStartColForSide(startCol, startRow, endRow, col, xlEdgeTop, xlEdgeBottom, LEFT, whichSide)
            Call drawBlackBorders(startRow, center - 1, startCol, col, xlEdgeTop)
            ' �E���v���C���[�̌r�� ��
            startCol = col
            Call adjustStartColForSide(startCol, startRow, endRow, col, xlEdgeTop, xlEdgeBottom, RIGHT, whichSide)
            Call drawBlackBorders(center, endRow, startCol, col, xlEdgeBottom)
            
        Case 2
            Debug.Print ("2���̏���")
            ' �����v���C���[�̌r�� ��
            startCol = col
            Call adjustStartColForSide(startCol, startRow, endRow, col, xlEdgeTop, xlEdgeBottom, LEFT, whichSide)
            If startCol <> col Then
                '2�Z����(=startCol��col���s��v)�̏ꍇ�̂ݍ���
                Call drawBlackBorders(startRow, center - 1, startCol, col, xlEdgeTop)
            End If
            
            ' �E���v���C���[�̌r�� ��
            startCol = col
            Call adjustStartColForSide(startCol, startRow, endRow, col, xlEdgeTop, xlEdgeBottom, RIGHT, whichSide)
            If startCol <> col Then
                '2�Z����(=startCol��col���s��v)�̏ꍇ�̂ݍ���
                Call drawBlackBorders(center, endRow, startCol, col, xlEdgeBottom)
            End If
        Case Else
            Debug.Print ("3���ȍ~�̏���")
    End Select
    
    ' �c�A�����̌r�� ��
    If (whichSide = LEFT) Then
        Call drawBlackBorders(startRow, endRow, startCol, col, xlEdgeRight)
        Call drawBlackBorders(center, center, col + 1, col + 1, xlEdgeTop)

    Else
        Call drawBlackBorders(startRow, endRow, startCol, col, xlEdgeLeft)
        Call drawBlackBorders(center, center, col - 1, col - 1, xlEdgeTop)
    End If

    
    ' �Ԑ� ############################################################################################################################
    ' �������̌r���̕��̒���
    startCol = col
    Call adjustStartColForSide(startCol, startRow, endRow, col, xlEdgeTop, xlEdgeBottom, winningSide, whichSide)
    

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

' 2���Ԃ̒������Z�o����B
' �߂�l�s��Border(xlTop)�Ɍr�����������ƁB
' �����F2���Ԃ̓���2�Z���̍s�ԍ�
' �@�@�@�؏グ�؎̂ăt���O
Function culCenter(topRow As Integer, bottomRow As Integer, isFloor As Boolean) As Integer
    If (isFloor = True) Then
        culCenter = WorksheetFunction.RoundDown((bottomRow + 1 - topRow) / 2 + topRow, 0)
    Else
        culCenter = WorksheetFunction.RoundUp((bottomRow + 1 - topRow) / 2 + topRow, 0)
    End If
End Function
' �r���̉����𒲐�����
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
' �r����`�悷��B
' 4�p�ƌr���̈ʒu���w�肷��A
Function drawRedBorders(startRow As Integer, endRow As Integer, startCol As Integer, endCol As Integer, selectedLine As XlBordersIndex)
    With tournamentWS.Range(tournamentWS.Cells(startRow, startCol), tournamentWS.Cells(endRow, endCol)).Borders(selectedLine)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(255, 0, 0)
    End With
End Function
' �r����`�悷��B
' 4�p�ƌr���̈ʒu���w�肷��A
Function drawBlackBorders(startRow As Integer, endRow As Integer, startCol As Integer, endCol As Integer, selectedLine As XlBordersIndex)
    With tournamentWS.Range(tournamentWS.Cells(startRow, startCol), tournamentWS.Cells(endRow, endCol)).Borders(selectedLine)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(0, 0, 0)
    End With
End Function
' �g�[�i�����g�̃w�b�_�[��ݒ�
Function setHeader()
    tournamentWS.PageSetup.CenterHeader = categoryRange.Value & " (&P)"
End Function
