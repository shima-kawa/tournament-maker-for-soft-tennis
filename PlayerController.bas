Attribute VB_Name = "PlayerController"
Sub insertPlayerInformation()

    setUp
    
    '--------------------------------------------------------------------------------
    Application.DisplayAlerts = False  '--- �m�F���b�Z�[�W���\��
    
    Dim row As Integer
    Dim lastRow As Integer
    
    Dim p As player

    xlLastRow = playerListWS.Cells(Rows.count, plgNoCol).row
    lastRow = playerListWS.Cells(xlLastRow, plgNoCol).End(xlUp).row
    
    If (teamsRange.Value <> lastRow - 1) Then
        MsgBox "�G���[�B�I��ꗗ�̑g���ƎQ���g������v���܂���B"
        Exit Sub
    End If
    
    
    ' ����
    xlLastRow = tournamentWS.Cells(Rows.count, G_numLeftCol).row  'Excel�̍ŏI�s���擾
    lastRow = tournamentWS.Cells(xlLastRow, G_numLeftCol).End(xlUp).row   '�ŏI�s���擾
    For row = 1 To lastRow Step 2
        Set p = findPlayer(tournamentWS.Cells(row, G_numLeftCol))
        tournamentWS.Cells(row, G_nameLeftCol) = p.AName
        tournamentWS.Cells(row + 1, G_nameLeftCol) = p.BName
        With tournamentWS
            .Cells(row, G_nameLeftCol).VerticalAlignment = xlBottom
            .Cells(row + 1, G_nameLeftCol).VerticalAlignment = xlTop
        End With
        If (p.ATeam = p.BTeam) Then
            With tournamentWS
                .Range(.Cells(row, G_teamLeftCol), .Cells(row + 1, G_teamLeftCol)).Merge
                .Range(.Cells(row, G_teamLeftCol - 1), .Cells(row + 1, G_teamLeftCol - 1)).Merge
                .Range(.Cells(row, G_teamLeftCol + 1), .Cells(row + 1, G_teamLeftCol + 1)).Merge
                .Cells(row, G_teamLeftCol) = p.ATeam
            End With
        Else
            With tournamentWS
                .Cells(row, G_teamLeftCol) = p.ATeam
                .Cells(row + 1, G_teamLeftCol) = p.BTeam
                .Cells(row, G_teamLeftCol - 1).VerticalAlignment = xlBottom
                .Cells(row, G_teamLeftCol).VerticalAlignment = xlBottom
                .Cells(row, G_teamLeftCol + 1).VerticalAlignment = xlBottom
                .Cells(row + 1, G_teamLeftCol - 1).VerticalAlignment = xlTop
                .Cells(row + 1, G_teamLeftCol).VerticalAlignment = xlTop
                .Cells(row + 1, G_teamLeftCol + 1).VerticalAlignment = xlTop
            End With
        End If
        
    Next row
    
    ' �E��
    xlLastRow = tournamentWS.Cells(Rows.count, G_numRightCol).row  'Excel�̍ŏI�s���擾
    lastRow = tournamentWS.Cells(xlLastRow, G_numRightCol).End(xlUp).row   '�ŏI�s���擾
    For row = 1 To lastRow Step 2
        Set p = findPlayer(tournamentWS.Cells(row, G_numRightCol))
        tournamentWS.Cells(row, G_nameRightCol) = p.AName
        tournamentWS.Cells(row + 1, G_nameRightCol) = p.BName
        With tournamentWS
            .Cells(row, G_nameRightCol).VerticalAlignment = xlBottom
            .Cells(row + 1, G_nameRightCol).VerticalAlignment = xlTop
        End With
        If (p.ATeam = p.BTeam) Then
            With tournamentWS
                .Range(.Cells(row, G_teamRightCol), .Cells(row + 1, G_teamRightCol)).Merge
                .Range(.Cells(row, G_teamRightCol - 1), .Cells(row + 1, G_teamRightCol - 1)).Merge
                .Range(.Cells(row, G_teamRightCol + 1), .Cells(row + 1, G_teamRightCol + 1)).Merge
                .Cells(row, G_teamRightCol) = p.ATeam
            End With
        Else
            With tournamentWS
                .Cells(row, G_teamRightCol) = p.ATeam
                .Cells(row + 1, G_teamRightCol) = p.BTeam
                .Cells(row, G_teamRightCol - 1).VerticalAlignment = xlBottom
                .Cells(row, G_teamRightCol).VerticalAlignment = xlBottom
                .Cells(row, G_teamRightCol + 1).VerticalAlignment = xlBottom
                .Cells(row + 1, G_teamRightCol - 1).VerticalAlignment = xlTop
                .Cells(row + 1, G_teamRightCol).VerticalAlignment = xlTop
                .Cells(row + 1, G_teamRightCol + 1).VerticalAlignment = xlTop
            End With
        End If
    Next row
    Application.DisplayAlerts = True   '--- �m�F���b�Z�[�W��\��
    
    isInsertedPlayerInfo.Value = "��"
End Sub
Function findPlayer(plgNo As Integer) As player
    
    Dim p As player
    Set p = New player
    
    p.programNum = plgNo
    
    Set res = playerListWS.Range("A:A").Find(plgNo, LookAt:=xlWhole, SearchOrder:=xlByRows)
    
    If (res Is Nothing) Then
        MsgBox "�G���[: �v���O����No" & plgNo & "�ɑΉ�����I�肪������܂���B"
    Else
        p.AName = playerListWS.Cells(res.row, playerANameCol)
        p.BName = playerListWS.Cells(res.row, playerBNameCol)
        p.ATeam = playerListWS.Cells(res.row, playerATeamCol)
        p.BTeam = playerListWS.Cells(res.row, playerBTeamCol)
    End If
    
    Set findPlayer = p
End Function

Function insertPlayer(plgNo As Integer, p As player)
    Dim row As Integer
    Dim res As Range
    ' �v���O����No�̌���
    Set res = playerListWS.Range("A:A").Find(plgNo, LookAt:=xlWhole, SearchOrder:=xlByRows)

    If (res Is Nothing) Then
        MsgBox "�G���[: �v���O����No" & plgNo & "�ɑΉ�����s��������܂���B"
        Exit Function
    End If
    playerListWS.Cells(res.row, playerANameCol) = p.AName
    playerListWS.Cells(res.row, playerBNameCol) = p.BName
    playerListWS.Cells(res.row, playerATeamCol) = p.ATeam
    playerListWS.Cells(res.row, playerBTeamCol) = p.BTeam
End Function
