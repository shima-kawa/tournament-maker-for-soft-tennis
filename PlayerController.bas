Attribute VB_Name = "PlayerController"
Sub insertPlayerInformation()

    setUp
    
    '--------------------------------------------------------------------------------
    
    Dim row As Integer
    Dim lastRow As Integer
    
    Dim p As player

    xlLastRow = playerListWS.Cells(Rows.count, plgNoCol).row
    lastRow = playerListWS.Cells(xlLastRow, plgNoCol).End(xlUp).row
    
    If (teamsRange.Value <> lastRow + 1) Then
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
        tournamentWS.Cells(row, G_teamLeftCol) = p.ATeam
        tournamentWS.Cells(row + 1, G_teamLeftCol) = p.BTeam
        
    Next row
    
    ' �E��
    xlLastRow = tournamentWS.Cells(Rows.count, G_numRightCol).row  'Excel�̍ŏI�s���擾
    lastRow = tournamentWS.Cells(xlLastRow, G_numRightCol).End(xlUp).row   '�ŏI�s���擾
    For row = 1 To lastRow Step 2
        Set p = findPlayer(tournamentWS.Cells(row, G_numRightCol))
        tournamentWS.Cells(row, G_nameRightCol) = p.AName
        tournamentWS.Cells(row + 1, G_nameRightCol) = p.BName
        tournamentWS.Cells(row, G_teamRightCol) = p.ATeam
        tournamentWS.Cells(row + 1, G_teamRightCol) = p.BTeam
    Next row
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
