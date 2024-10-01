Attribute VB_Name = "Editor"
Option Explicit

Sub startEditMode()

    setUp
    
    Dim leftEntryNumCol As Integer
    Dim rightEntryNumCol As Integer
    Dim i As Integer
    Dim lastRow As Integer
    Dim entryPlayersRange As Range
    Dim ref As Range

    Dim msgRes As VbMsgBoxResult
    

    ' エラー処理-------------------------------------------------------------------------------
    If (isTournamentGeneratedRange.Value <> "済") Then
        MsgBox "トーナメントが作成されていません。トーナメントを作成してください。", _
            Title:="編集モード"
        Exit Sub
    End If
    
    If (isEditModeRange.Value = "済") Then
        ' temp
        MsgBox "組合せは決定済みです。リセットする際は、「トーナメント作成」ボタンを押してください。", _
            Title:="編集モード"
        Exit Sub
    End If
    If (isEditModeRange.Value = "途中") Then
        MsgBox "現在、編集モードです。", _
            Title:="編集モード"
        Exit Sub
    End If
    
    msgRes = MsgBox( _
        Prompt:="トーナメント編集モードへ入ります。よろしいですか？", _
        Buttons:=vbOKCancel, _
        Title:="編集モード" _
    )
    
    If (msgRes = vbCancel) Then
        Exit Sub
    End If
    
    ' エントリー名簿シートの確認----------------------------------------------------------------
    If (flgExsistSheet("エントリー名簿") = False) Then
        Worksheets.Add after:=Sheets(Worksheets.count)
        ActiveSheet.Name = "エントリー名簿"
        Set entryPlayersWS = ThisWorkbook.Worksheets("エントリー名簿")
        Call makeEntryPlayersSheet
    End If
    
    Set entryPlayersRange = entryPlayersWS.Range("A:E") ' TODO: 別の場所から変更したり、動的に設定できるようにする

    
    ' トーナメントシートの変更------------------------------------------------------------------
    leftEntryNumCol = G_numLeftCol
    rightEntryNumCol = G_numRightCol + 2
    tournamentWS.Columns(G_numLeftCol).Insert ' トーナメントの左側。右側はすでに空いているので、それを利用
    
    tournamentWS.Columns(leftEntryNumCol).ColumnWidth = 4
    tournamentWS.Columns(rightEntryNumCol).ColumnWidth = 4
    
    ' セルへの関数の挿入------------------------------------------------------------------------
    ' 左側
    lastRow = tournamentWS.Cells(tournamentWS.Rows.count, leftEntryNumCol + 1).End(xlUp).row
    For i = 1 To lastRow Step 2
        With tournamentWS
            With .Range(.Cells(i, leftEntryNumCol), .Cells(i + 1, leftEntryNumCol))
                .Merge
                .Interior.Color = vbYellow
                .Borders.LineStyle = xlContinuous
            End With
            Set ref = .Cells(i, leftEntryNumCol)
            .Cells(i, leftEntryNumCol + 2) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",2,FALSE),"""")"
            .Cells(i + 1, leftEntryNumCol + 2) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",3,FALSE),"""")"
            .Cells(i, leftEntryNumCol + 4) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",4,FALSE),"""")"
            .Cells(i + 1, leftEntryNumCol + 4) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",5,FALSE),"""")"
        End With
    Next i

    ' 右側
    lastRow = tournamentWS.Cells(tournamentWS.Rows.count, rightEntryNumCol - 1).End(xlUp).row
    For i = 1 To lastRow Step 2
        With tournamentWS
            With .Range(.Cells(i, rightEntryNumCol), .Cells(i + 1, rightEntryNumCol))
                .Merge
                .Interior.Color = vbYellow
                .Borders.LineStyle = xlContinuous
            End With
            Set ref = .Cells(i, rightEntryNumCol)
            .Cells(i, rightEntryNumCol - 5) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",2,FALSE),"""")"
            .Cells(i + 1, rightEntryNumCol - 5) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",3,FALSE),"""")"
            .Cells(i, rightEntryNumCol - 3) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",4,FALSE),"""")"
            .Cells(i + 1, rightEntryNumCol - 3) = "=IFERROR(VLOOKUP(" & ref.Address & "," & entryPlayersRange.Address(External:=True) & ",5,FALSE),"""")"
        End With
    Next i
    
    isEditModeRange.Value = "途中"
        
End Sub
Sub finishEditMode()
    setUp
    
    Dim msgRes As VbMsgBoxResult
    
    If (isEditModeRange.Value <> "途中") Then
        MsgBox "現在、編集モードではありません。", _
            Title:="編集モード"
        Exit Sub
    End If

    msgRes = MsgBox( _
        Prompt:="トーナメント編集モードを終了します。よろしいですか？", _
        Buttons:=vbOKCancel, _
        Title:="編集モード" _
    )

    If (msgRes = vbCancel) Then
        Exit Sub
    End If
    
    Call updatePlayerListFromTournament(G_numLeftCol, G_numRightCol + 2)
    
    tournamentWS.Columns(G_numLeftCol).Delete
    tournamentWS.Columns(G_numRightCol + 1).Delete
    
    isEditModeRange.Value = "済"
    
    Call insertPlayerInformation
End Sub
' ワークシートが存在するかチェックする
' 参考：https://qiita.com/Zitan/items/1b671510d3da5557ba1a
Function flgExsistSheet(ByVal WorkSheetName As String) As Boolean
Dim sht As Worksheet
  For Each sht In ActiveWorkbook.Worksheets
    If sht.Name = WorkSheetName Then
        flgExsistSheet = True
        Exit Function
    End If
  Next sht
flgExsistSheet = False
End Function

Function makeEntryPlayersSheet()
    setUp
    
    Dim i As Integer

    entryPlayersWS.Cells(1, 1) = "エントリー番号"
    entryPlayersWS.Cells(1, 2) = "後衛名前"
    entryPlayersWS.Cells(1, 3) = "前衛名前"
    entryPlayersWS.Cells(1, 4) = "後衛所属"
    entryPlayersWS.Cells(1, 5) = "前衛所属"
    
    For i = 2 To teamsRange.Value + 1
        entryPlayersWS.Cells(i, 1) = i - 1
    Next i
End Function

Function updatePlayerListFromTournament(leftEntryNumCol As Integer, rightEntryNumCol As Integer)
    Dim row As Integer
    Dim entryNum As Integer
    Dim tournamentNum As Integer
    Dim lastRow As Integer
    Dim p As player
    
    ' 左側
    lastRow = tournamentWS.Cells(tournamentWS.Rows.count, leftEntryNumCol).End(xlUp).row
    
    For row = 1 To lastRow
        If (tournamentWS.Cells(row, leftEntryNumCol) <> "" And tournamentWS.Cells(row, leftEntryNumCol + 1) <> "") Then
            entryNum = tournamentWS.Cells(row, leftEntryNumCol)
            tournamentNum = tournamentWS.Cells(row, leftEntryNumCol + 1)
            Set p = findEntryPlayer(entryNum)
            Call insertPlayer(tournamentNum, p)
        End If
    Next row
    
    ' 右側
    lastRow = tournamentWS.Cells(tournamentWS.Rows.count, rightEntryNumCol).End(xlUp).row
    
    For row = 1 To lastRow
        If (tournamentWS.Cells(row, rightEntryNumCol) <> "" And tournamentWS.Cells(row, rightEntryNumCol - 1) <> "") Then
            entryNum = tournamentWS.Cells(row, rightEntryNumCol)
            tournamentNum = tournamentWS.Cells(row, rightEntryNumCol - 1)
            Set p = findEntryPlayer(entryNum)
            Call insertPlayer(tournamentNum, p)
        End If
    Next row
End Function

Function findEntryPlayer(key As Integer) As player
    Dim row As Integer
    Dim lastRow As Integer
    Dim p As player
    
    Set p = New player
    
    lastRow = entryPlayersWS.Cells(entryPlayersWS.Rows.count, 1).End(xlUp).row
    
    For row = 1 To lastRow
        If (entryPlayersWS.Cells(row, 1) = key) Then
            p.AName = entryPlayersWS.Cells(row, 2)
            p.BName = entryPlayersWS.Cells(row, 3)
            p.ATeam = entryPlayersWS.Cells(row, 4)
            p.BTeam = entryPlayersWS.Cells(row, 5)
            
            Set findEntryPlayer = p
            Exit Function
        End If
    Next row
End Function
