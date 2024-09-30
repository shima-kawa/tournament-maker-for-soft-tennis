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
    

    ' �G���[����-------------------------------------------------------------------------------
    If (isTournamentGeneratedRange.Value <> "��") Then
        MsgBox "�g�[�i�����g���쐬����Ă��܂���B�g�[�i�����g���쐬���Ă��������B", _
            Title:="�ҏW���[�h"
        Exit Sub
    End If
    
    If (isEditModeRange.Value = "��") Then
        ' temp
        MsgBox "�g�����͌���ς݂ł��B���Z�b�g����ۂ́A�u�g�[�i�����g�쐬�v�{�^���������Ă��������B", _
            Title:="�ҏW���[�h"
        Exit Sub
    End If
    If (isEditModeRange.Value = "�r��") Then
        MsgBox "���݁A�ҏW���[�h�ł��B", _
            Title:="�ҏW���[�h"
        Exit Sub
    End If
    
    msgRes = MsgBox( _
        Prompt:="�g�[�i�����g�ҏW���[�h�֓���܂��B��낵���ł����H", _
        Buttons:=vbOKCancel, _
        Title:="�ҏW���[�h" _
    )
    
    If (msgRes = vbCancel) Then
        Exit Sub
    End If
    
    ' �G���g���[����V�[�g�̊m�F----------------------------------------------------------------
    If (flgExsistSheet("�G���g���[����") = False) Then
        Worksheets.Add after:=Sheets(Worksheets.count)
        ActiveSheet.Name = "�G���g���[����"
        Set entryPlayersWS = ThisWorkbook.Worksheets("�G���g���[����")
        Call makeEntryPlayersSheet
    End If
    
    Set entryPlayersRange = entryPlayersWS.Range("A:E") ' TODO: �ʂ̏ꏊ����ύX������A���I�ɐݒ�ł���悤�ɂ���

    
    ' �g�[�i�����g�V�[�g�̕ύX------------------------------------------------------------------
    leftEntryNumCol = G_numLeftCol
    rightEntryNumCol = G_numRightCol + 2
    tournamentWS.Columns(G_numLeftCol).Insert ' �g�[�i�����g�̍����B�E���͂��łɋ󂢂Ă���̂ŁA����𗘗p
    
    tournamentWS.Columns(leftEntryNumCol).ColumnWidth = 4
    tournamentWS.Columns(rightEntryNumCol).ColumnWidth = 4
    
    ' �Z���ւ̊֐��̑}��------------------------------------------------------------------------
    ' ����
    lastRow = tournamentWS.Cells(matchesWS.Rows.count, leftEntryNumCol + 1).End(xlUp).row
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

    ' �E��
    lastRow = tournamentWS.Cells(matchesWS.Rows.count, rightEntryNumCol - 1).End(xlUp).row
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
    
    isEditModeRange.Value = "�r��"
        
End Sub
Sub finishEditMode()
    setUp
    
    Dim msgRes As VbMsgBoxResult
    
    If (isEditModeRange.Value <> "�r��") Then
        MsgBox "���݁A�ҏW���[�h�ł͂���܂���B", _
            Title:="�ҏW���[�h"
        Exit Sub
    End If

    msgRes = MsgBox( _
        Prompt:="�g�[�i�����g�ҏW���[�h���I�����܂��B��낵���ł����H", _
        Buttons:=vbOKCancel, _
        Title:="�ҏW���[�h" _
    )

    If (msgRes = vbCancel) Then
        Exit Sub
    End If
    
    tournamentWS.Columns(G_numLeftCol).Delete
    tournamentWS.Columns(G_numRightCol + 1).Delete
    
    isEditModeRange.Value = "��"
End Sub
' ���[�N�V�[�g�����݂��邩�`�F�b�N����
' �Q�l�Fhttps://qiita.com/Zitan/items/1b671510d3da5557ba1a
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

    entryPlayersWS.Cells(1, 1) = "�G���g���[�ԍ�"
    entryPlayersWS.Cells(1, 2) = "��q���O"
    entryPlayersWS.Cells(1, 3) = "�O�q���O"
    entryPlayersWS.Cells(1, 4) = "��q����"
    entryPlayersWS.Cells(1, 5) = "�O�q����"
    
    For i = 2 To teamsRange.Value + 1
        entryPlayersWS.Cells(i, 1) = i - 1
    Next i
End Function
