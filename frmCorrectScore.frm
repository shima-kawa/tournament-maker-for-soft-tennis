VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorrectScore 
   Caption         =   "�X�R�A����"
   ClientHeight    =   4830
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8160
   OleObjectBlob   =   "frmCorrectScore.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmCorrectScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ON: Object Name
Private Const ONleftPlayerNum As String = "lblLeftPlayerNum" ' & i
Private Const ONleftPlayerNameA As String = "lblLeftNameA" ' & i
Private Const ONleftPlayerNameB As String = "lblLeftNameB" ' & i
Private Const ONRightPlayerNum As String = "lblRightPlayerNum" ' & i
Private Const ONRightPlayerNameA As String = "lblRightNameA" ' & i
Private Const ONRightPlayerNameB As String = "lblRightNameB" ' & i
Private Const ONLeftScore As String = "txtLeftScore" ' & i
Private Const ONRightScore As String = "txtRightScore" ' & i

Private countMatches As Integer
Private matches() As Match
Private results() As Result


Private Sub btnCorrect_Click()
    Dim i As Integer
    Dim winner As Integer
    Dim isModified() As Boolean
    Dim isChangedWinner() As Boolean
    Dim dangerFlg As Boolean
    Dim r As Result
    
    ' �������̊m�F
    If (checkInputResults = False) Then
        Exit Sub
    End If
    
    ' �ύX�_�̊m�F
    isModified = checkModified()
    
    ' ���s���ύX���ꂽ���ǂ����̊m�F
    ReDim isChangedWinner(countMatches) As Boolean
    For i = 1 To countMatches
        If (Me.Controls(ONLeftScore & i).Value > Me.Controls(ONRightScore & i).Value) Then
            winner = matches(i).leftNum
        Else
            winner = matches(i).rightNum
        End If
        
        ' ���҂̕ύX�͎��ΐ킪�I�����Ă��Ȃ��ꍇ�̂ݔF�߂�
        If (winner <> results(i).winner) Then
            isChangedWinner(i) = True
            If (getNextMatchStatus(matches(i).matchID) = MATCH_FINISHED) Then
                dangerFlg = True
            End If
        End If
    Next i
    
    
    If (dangerFlg = True) Then
        MsgBox "���̎������I���ς݂̂��߁A���̕ύX�͂ł��܂���"
        Exit Sub
    End If
    
        
    ' �X�R�A�̕ύX����
    For i = 1 To countMatches
        If (isModified(i)) Then
            Set r = New Result
            r.matchID = matches(i).matchID
            r.leftScore = Me.Controls(ONLeftScore & i).Value
            r.rightScore = Me.Controls(ONRightScore & i).Value
            If (r.leftScore > r.rightScore) Then
                r.winner = matches(i).leftNum
            Else
                r.winner = matches(i).rightNum
            End If
            Call registerResult(r)
        End If
    Next i
    
    MsgBox "�������܂���"
    Unload Me
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub


Private Function makeList()
    
    Dim i As Integer
    Dim newLblLeftPlayerNum As Object
    Dim newLblLeftNameA As Object
    Dim newLblLeftNameB As Object
    Dim newLblRightPlayerNum As Object
    Dim newLblRightNameA As Object
    Dim newLblRightNameB As Object
    Dim newTxtLeftScore As Object
    Dim newTxtRightScore As Object
    Dim newLblHyphen As Object
    Dim p1 As player
    Dim p2 As player
    
    ' ���ɂ���ꗗ���폜
    Call clearForm

    ' �ꗗ���쐬
    For i = 1 To countMatches
        Set p1 = findPlayer(matches(i).leftNum)
        Set p2 = findPlayer(matches(i).rightNum)
    
        Set newLblLeftPlayerNum = frmCorrectScore.frmMatchList.Add("Forms.Label.1", ONleftPlayerNum & i)
        Set newLblLeftNameA = frmCorrectScore.frmMatchList.Add("Forms.Label.1", ONleftPlayerNameA & i)
        Set newLblLeftNameB = frmCorrectScore.frmMatchList.Add("Forms.Label.1", ONleftPlayerNameB & i)
        Set newLblRightPlayerNum = frmCorrectScore.frmMatchList.Add("Forms.Label.1", ONRightPlayerNum & i)
        Set newLblRightNameA = frmCorrectScore.frmMatchList.Add("Forms.Label.1", ONRightPlayerNameA & i)
        Set newLblRightNameB = frmCorrectScore.frmMatchList.Add("Forms.Label.1", ONRightPlayerNameB & i)
        
        Set newTxtLeftScore = frmCorrectScore.frmMatchList.Add("Forms.TextBox.1", ONLeftScore & i)
        Set newTxtRightScore = frmCorrectScore.frmMatchList.Add("Forms.TextBox.1", ONRightScore & i)
        Set newLblHyphen = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblHyphen" & i)
        
        With newLblLeftPlayerNum
            .Width = 20
            .Height = 20
            .LEFT = 10
            .Top = (i - 1) * 25 + 10
            .Caption = p1.programNum
            .BorderStyle = 1
        End With
        With newLblLeftNameA
            .Width = 75
            .Height = 20
            .LEFT = 35
            .Top = (i - 1) * 25 + 10
            .Caption = p1.AName
            .BorderStyle = 1
        End With
        With newLblLeftNameB
            .Width = 75
            .Height = 20
            .LEFT = 115
            .Top = (i - 1) * 25 + 10
            .Caption = p1.BName
            .BorderStyle = 1
        End With
        With newTxtLeftScore
            .Width = 20
            .Height = 20
            .LEFT = 195
            .Top = (i - 1) * 25 + 10
            .Value = results(i).leftScore
            .BorderStyle = 1
        End With
        With newLblHyphen
            .Width = 10
            .Height = 20
            .LEFT = 220
            .Top = (i - 1) * 25 + 10
            .Caption = "-"
            .BorderStyle = 1
        End With
        With newTxtRightScore
            .Width = 20
            .Height = 20
            .LEFT = 235
            .Top = (i - 1) * 25 + 10
            .Value = results(i).rightScore
            .BorderStyle = 1
        End With
        With newLblRightNameA
            .Width = 75
            .Height = 20
            .LEFT = 260
            .Top = (i - 1) * 25 + 10
            .Caption = p2.AName
            .BorderStyle = 1
        End With
        With newLblRightNameB
            .Width = 75
            .Height = 20
            .LEFT = 340
            .Top = (i - 1) * 25 + 10
            .Caption = p2.BName
            .BorderStyle = 1
        End With
        With newLblRightPlayerNum
            .Width = 20
            .Height = 20
            .LEFT = 420
            .Top = (i - 1) * 25 + 10
            .Caption = p2.programNum
            .BorderStyle = 1
        End With
    Next

    '�t���[���̃X�N���[���ʂ𒲐߂���
    frmCorrectScore.frmMatchList.ScrollHeight = i * 25

    
End Function

Private Sub btnFind_Click()
    Dim key As Integer
    Dim i As Integer
    
    key = Me.txtPlayerID.Value
    
    matches = findAllMatchesWithStatus(key, MATCH_FINISHED)
    countMatches = UBound(matches)
    ReDim results(countMatches)
    For i = 1 To countMatches
        Set results(i) = findResult(matches(i).matchID)
    Next i
    If (UBound(matches) = 0) Then
        MsgBox "���͂��ꂽ�����͂���܂���B"
        Exit Sub
    End If
    Call makeList

End Sub
Private Function clearForm()
    Dim ctrl As Control
    Dim i As Integer
        
    For i = Me.Controls.Count - 1 To 0 Step -1
        Set ctrl = Me.Controls(i)
        If (ctrl.parent Is Me.frmMatchList) Then
            Me.Controls.Remove ctrl.Name
        End If
    Next i
End Function


Private Sub UserForm_Initialize()
    setUp
End Sub

Private Function checkModified() As Boolean()
    Dim i As Integer
    
    Dim isModified() As Boolean
    ReDim isModified(1 To countMatches) As Boolean
    
    ' ������
    For i = 1 To countMatches
        isModified(i) = False
    Next i
    
    ' �T�[�`
    For i = 1 To countMatches
        If (Me.Controls(ONLeftScore & i).Value <> results(i).leftScore) Then
            isModified(i) = True
        End If
        If (Me.Controls(ONRightScore & i).Value <> results(i).rightScore) Then
            isModified(i) = True
        End If
        
    Next i
    
    checkModified = isModified
End Function
Private Function checkInputResults() As Boolean
    Dim i As Integer
    Dim leftCount As Integer
    Dim rightCount As Integer
    Dim errorMessage As String
    Dim errorCount As Integer
    Dim validCount As Integer
    
    errorMessage = ""
    errorCount = 0
    validCount = 0
        
    For i = 1 To countMatches
        
        ' �󔒃`�F�b�N------------------------------------------------
        If (Me.Controls(ONLeftScore & i) = "" Or Me.Controls(ONRightScore & i) = "") Then
            errorMessage = errorMessage & "id = " & i & "�X�R�A����͂��Ă��������B" & vbLf
                errorCount = errorCount + 1
            GoTo CONTINUE
        End If
        
        ' ���U���g�`�F�b�N--------------------------------------------
        leftCount = Me.Controls(ONLeftScore & i)
        rightCount = Me.Controls(ONRightScore & i)
        If (leftCount + rightCount > matches(i).matchGames) Then
            errorMessage = errorMessage & "id = " & i & "�B�Q�[�������s���ł��B" & vbLf & "�����B���̎�����" & matches(i).matchGames & "�Q�[���}�b�`�ł��B" & vbLf
            errorCount = errorCount + 1
        ElseIf (leftCount <> getRequiredGames(matches(i).matchGames) And rightCount <> getRequiredGames(matches(i).matchGames)) Then
            errorMessage = errorMessage & "id = " & i & "�B�Q�[�������s���ł��B" & vbLf & "�K�v�ȃQ�[��������Ă��Ȃ��A�܂��́A�����B���̎�����" & matches(i).matchGames & "�Q�[���}�b�`�ł��B" & vbLf
            errorCount = errorCount + 1
        Else
            validCount = validCount + 1
        End If
        
CONTINUE:
    Next i
    
    If (errorCount > 0) Then
        MsgBox "�G���[�F" & errorCount & vbLf & errorMessage
        checkInputResults = False
    Else
        checkInputResults = True
    End If
    

End Function

