VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCorrectScore 
   Caption         =   "スコア訂正"
   ClientHeight    =   6045
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   10230
   OleObjectBlob   =   "frmCorrectScore.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmCorrectScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()
    Unload Me
End Sub


Private Function makeList(matches() As match, results() As Result)
    
    Dim max_num As Integer
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
    max_num = UBound(matches)
    
    ' 既にある一覧を削除
    Call clearForm

    ' 一覧を作成
    For i = 1 To max_num
        Set p1 = findPlayer(matches(i).leftNum)
        Set p2 = findPlayer(matches(i).rightNum)
    
        Set newLblLeftPlayerNum = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblLeftPlayerNum" & i)
        Set newLblLeftNameA = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblLeftNameA" & i)
        Set newLblLeftNameB = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblLeftNameB" & i)
        Set newLblRightPlayerNum = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblRightPlayerNum" & i)
        Set newLblRightNameA = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblRightNameA" & i)
        Set newLblRightNameB = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lblRightNameB" & i)
        
        Set newTxtLeftScore = frmCorrectScore.frmMatchList.Add("Forms.TextBox.1", "txtLeftScore" & i)
        Set newTxtRightScore = frmCorrectScore.frmMatchList.Add("Forms.TextBox.1", "txtRightScore" & i)
        Set newLblHyphen = frmCorrectScore.frmMatchList.Add("Forms.Label.1", "lbl" & i)
        
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

    'フレームのスクロール量を調節する
    frmCorrectScore.frmMatchList.ScrollHeight = i * 25

    
End Function

Private Sub btnFind_Click()
    Dim matches() As match
    Dim results() As Result
    Dim key As Integer
    Dim i As Integer
    Dim countMatches As Integer
    
    key = Me.txtPlayerID.Value
    
    matches = findAllMatchesWithStatus(key, MATCH_FINISHED)
    countMatches = UBound(matches)
    ReDim results(countMatches)
    For i = 1 To countMatches
        Set results(i) = findResult(matches(i).matchID)
    Next i
    If (UBound(matches) = 0) Then
        MsgBox "入力された試合はありません。"
        Exit Sub
    End If
    Call makeList(matches, results)

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
