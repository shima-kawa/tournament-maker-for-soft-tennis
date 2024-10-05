VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResultInput 
   Caption         =   "結果登録"
   ClientHeight    =   7575
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   13890
   OleObjectBlob   =   "frmResultInput.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmResultInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private leftNums(8) As Integer
Private matches(8) As match
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnInput_Click()
    Dim i As Integer
    Dim validCount As Integer
    Dim r As Result
    
    ' ON : Object Name
    Dim ONleftPlayerNum As String
    Dim ONleftCount As String
    Dim ONrightCount As String

    validCount = 0
    
    If (checkInputResults = False) Then
        MsgBox "登録に失敗！"
        Exit Sub
    End If
    
    For i = 1 To 8
        If (matches(i) Is Nothing) Then
        Else
            validCount = validCount + 1
        End If
    Next i
    MsgBox validCount & "件登録します。"
    
    For i = 1 To 8
        If (matches(i) Is Nothing) Then
        Else
            ONleftPlayerNum = "txtLeftPlayerNum" & i
            ONleftCount = "txtLeftScore" & i
            ONrightCount = "txtRightScore" & i
            Set r = New Result
            r.matchID = matches(i).matchID
            r.leftScore = Me.Controls(ONleftCount)
            r.rightScore = Me.Controls(ONrightCount)
            If (r.leftScore > r.rightScore) Then
                r.winner = matches(i).leftNum
            Else
                r.winner = matches(i).rightNum
            End If
            Call registerResult(r)
        End If
    Next i
    
    Unload Me
End Sub


Private Sub lblCategory_Click()

End Sub

Private Sub txtLeftPlayerNum1_AfterUpdate()
    If (Me.txtLeftPlayerNum1.Value <> "") Then
        leftNums(1) = Me.txtLeftPlayerNum1.Value
    End If
End Sub
Private Sub txtLeftPlayerNum2_AfterUpdate()
    If (Me.txtLeftPlayerNum2.Value <> "") Then
        leftNums(2) = Me.txtLeftPlayerNum2.Value
    End If
End Sub
Private Sub txtLeftPlayerNum3_AfterUpdate()
    If (Me.txtLeftPlayerNum3.Value <> "") Then
        leftNums(3) = Me.txtLeftPlayerNum3.Value
    End If
End Sub
Private Sub txtLeftPlayerNum4_AfterUpdate()
    If (Me.txtLeftPlayerNum4.Value <> "") Then
        leftNums(4) = Me.txtLeftPlayerNum4.Value
    End If
End Sub
Private Sub txtLeftPlayerNum5_AfterUpdate()
    If (Me.txtLeftPlayerNum5.Value <> "") Then
        leftNums(5) = Me.txtLeftPlayerNum5.Value
    End If
End Sub
Private Sub txtLeftPlayerNum6_AfterUpdate()
    If (Me.txtLeftPlayerNum6.Value <> "") Then
        leftNums(6) = Me.txtLeftPlayerNum6.Value
    End If
End Sub
Private Sub txtLeftPlayerNum7_AfterUpdate()
    If (Me.txtLeftPlayerNum7.Value <> "") Then
        leftNums(7) = Me.txtLeftPlayerNum7.Value
    End If
End Sub
Private Sub txtLeftPlayerNum8_AfterUpdate()
    If (Me.txtLeftPlayerNum8.Value <> "") Then
        leftNums(8) = Me.txtLeftPlayerNum8.Value
    End If
End Sub

Private Sub txtLeftPlayerNum1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 1, Cancel
End Sub

Private Sub txtLeftPlayerNum2_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 2, Cancel
End Sub
Private Sub txtLeftPlayerNum3_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 3, Cancel
End Sub
Private Sub txtLeftPlayerNum4_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 4, Cancel
End Sub
Private Sub txtLeftPlayerNum5_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 5, Cancel
End Sub
Private Sub txtLeftPlayerNum6_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 6, Cancel
End Sub
Private Sub txtLeftPlayerNum7_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 7, Cancel
End Sub
Private Sub txtLeftPlayerNum8_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    checkInputLeftNum 8, Cancel
End Sub

Private Sub txtLeftPlayerNum2_Change()

End Sub


Private Sub txtLeftScore1_Change()

End Sub

Private Sub UserForm_Initialize()
    setUp
    Me.lblCategory = categoryRange.Value
End Sub

Private Function checkInputLeftNum(setID As Integer, ByRef Cancel As MSForms.ReturnBoolean)
    Dim m As match
    Dim p As player
    Dim i As Integer
    
    ' ON -> Object Name
    Dim ONleftNum As String
    Dim ONleftNameA As String
    Dim ONleftNameB As String
    Dim ONrightNum As String
    Dim ONrightNameA As String
    Dim ONrightNameB As String
    
    ONleftNum = "txtLeftPlayerNum" & setID
    ONleftNameA = "lblLeftNameA" & setID
    ONleftNameB = "lblLeftNameB" & setID
    ONrightNum = "lblRightPlayerNum" & setID
    ONrightNameA = "lblRightNameA" & setID
    ONrightNameB = "lblRightNameB" & setID
    
    ' 空白チェック----------------------------------------------
    If (Me.Controls(ONleftNum).Value = "") Then
        leftNums(setID) = 0
        clear1Row (setID)
        Exit Function
    End If
    
    ' 重複チェック----------------------------------------------
    For i = 1 To 8
        If (leftNums(i) = Me.Controls(ONleftNum).Value) Then
            MsgBox "入力された試合が重複しています。"
            ' エラー範囲再選択
            Me.Controls(ONleftNum).SetFocus
            Me.Controls(ONleftNum).SelStart = 0
            Me.Controls(ONleftNum).SelLength = Len(Me.Controls(ONleftNum).Value)
            Cancel = True
            Exit Function
        End If
    Next i
    
    ' 試合チェック----------------------------------------------
    Set m = findMatch(Me.Controls(ONleftNum).Value)
    
    If (m Is Nothing) Then
        MsgBox "エラー。試合がありません。"
        Me.Controls(ONleftNameA) = ""
        Me.Controls(ONleftNameB) = ""
        Me.Controls(ONrightNameA) = ""
        Me.Controls(ONrightNameB) = ""
        ' エラー範囲再選択
        Me.Controls(ONleftNum).SetFocus
        Me.Controls(ONleftNum).SelStart = 0
        Me.Controls(ONleftNum).SelLength = Len(Me.Controls(ONleftNum).Value)
        Cancel = True
    Else
        Set matches(setID) = m
        Me.Controls(ONrightNum) = m.rightNum
        Set p = findPlayer(m.leftNum)
        Me.Controls(ONleftNameA) = p.AName
        Me.Controls(ONleftNameB) = p.BName
        Set p = findPlayer(m.rightNum)
        Me.Controls(ONrightNameA) = p.AName
        Me.Controls(ONrightNameB) = p.BName
    End If
End Function

Function checkInputResults() As Boolean
    Dim i As Integer
    Dim leftCount As Integer
    Dim rightCount As Integer
    Dim errorMessage As String
    Dim errorCount As Integer
    Dim validCount As Integer
    
    errorMessage = ""
    errorCount = 0
    validCount = 0
    
    ' ON : Object Name
    Dim ONleftPlayerNum As String
    Dim ONleftCount As String
    Dim ONrightCount As String
    
    For i = 1 To 8
        ONleftPlayerNum = "txtLeftPlayerNum" & i
        ONleftCount = "txtLeftScore" & i
        ONrightCount = "txtRightScore" & i
        
        ' 空白チェック------------------------------------------------
        If (Me.Controls(ONleftPlayerNum) = "" And Me.Controls(ONleftCount) = "" And Me.Controls(ONrightCount) = "") Then
            ' 1行空白
            GoTo CONTINUE
        End If
        If (Me.Controls(ONleftPlayerNum) = "") Then
            errorMessage = errorMessage & "id = " & i & "選手番号を入力してください。" & vbLf
            errorCount = errorCount + 1
            GoTo CONTINUE
        End If
        If (Me.Controls(ONleftCount) = "" Or Me.Controls(ONrightCount) = "") Then
            errorMessage = errorMessage & "id = " & i & "スコアを入力してください。" & vbLf
                errorCount = errorCount + 1
            GoTo CONTINUE
        End If
        
        ' リザルトチェック--------------------------------------------
        leftCount = Me.Controls(ONleftCount)
        rightCount = Me.Controls(ONrightCount)
        If (leftCount + rightCount > matches(i).matchGames) Then
            errorMessage = errorMessage & "id = " & i & "。ゲーム数が不正です。" & vbLf & "多い。この試合は" & matches(i).matchGames & "ゲームマッチです。" & vbLf
            errorCount = errorCount + 1
        ElseIf (leftCount <> getRequiredGames(matches(i).matchGames) And rightCount <> getRequiredGames(matches(i).matchGames)) Then
            errorMessage = errorMessage & "id = " & i & "。ゲーム数が不正です。" & vbLf & "必要なゲームを取っていない、または、多い。この試合は" & matches(i).matchGames & "ゲームマッチです。" & vbLf
            errorCount = errorCount + 1
        Else
            validCount = validCount + 1
        End If
        
CONTINUE:
    Next i
    
    If (errorCount > 0) Then
        MsgBox "エラー：" & errorCount & vbLf & errorMessage
        checkInputResults = False
    Else
        checkInputResults = True
    End If
    

End Function

Function clear1Row(setID As Integer)
    Dim ONleftNum As String
    Dim ONleftNameA As String
    Dim ONleftNameB As String
    Dim ONrightNum As String
    Dim ONrightNameA As String
    Dim ONrightNameB As String
    
    ONleftNum = "txtLeftPlayerNum" & setID
    ONleftNameA = "lblLeftNameA" & setID
    ONleftNameB = "lblLeftNameB" & setID
    ONrightNum = "lblRightPlayerNum" & setID
    ONrightNameA = "lblRightNameA" & setID
    ONrightNameB = "lblRightNameB" & setID

    Me.Controls(ONleftNameA) = ""
    Me.Controls(ONleftNameB) = ""
    Me.Controls(ONrightNum) = ""
    Me.Controls(ONrightNameA) = ""
    Me.Controls(ONrightNameB) = ""
End Function
