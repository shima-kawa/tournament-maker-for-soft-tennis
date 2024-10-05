Attribute VB_Name = "JudgePaperMaker"
Option Explicit

Sub printJudgePaper()
    setUp
    
    ' ---------------------------------------------------
    ' 個人ジャッペシート用
    Dim categoryRange As Range
    Dim cordRange As Range
    Dim roundRange As Range
    Dim LeftNumberRange As Range
    Dim RightNumberRange As Range
    Dim LeftTeamRange As Range
    Dim RightTeamRange As Range
    Dim LeftANameRange As Range
    Dim LeftBNameRange As Range
    Dim RightANameRange As Range
    Dim RightBNameRange As Range
    With judgePaperWS
        Set categoryRange = .Range("AU2")
        Set cordRange = .Range("AU3")
        Set roundRange = .Range("AU4")
        Set LeftNumberRange = .Range("AU5")
        Set RightNumberRange = .Range("AU6")
        Set LeftTeamRange = .Range("AU7")
        Set RightTeamRange = .Range("AU8")
        Set LeftANameRange = .Range("AU9")
        Set LeftBNameRange = .Range("AU10")
        Set RightANameRange = .Range("AU11")
        Set RightBNameRange = .Range("AU12")
    End With
    

    
    ' ---------------------------------------------------
    Dim printNum As Integer
    Dim i As Integer
    Dim lastRow As Integer
    Dim p As player
    
    printNum = 0
    
    ' 印刷数の確認
    printNum = WorksheetFunction.CountIf(matchesWS.Cells(1, G_statusCol).EntireColumn, MATCH_ALLOWED_NOPRINT)
    If (printNum > 0) Then
        MsgBox "プリント可能数：" & printNum & "枚" & vbLf & "採点表をセットしてください。"
    Else
        MsgBox "印刷可能な試合がありません。"
        End
    End If
    
    ' 印刷
    
    lastRow = matchesWS.Cells(matchesWS.Rows.Count, 1).End(xlUp).row
    For i = 1 To lastRow
        If (matchesWS.Cells(i, G_statusCol) = MATCH_ALLOWED_NOPRINT) Then
            clearJudgePaper
            With matchesWS
                'categoryRange.Value =
                'cordRange.Value =
                roundRange.Value = .Cells(i, G_roundCol)
                LeftNumberRange.Value = .Cells(i, G_leftCol)
                RightNumberRange.Value = .Cells(i, G_rightCol)
                Set p = findPlayer(.Cells(i, G_leftCol))
                LeftANameRange.Value = p.AName
                LeftBNameRange.Value = p.BName
                If (p.ATeam = p.BTeam) Then
                    LeftTeamRange.Value = p.ATeam
                Else
                    LeftTeamRange.Value = p.ATeam & vbLf & p.BTeam
                End If
                Set p = findPlayer(.Cells(i, G_rightCol))
                RightANameRange.Value = p.AName
                RightBNameRange.Value = p.BName
                If (p.ATeam = p.BTeam) Then
                    RightTeamRange.Value = p.ATeam
                Else
                    RightTeamRange.Value = p.ATeam & vbLf & p.BTeam
                End If
            End With
            'tempDir = tempDir & i & ".pdf"
            'judgePaperWS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=tempDir
            'judgePaperWS.PrintOut
            matchesWS.Cells(i, G_statusCol) = MATCH_ALLOWED_PRINTED
        End If
    Next i

End Sub

Function clearJudgePaper()
    
    Dim judgePaperWS As Worksheet
    Set judgePaperWS = ThisWorkbook.Sheets("個人ジャッペ")
    ' 個人ジャッペシート用
    Dim categoryRange As Range
    Dim cordRange As Range
    Dim roundRange As Range
    Dim LeftNumberRange As Range
    Dim RightNumberRange As Range
    Dim LeftTeamRange As Range
    Dim RightTeamRange As Range
    Dim LeftANameRange As Range
    Dim LeftBNameRange As Range
    Dim RightANameRange As Range
    Dim RightBNameRange As Range
    With judgePaperWS
        Set categoryRange = .Range("AU2")
        Set cordRange = .Range("AU3")
        Set roundRange = .Range("AU4")
        Set LeftNumberRange = .Range("AU5")
        Set RightNumberRange = .Range("AU6")
        Set LeftTeamRange = .Range("AU7")
        Set RightTeamRange = .Range("AU8")
        Set LeftANameRange = .Range("AU9")
        Set LeftBNameRange = .Range("AU10")
        Set RightANameRange = .Range("AU11")
        Set RightBNameRange = .Range("AU12")
    End With

    categoryRange.ClearContents
    cordRange.ClearContents
    roundRange.ClearContents
    LeftNumberRange.ClearContents
    RightNumberRange.ClearContents
    LeftTeamRange.ClearContents
    RightTeamRange.ClearContents
    LeftANameRange.ClearContents
    LeftBNameRange.ClearContents
    RightANameRange.ClearContents
    RightBNameRange.ClearContents
End Function

