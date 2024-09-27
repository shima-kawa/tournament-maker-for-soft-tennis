Attribute VB_Name = "Constants"
Public mainWS As Worksheet
Public tournamentWS As Worksheet
Public testWorksheet As Worksheet
Public baseMatchesWS As Worksheet
Public matchesWS As Worksheet
Public playerListWS As Worksheet
Public judgePaperWS As Worksheet

Public teamsRange As Range
Public plgStartNoRange As Range
Public maxNumPerPageRange As Range
Public categoryRange As Range

' 試合シート用
Public Const G_idCol As Integer = 1
Public Const G_baseMatchIdCol As Integer = 2
Public Const G_roundCol As Integer = 3
Public Const G_fromCol As Integer = 4
Public Const G_toCol As Integer = 5
Public Const G_statusCol As Integer = 6
Public Const G_matchGamesCol As Integer = 7
Public Const G_leftCol As Integer = 8
Public Const G_rightCol As Integer = 9
Public Const G_winnerCol As Integer = 10
Public Const G_scoreLeftCol As Integer = 11
Public Const G_scoreRightCol As Integer = 12
Public Const G_addressLeftRowCol As Integer = 13
Public Const G_addressLeftColCol As Integer = 14
Public Const G_addressRightRowCol As Integer = 15
Public Const G_addressRightColCol As Integer = 16
Public Const G_nextMatchRowCol As Integer = 17
Public Const G_nextMatchColCol As Integer = 18
Public Const G_LRCol = 19

Public Const MATCH_NOT_ALLOWED As Integer = 0
Public Const MATCH_ALLOWED_NOPRINT As Integer = 1
Public Const MATCH_ALLOWED_PRINTED As Integer = 2
Public Const MATCH_FINISHED As Integer = 3

Public Const LEFT As Integer = 0
Public Const RIGHT As Integer = 1

' トーナメントシート用
Public Const G_numLeftCol As Integer = 1
Public Const G_nameLeftCol As Integer = 2
Public Const G_teamLeftCol As Integer = 4
Public Const G_numRightCol As Integer = 24
Public Const G_nameRightCol As Integer = 20
Public Const G_teamRightCol As Integer = 22
Public Const G_startTournamentArea As Integer = 6
Public Const G_endTournamentArea As Integer = 19


' 選手一覧シート用
Public Const plgNoCol As Integer = 1
Public Const playerANameCol As Integer = 2
Public Const playerATeamCol As Integer = 3
Public Const playerBNameCol As Integer = 4
Public Const playerBTeamCol As Integer = 5
Sub setUp()
    
    Set mainWS = ThisWorkbook.Worksheets("メイン")
    Set testWorksheet = ThisWorkbook.Worksheets("ベース")
    Set matchesWS = ThisWorkbook.Worksheets("試合")
    Set baseMatchesWS = ThisWorkbook.Worksheets("ベース")
    Set tournamentWS = ThisWorkbook.Worksheets("トーナメント")
    Set playerListWS = ThisWorkbook.Worksheets("選手一覧")
    Set judgePaperWS = ThisWorkbook.Worksheets("個人ジャッペ")
    
    Set teamsRange = mainWS.Range("B1")
    Set plgStartNoRange = mainWS.Range("B4")
    Set maxNumPerPageRange = mainWS.Range("B2")
    Set categoryRange = mainWS.Range("B3")
    
    If ((G_endTournamentArea - G_startTournamentArea) Mod 2 = 0) Then
        MsgBox エラー｡トーナメント範囲が無効です｡範囲は偶数個のセルが必要です｡
        Exit Sub
    End If
    
    Call localSetUp
    
    
End Sub
