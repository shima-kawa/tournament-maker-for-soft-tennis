Attribute VB_Name = "ForTest"
Option Explicit
' �g�[�i�����g�̌`���m�F����
' �w�肳�ꂽ�͈͓��Ńg�[�i�����g�𐶐����Apdf�ŏo�͂���
Sub testFormOfTournament()
    Call setUp
    
    Dim startTeamNum As Integer
    Dim endTeamNum As Integer
    Dim i As Integer
    Dim fullPathForTempFile As String
    
    
    startTeamNum = 4
    endTeamNum = 64
    
    For i = startTeamNum To endTeamNum
        teamsRange = i
        Call test
        Call makeTournament
        fullPathForTempFile = tempDir & i & ".pdf"
        tournamentWS.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fullPathForTempFile
        'tournamentWS.PrintOut

    Next i
End Sub
