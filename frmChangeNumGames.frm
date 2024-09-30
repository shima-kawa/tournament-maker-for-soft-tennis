VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangeNumGames 
   Caption         =   "�Q�[�����̕ύX"
   ClientHeight    =   6090
   ClientLeft      =   180
   ClientTop       =   705
   ClientWidth     =   9330.001
   OleObjectBlob   =   "frmChangeNumGames.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "frmChangeNumGames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnChange_Click()
    Dim selectedIndex As Integer
    Dim numGames As Integer
    
    selectedIndex = cmbRounds.ListIndex
    numGames = txtNumGames.Value
    Call changeNumOfGames(selectedIndex + 1, numGames)
    
    Unload Me
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub txtNumGames_Change()
    Dim txt As String
    Dim i As Integer

    '��������0�ɂȂ�����I��
    If Len(txtNumGames.Text) = 0 Then Exit Sub
    
    '�e�L�X�g�{�b�N�X�̒l���擾
    txt = txtNumGames.Text
    
    '������������1��������납�烋�[�v
    For i = Len(txt) To 1 Step -1
        '�����������ȊO�̏ꍇ
        If IsNumeric(Mid(txt, i, 1)) = False Then
            '�����ȊO�̕������폜
            txt = Replace(txt, Mid(txt, i, 1), "")
        End If
    Next
    
    '�e�L�X�g�{�b�N�X�ɒl�����
    txtNumGames.Text = txt

End Sub

Private Sub UserForm_Initialize()
    setUp
    
    Dim teams As Integer
    Dim baseTeams As Integer
    Dim needRound As Integer
    Dim i As Integer
    
    teams = teamsRange.Value
    baseTeams = culNumberOfBaseTeams(teams)
    needRound = culNumberOfNeedRounds(baseTeams)

    For i = 1 To needRound - 1
        cmbRounds.AddItem i & "���"
    Next i
    cmbRounds.AddItem "����"
    
End Sub
