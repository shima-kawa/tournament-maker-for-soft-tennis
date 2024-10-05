Attribute VB_Name = "ForGit"
Option Explicit

Sub ExportAll()
    Dim module                  As VBComponent      '// ���W���[��
    Dim moduleList              As VBComponents     '// VBA�v���W�F�N�g�̑S���W���[��
    Dim extension                                   '// ���W���[���̊g���q
    Dim sPath                                       '// �����Ώۃu�b�N�̃p�X
    Dim sFilePath                                   '// �G�N�X�|�[�g�t�@�C���p�X
    Dim TargetBook                                  '// �����Ώۃu�b�N�I�u�W�F�N�g
    
    '// �u�b�N���J����Ă��Ȃ��ꍇ�͌l�p�}�N���u�b�N�ipersonal.xlsb�j��ΏۂƂ���
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    '// �u�b�N���J����Ă���ꍇ�͕\�����Ă���u�b�N��ΏۂƂ���
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.Path
    
    '// �����Ώۃu�b�N�̃��W���[���ꗗ���擾
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBA�v���W�F�N�g�Ɋ܂܂��S�Ẵ��W���[�������[�v
    For Each module In moduleList
        '// �N���X
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// �t�H�[��
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frx���ꏏ�ɃG�N�X�|�[�g�����
            extension = "frm"
        '// �W�����W���[��
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// ���̑�
        Else
            '// �G�N�X�|�[�g�ΏۊO�̂��ߎ����[�v��
            GoTo CONTINUE
        End If
        
        '// �G�N�X�|�[�g���{
        sFilePath = sPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        '// �o�͐�m�F�p���O�o��
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

'// �w�胏�[�N�u�b�N�Ɏw��t�H���_�z���̃��W���[�����C���|�[�g����
'// �����P�F���[�N�u�b�N
'// �����Q�F���W���[���i�[�t�H���_�p�X
Sub ImportAll(a_TargetBook As Workbook, a_sModulePath As String)
    On Error Resume Next
    
    Dim oFso        As New FileSystemObject     '// FileSystemObject�I�u�W�F�N�g
    Dim sArModule() As String                   '// ���W���[���t�@�C���z��
    Dim sModule                                 '// ���W���[���t�@�C��
    Dim sExt        As String                   '// �g���q
    Dim iMsg                                    '// MsgBox�֐��߂�l
    
    iMsg = MsgBox("�����̃��W���[���͏㏑�����܂��B��낵���ł����H", vbOKCancel, "�㏑���m�F")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// �S���W���[���̃t�@�C���p�X���擾
    Call searchAllFile(a_sModulePath, sArModule)
    
    '// �S���W���[�������[�v
    For Each sModule In sArModule
        '// �g���q���������Ŏ擾
        sExt = LCase(oFso.GetExtensionName(sModule))
        
        '// �g���q��cls�Afrm�Abas�̂����ꂩ�̏ꍇ
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// �������W���[�����폜
            Call a_TargetBook.VBProject.VBComponents.Remove(a_TargetBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// ���W���[����ǉ�
            Call a_TargetBook.VBProject.VBComponents.Import(sModule)
            '// Import�m�F�p���O�o��
            Debug.Print sModule
        End If
    Next
End Sub

'// �w��t�H���_�z���̃t�@�C���p�X���擾
'// �����P�F�t�H���_�p�X
'// �����Q�F�t�@�C���p�X�z��
Sub searchAllFile(a_sFolder As String, s_ArFile() As String)
    Dim oFso        As New FileSystemObject
    Dim oFolder     As Folder
    Dim oSubFolder  As Folder
    Dim oFile       As File
    Dim i
    
    '// �t�H���_���Ȃ��ꍇ
    If (oFso.FolderExists(a_sFolder) = False) Then
        Exit Sub
    End If
    
    Set oFolder = oFso.GetFolder(a_sFolder)
    
    '// �T�u�t�H���_���ċA�i�T�u�t�H���_��T���K�v���Ȃ��ꍇ�͂���For�����폜���Ă��������j
    For Each oSubFolder In oFolder.SubFolders
        Call searchAllFile(oSubFolder.Path, s_ArFile)
    Next
    
    i = UBound(s_ArFile)
    
    '// �J�����g�t�H���_���̃t�@�C�����擾
    For Each oFile In oFolder.Files
        If (i <> 0 Or s_ArFile(i) <> "") Then
            i = i + 1
            ReDim Preserve s_ArFile(i)
        End If
        
        '// �t�@�C���p�X��z��Ɋi�[
        s_ArFile(i) = oFile.Path
    Next
End Sub

