Attribute VB_Name = "ForGit"
Option Explicit

Sub ExportAll()
    Dim module                  As VBComponent      '// モジュール
    Dim moduleList              As VBComponents     '// VBAプロジェクトの全モジュール
    Dim extension                                   '// モジュールの拡張子
    Dim sPath                                       '// 処理対象ブックのパス
    Dim sFilePath                                   '// エクスポートファイルパス
    Dim TargetBook                                  '// 処理対象ブックオブジェクト
    
    '// ブックが開かれていない場合は個人用マクロブック（personal.xlsb）を対象とする
    If (Workbooks.Count = 1) Then
        Set TargetBook = ThisWorkbook
    '// ブックが開かれている場合は表示しているブックを対象とする
    Else
        Set TargetBook = ActiveWorkbook
    End If
    
    sPath = TargetBook.Path
    
    '// 処理対象ブックのモジュール一覧を取得
    Set moduleList = TargetBook.VBProject.VBComponents
    
    '// VBAプロジェクトに含まれる全てのモジュールをループ
    For Each module In moduleList
        '// クラス
        If (module.Type = vbext_ct_ClassModule) Then
            extension = "cls"
        '// フォーム
        ElseIf (module.Type = vbext_ct_MSForm) Then
            '// .frxも一緒にエクスポートされる
            extension = "frm"
        '// 標準モジュール
        ElseIf (module.Type = vbext_ct_StdModule) Then
            extension = "bas"
        '// その他
        Else
            '// エクスポート対象外のため次ループへ
            GoTo CONTINUE
        End If
        
        '// エクスポート実施
        sFilePath = sPath & "\" & module.Name & "." & extension
        Call module.Export(sFilePath)
        
        '// 出力先確認用ログ出力
        Debug.Print sFilePath
CONTINUE:
    Next
End Sub

'// 指定ワークブックに指定フォルダ配下のモジュールをインポートする
'// 引数１：ワークブック
'// 引数２：モジュール格納フォルダパス
Sub ImportAll(a_TargetBook As Workbook, a_sModulePath As String)
    On Error Resume Next
    
    Dim oFso        As New FileSystemObject     '// FileSystemObjectオブジェクト
    Dim sArModule() As String                   '// モジュールファイル配列
    Dim sModule                                 '// モジュールファイル
    Dim sExt        As String                   '// 拡張子
    Dim iMsg                                    '// MsgBox関数戻り値
    
    iMsg = MsgBox("同名のモジュールは上書きします。よろしいですか？", vbOKCancel, "上書き確認")
    If (iMsg <> vbOK) Then
        Exit Sub
    End If
    
    ReDim sArModule(0)
    
    '// 全モジュールのファイルパスを取得
    Call searchAllFile(a_sModulePath, sArModule)
    
    '// 全モジュールをループ
    For Each sModule In sArModule
        '// 拡張子を小文字で取得
        sExt = LCase(oFso.GetExtensionName(sModule))
        
        '// 拡張子がcls、frm、basのいずれかの場合
        If (sExt = "cls" Or sExt = "frm" Or sExt = "bas") Then
            '// 同名モジュールを削除
            Call a_TargetBook.VBProject.VBComponents.Remove(a_TargetBook.VBProject.VBComponents(oFso.GetBaseName(sModule)))
            '// モジュールを追加
            Call a_TargetBook.VBProject.VBComponents.Import(sModule)
            '// Import確認用ログ出力
            Debug.Print sModule
        End If
    Next
End Sub

'// 指定フォルダ配下のファイルパスを取得
'// 引数１：フォルダパス
'// 引数２：ファイルパス配列
Sub searchAllFile(a_sFolder As String, s_ArFile() As String)
    Dim oFso        As New FileSystemObject
    Dim oFolder     As Folder
    Dim oSubFolder  As Folder
    Dim oFile       As File
    Dim i
    
    '// フォルダがない場合
    If (oFso.FolderExists(a_sFolder) = False) Then
        Exit Sub
    End If
    
    Set oFolder = oFso.GetFolder(a_sFolder)
    
    '// サブフォルダを再帰（サブフォルダを探す必要がない場合はこのFor文を削除してください）
    For Each oSubFolder In oFolder.SubFolders
        Call searchAllFile(oSubFolder.Path, s_ArFile)
    Next
    
    i = UBound(s_ArFile)
    
    '// カレントフォルダ内のファイルを取得
    For Each oFile In oFolder.Files
        If (i <> 0 Or s_ArFile(i) <> "") Then
            i = i + 1
            ReDim Preserve s_ArFile(i)
        End If
        
        '// ファイルパスを配列に格納
        s_ArFile(i) = oFile.Path
    Next
End Sub

