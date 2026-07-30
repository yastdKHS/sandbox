Attribute VB_Name = "modBuildTool"
Option Explicit
'==============================================================================
' modBuildTool
' 目的 : 本ツールのVBAソース一式（vba/modules 配下の .bas ファイル、および
'        vba/ThisWorkbook.txt）から、実際に動作する .xlsm ブックを
'        自動的に組み立てる。
'
' このモジュールは「組み立て専用」であり、完成後のブックにも残しておいて
' 差し支えない（将来バージョンアップ時に同じ手順で再ビルドできるよう、
' 保守用ユーティリティとして同梱する）。
'
' 前提条件（詳細は同フォルダの BUILD_README.md を参照）:
'   ・Windows版 Excel 2019 が必要（VBAのIDE操作を行うため）。
'   ・[ファイル]-[オプション]-[トラストセンター]-[トラストセンターの設定]-
'     [マクロの設定] で「VBA プロジェクト オブジェクト モデルへのアクセスを
'     信頼する」にチェックを入れておくこと（VBComponents.Import 等の
'     プログラムからのVBA操作に必要な設定）。
'
' 使い方:
'   1. 空の新規ブックを開く。
'   2. Alt+F11 → 挿入 → 標準モジュール にこのファイルの内容を貼り付ける。
'   3. イミディエイトウィンドウで BuildWorkbook を実行する。
'   4. ダイアログで vba/modules フォルダを選択する。
'   5. 完成した .xlsm の保存先ファイル名を指定する。
'==============================================================================

Public Sub BuildWorkbook()
    On Error GoTo ErrHandler

    ' トラストアクセスの事前確認（無効な場合はここで例外が発生する）
    Dim testCount As Long
    testCount = Application.VBE.VBProjects.Count

    Dim srcFolder As Variant
    srcFolder = PickFolder("VBAソースフォルダ（vba¥modules）を選択してください")
    If srcFolder = False Then Exit Sub

    Dim wbNew As Workbook
    Set wbNew = Workbooks.Add(xlWBATWorksheet)

    ImportAllModules wbNew, CStr(srcFolder)
    InjectThisWorkbookCode wbNew, CStr(srcFolder)

    ' 取り込んだ mod_UI.SetupAllSheets を実行し、画面・テーブルを初期構築する。
    Application.Run "'" & wbNew.Name & "'!mod_UI.SetupAllSheets"

    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="評価データ分析ツール.xlsm", _
        FileFilter:="Excel マクロ有効ブック (*.xlsm), *.xlsm")
    If savePath = False Then
        MsgBox "保存先が指定されなかったため、ブックは保存せずに残しています。" & vbCrLf & _
               "必要な場合は手動で「名前を付けて保存」（Excelマクロ有効ブック）を実行してください。", _
               vbExclamation, "ビルド中断"
        Exit Sub
    End If

    Application.DisplayAlerts = False
    wbNew.SaveAs Filename:=savePath, FileFormat:=52   ' xlOpenXMLWorkbookMacroEnabled (.xlsm)
    Application.DisplayAlerts = True

    MsgBox "ビルドが完了しました。" & vbCrLf & savePath, vbInformation, "ビルド完了"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "ビルド中にエラーが発生しました。" & vbCrLf & _
           "[トラストセンター]-[マクロの設定]-" & vbCrLf & _
           "「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が" & vbCrLf & _
           "有効になっているかご確認ください。" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "ビルド失敗"
End Sub

Private Function PickFolder(ByVal title As String) As Variant
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = title
    If fd.Show = -1 Then
        PickFolder = fd.SelectedItems(1)
    Else
        PickFolder = False
    End If
End Function

' 指定フォルダ内のすべての .bas ファイル（標準モジュール）をインポートする。
' 自分自身（modBuildTool.bas）がそのフォルダに置かれていた場合は多重登録を避けるため
' スキップする。
Private Sub ImportAllModules(ByVal wb As Workbook, ByVal folderPath As String)
    Dim fPath As String
    fPath = Dir(folderPath & Application.PathSeparator & "*.bas")
    Do While Len(fPath) > 0
        If LCase$(fPath) <> "modbuildtool.bas" Then
            wb.VBProject.VBComponents.Import folderPath & Application.PathSeparator & fPath
        End If
        fPath = Dir()
    Loop
End Sub

' vba/ThisWorkbook.txt の内容を、新規ブックの既存 ThisWorkbook コンポーネントへ追記する。
' ThisWorkbook は標準モジュールと異なり Import で新規作成できないコンポーネントのため、
' コードモジュールへ直接文字列として書き込む。
Private Sub InjectThisWorkbookCode(ByVal wb As Workbook, ByVal folderPath As String)
    Dim twPath As String
    twPath = folderPath & Application.PathSeparator & ".." & Application.PathSeparator & "ThisWorkbook.txt"
    If Dir(twPath) = "" Then
        MsgBox "ThisWorkbook.txt が見つかりませんでした（" & twPath & "）。" & vbCrLf & _
               "手動で ThisWorkbook モジュールへ貼り付けてください。", vbExclamation
        Exit Sub
    End If

    Dim codeText As String
    Dim fileNum As Integer
    fileNum = FreeFile
    Open twPath For Input As #fileNum
    Dim ln As String
    Do While Not EOF(fileNum)
        Line Input #fileNum, ln
        codeText = codeText & ln & vbCrLf
    Loop
    Close #fileNum

    wb.VBProject.VBComponents("ThisWorkbook").CodeModule.AddFromString codeText
End Sub
