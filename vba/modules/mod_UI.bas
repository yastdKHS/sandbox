Attribute VB_Name = "mod_UI"
Option Explicit
'==============================================================================
' mod_UI
' 目的 : 6画面（ホーム／データ取込／分析／設定／データベース／ログ）の
'        画面構築（フォームコントロールの生成）と、ボタン等からの
'        イベント振り分けを担当する。
'
' 設計方針:
'   ・すべてのコントロールは Excel 標準の「フォームコントロール」
'     （ActiveXコントロールではない）を使用する。ActiveXは環境依存の
'     不具合や互換性問題が発生しやすく、10年以上の長期利用に適さないため。
'   ・画面のセル位置・コントロール名は mod_Common の定数で一元管理する。
'   ・SetupAllSheets はブック構築時だけでなく、レイアウトが崩れた場合の
'     修復用マクロとしても再実行可能（べき等）。
'   ・本モジュールは「表示」のみを担当し、統計計算・データ加工は行わない
'     （mod_Statistics / mod_Analysis / mod_Database 等に委譲する）。
'==============================================================================

'==============================================================================
' 画面構築（初期セットアップ・修復用）
'==============================================================================

' 全画面を初期構築する。ビルド時、および保守目的での再実行（修復）に使用する。
Public Sub SetupAllSheets()
    ' 画面構築中は Application.ScreenUpdating を False にしない。
    ' セル結合（Range.Merge）は本モジュールでは一切使用しない方針とした。
    ' プログラムからのセル結合は、直後に同じセルへ ClearContents や
    ' Value代入などの操作を行うと実行時エラー1004を起こすことがあり
    ' （実機で複数回確認済み）、ScreenUpdatingの状態にも左右されず
    ' 不安定であるため、複数セルにまたがる表示が必要な箇所は
    ' 列幅を広げ、空白の隣接セルへ自然にはみ出させる方式に統一している。
    On Error GoTo ErrHandler
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    mod_Database.EnsureSchema
    SetupHomeSheet
    SetupImportSheet
    SetupAnalysisSheet
    SetupSettingsSheet
    SetupDatabaseSheet
    SetupLogSheet
    ThisWorkbook.Worksheets(mod_Common.SH_HOME).Activate

    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    mod_Common.HandleError "mod_UI", "SetupAllSheets"
End Sub

Public Sub SetupHomeSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_HOME)
    ws.Activate

    ws.Range("B2").Value = mod_Common.APP_NAME
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 16
    ws.Range("B3").Value = "バージョン " & mod_Common.APP_VERSION

    AddButtonControl ws, mod_Common.HOME_BTN_IMPORT, mod_Common.HOME_RANGE_BTN_IMPORT, _
        "① データ取込", "mod_UI.Btn_Home_GoImport"
    AddButtonControl ws, mod_Common.HOME_BTN_ANALYSIS, mod_Common.HOME_RANGE_BTN_ANALYSIS, _
        "② 分析", "mod_UI.Btn_Home_GoAnalysis"
    AddButtonControl ws, mod_Common.HOME_BTN_SETTINGS, mod_Common.HOME_RANGE_BTN_SETTINGS, _
        "③ 設定", "mod_UI.Btn_Home_GoSettings"
    AddButtonControl ws, mod_Common.HOME_BTN_DATABASE, mod_Common.HOME_RANGE_BTN_DATABASE, _
        "④ データベース", "mod_UI.Btn_Home_GoDatabase"
    AddButtonControl ws, mod_Common.HOME_BTN_LOG, mod_Common.HOME_RANGE_BTN_LOG, _
        "⑤ ログ", "mod_UI.Btn_Home_GoLog"

    ws.Columns("F").ColumnWidth = 60
    ws.Range(mod_Common.HOME_CELL_SUMMARY).WrapText = True
    ws.Range(mod_Common.HOME_CELL_SUMMARY).VerticalAlignment = xlTop

    CreateNavBar ws
    RefreshHomeSummary
End Sub

Public Sub SetupImportSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_IMPORT)
    ws.Activate

    ws.Range(mod_Common.IMP_CELL_TITLE).Value = "データ取込"
    ws.Range(mod_Common.IMP_CELL_TITLE).Font.Bold = True
    ws.Range(mod_Common.IMP_CELL_TITLE).Font.Size = 14

    ws.Range("B4").Value = "① 下のボタンから外部Excelファイルを選択してください。"
    AddButtonControl ws, mod_Common.IMP_BTN_SELECTFILE, mod_Common.IMP_RANGE_BTN_SELECTFILE, _
        "データ取込を開始", "mod_UI.Btn_Import_Start"
    ws.Range("B7").Value = "② 総合成績シートを確認してください（候補が自動選択されます）。"
    AddListBoxControl ws, mod_Common.IMP_LISTBOX_SHEETS, mod_Common.IMP_RANGE_LISTBOX_SHEETS, False
    ws.Range("G8").Value = "年度（元データに列が無い場合のみ入力）"
    ws.Range("G10").Value = "クラスコード（元データに列が無い場合のみ入力）"
    AddButtonControl ws, mod_Common.IMP_BTN_CONFIRMSHEET, mod_Common.IMP_RANGE_BTN_CONFIRMSHEET, _
        "シート決定", "mod_UI.Btn_Import_ConfirmSheet"

    ws.Range(mod_Common.IMP_CELL_MAPPING_HINT).Value = "③ 評価項目のマッピングを確認・修正してください（プルダウンで変更できます）。"
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 22
    AddButtonControl ws, mod_Common.IMP_BTN_PROCEED_CONFIRM, mod_Common.IMP_RANGE_BTN_PROCEED_CONFIRM, _
        "内容確認へ", "mod_UI.Btn_Import_ProceedConfirm"
    AddButtonControl ws, mod_Common.IMP_BTN_SAVETEMPLATE, mod_Common.IMP_RANGE_BTN_SAVETEMPLATE, _
        "テンプレート保存", "mod_UI.Btn_Import_SaveTemplate"

    ws.Range("B52").Value = "④ 登録内容を確認し、問題なければ「登録」を押してください。"
    AddCheckBoxControl ws, mod_Common.IMP_CHK_OVERWRITE, mod_Common.IMP_RANGE_CHK_OVERWRITE, _
        "既存データ（同一年度・クラス・学生番号・項目）を上書きする"
    AddButtonControl ws, mod_Common.IMP_BTN_COMMIT, mod_Common.IMP_RANGE_BTN_COMMIT, _
        "登録", "mod_UI.Btn_Import_Commit"
    AddButtonControl ws, mod_Common.IMP_BTN_CANCEL, mod_Common.IMP_RANGE_BTN_CANCEL, _
        "キャンセル", "mod_UI.Btn_Import_Cancel"

    ws.Columns(mod_Common.IMP_MAPPING_HELPER_COL).Hidden = True

    CreateNavBar ws
    ResetImportSheet
End Sub

Public Sub SetupAnalysisSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_ANALYSIS)
    ws.Activate

    ws.Range(mod_Common.ANL_CELL_TITLE).Value = "分析"
    ws.Range(mod_Common.ANL_CELL_TITLE).Font.Bold = True
    ws.Range(mod_Common.ANL_CELL_TITLE).Font.Size = 14

    ws.Range("B4").Value = "年度（複数選択可）"
    ws.Range("D4").Value = "クラスコード（複数選択可）"
    ws.Range("F4").Value = "評価項目（1つ選択）"
    AddListBoxControl ws, mod_Common.ANL_LISTBOX_YEARS, mod_Common.ANL_RANGE_LISTBOX_YEARS, True
    AddListBoxControl ws, mod_Common.ANL_LISTBOX_CLASSES, mod_Common.ANL_RANGE_LISTBOX_CLASSES, True
    AddListBoxControl ws, mod_Common.ANL_LISTBOX_ITEM, mod_Common.ANL_RANGE_LISTBOX_ITEM, False

    AddCheckBoxControl ws, mod_Common.ANL_CHK_NORMAL, mod_Common.ANL_RANGE_CHK_NORMAL, _
        "正規分布を重畳表示", True
    AddCheckBoxControl ws, mod_Common.ANL_CHK_OVERLAY, mod_Common.ANL_RANGE_CHK_OVERLAY, _
        "複数年度を重畳表示", True
    AddOptionButtonControl ws, mod_Common.ANL_OPT_BINCOUNT, mod_Common.ANL_RANGE_OPT_BINCOUNT, _
        "ビン数指定", True
    AddOptionButtonControl ws, mod_Common.ANL_OPT_BINWIDTH, mod_Common.ANL_RANGE_OPT_BINWIDTH, _
        "ビン幅指定", False
    ws.Range(mod_Common.ANL_CELL_BINCOUNT).Value = 10
    ws.Range(mod_Common.ANL_CELL_BINWIDTH).Value = 10

    AddButtonControl ws, mod_Common.ANL_BTN_RUN, mod_Common.ANL_RANGE_BTN_RUN, _
        "分析実行", "mod_UI.Btn_Analysis_Run"

    EnsureChartObject ws, mod_Common.ANL_CHART_HIST, mod_Common.ANL_RANGE_CHART_HIST
    EnsureChartObject ws, mod_Common.ANL_CHART_CUM, mod_Common.ANL_RANGE_CHART_CUM
    EnsureChartObject ws, mod_Common.ANL_CHART_BOX, mod_Common.ANL_RANGE_CHART_BOX

    ws.Range(mod_Common.ANL_STUDENTLIST_TITLE).Value = "個人別一覧（偏差値）"
    Dim hdrRow As Long, colBase As Long
    hdrRow = mod_Common.ANL_STUDENTLIST_HEADER_ROW
    colBase = mod_Common.ANL_STUDENTLIST_COL
    ws.Cells(hdrRow, colBase).Value = "学生番号"
    ws.Cells(hdrRow, colBase + 1).Value = "氏名"
    ws.Cells(hdrRow, colBase + 2).Value = "点数"
    ws.Cells(hdrRow, colBase + 3).Value = "偏差値"
    ws.Range(ws.Cells(hdrRow, colBase), ws.Cells(hdrRow, colBase + 3)).Font.Bold = True

    ws.Range(mod_Common.ANL_STATS_START).Offset(-1, 0).Value = "基本統計量"
    ws.Range(mod_Common.ANL_STATS_START).Offset(-1, 0).Font.Bold = True

    RefreshAnalysisSelectionLists ws
    CreateNavBar ws
End Sub

Public Sub SetupSettingsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_SETTINGS)
    ws.Activate

    ws.Range("B2").Value = "設定"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 14
    ws.Range("B3").Value = "評価項目の表示名・重み(%)は下表のセルを直接編集できます。編集後は「反映」を押してください。"
    ws.Range("I3").Value = "詳細設定（キーと値。通常は変更不要です）"

    AddButtonControl ws, mod_Common.SET_BTN_APPLY, mod_Common.SET_RANGE_BTN_APPLY, _
        "反映（総合評価を再計算）", "mod_UI.Btn_Settings_Apply"

    ws.Range("I12").Value = "保存済みテンプレート"
    AddListBoxControl ws, mod_Common.SET_LISTBOX_TEMPLATES, mod_Common.SET_RANGE_LISTBOX_TEMPLATES, False
    AddButtonControl ws, mod_Common.SET_BTN_DELETETEMPLATE, mod_Common.SET_RANGE_BTN_DELETETEMPLATE, _
        "選択したテンプレートを削除", "mod_UI.Btn_Settings_DeleteTemplate"

    CreateNavBar ws
    RefreshTemplateList ws
End Sub

Public Sub SetupDatabaseSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_DATABASE)
    ws.Activate

    ws.Range("B2").Value = "データベース"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 14
    ws.Range("B3").Value = "学生ごとの総合評価一覧です。表のオートフィルタで年度・クラスを絞り込めます。"

    ws.Range(mod_Common.DB_CELL_DELETE_YEAR_LABEL).Value = "年度指定（削除用）"
    ws.Range(mod_Common.DB_CELL_DELETE_CLASS_LABEL).Value = "クラスコード指定（削除用）"
    AddButtonControl ws, mod_Common.DB_BTN_DELETE, mod_Common.DB_RANGE_BTN_DELETE, _
        "指定データを削除", "mod_UI.Btn_Database_Delete"
    AddButtonControl ws, mod_Common.DB_BTN_REFRESH, mod_Common.DB_RANGE_BTN_REFRESH, _
        "再集計・更新", "mod_UI.Btn_Database_Refresh"

    CreateNavBar ws
End Sub

Public Sub SetupLogSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_LOG)
    ws.Activate

    ws.Range("B2").Value = "ログ"
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 14
    ws.Range("B3").Value = "操作・エラーの履歴です。保持件数の上限を超えると古い記録から自動的に削除されます。"

    AddButtonControl ws, mod_Common.LOG_BTN_CLEAR, mod_Common.LOG_RANGE_BTN_CLEAR, _
        "ログを全件クリア", "mod_UI.Btn_Log_Clear"

    CreateNavBar ws
End Sub

'==============================================================================
' ナビゲーションバー（全画面共通）
'==============================================================================
Public Sub CreateNavBar(ByVal ws As Worksheet)
    AddButtonControl ws, mod_Common.NAV_BTN_HOME, mod_Common.NAV_RANGE_HOME, "ホーム", "mod_UI.Nav_Home"
    AddButtonControl ws, mod_Common.NAV_BTN_IMPORT, mod_Common.NAV_RANGE_IMPORT, "データ取込", "mod_UI.Nav_Import"
    AddButtonControl ws, mod_Common.NAV_BTN_ANALYSIS, mod_Common.NAV_RANGE_ANALYSIS, "分析", "mod_UI.Nav_Analysis"
    AddButtonControl ws, mod_Common.NAV_BTN_SETTINGS, mod_Common.NAV_RANGE_SETTINGS, "設定", "mod_UI.Nav_Settings"
    AddButtonControl ws, mod_Common.NAV_BTN_DATABASE, mod_Common.NAV_RANGE_DATABASE, "データベース", "mod_UI.Nav_Database"
    AddButtonControl ws, mod_Common.NAV_BTN_LOG, mod_Common.NAV_RANGE_LOG, "ログ", "mod_UI.Nav_Log"
End Sub

Public Sub NavigateTo(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Activate

    Select Case sheetName
        Case mod_Common.SH_HOME
            RefreshHomeSummary
        Case mod_Common.SH_SETTINGS
            RefreshTemplateList ws
        Case mod_Common.SH_ANALYSIS
            RefreshAnalysisSelectionLists ws
    End Select
End Sub

Public Sub Nav_Home(): NavigateTo mod_Common.SH_HOME: End Sub
Public Sub Nav_Import(): NavigateTo mod_Common.SH_IMPORT: End Sub
Public Sub Nav_Analysis(): NavigateTo mod_Common.SH_ANALYSIS: End Sub
Public Sub Nav_Settings(): NavigateTo mod_Common.SH_SETTINGS: End Sub
Public Sub Nav_Database(): NavigateTo mod_Common.SH_DATABASE: End Sub
Public Sub Nav_Log(): NavigateTo mod_Common.SH_LOG: End Sub

'==============================================================================
' ホーム画面
'==============================================================================
Public Sub Btn_Home_GoImport()
    NavigateTo mod_Common.SH_IMPORT
    mod_Import.StartImport
End Sub
Public Sub Btn_Home_GoAnalysis(): NavigateTo mod_Common.SH_ANALYSIS: End Sub
Public Sub Btn_Home_GoSettings(): NavigateTo mod_Common.SH_SETTINGS: End Sub
Public Sub Btn_Home_GoDatabase(): NavigateTo mod_Common.SH_DATABASE: End Sub
Public Sub Btn_Home_GoLog(): NavigateTo mod_Common.SH_LOG: End Sub

Public Sub RefreshHomeSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_HOME)

    Dim years As Collection
    Set years = mod_Database.GetDistinctYears()

    Dim scoresTbl As ListObject, sumTbl As ListObject
    Set scoresTbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    Set sumTbl = mod_Common.GetTableSafe(mod_Common.SH_DATABASE, mod_Common.TBL_SUMMARY)

    Dim msg As String
    msg = "登録年度数: " & years.Count & vbCrLf & _
          "登録学生（延べ）件数: " & mod_Common.TableRowCount(sumTbl) & vbCrLf & _
          "評価点数レコード数: " & mod_Common.TableRowCount(scoresTbl) & vbCrLf & vbCrLf & _
          "「① データ取込」から外部Excelを取り込み、「② 分析」で年度・クラス・評価項目を選んで比較できます。"
    ws.Range(mod_Common.HOME_CELL_SUMMARY).Value = msg
End Sub

'==============================================================================
' データ取込画面のボタン
'==============================================================================
Public Sub Btn_Import_Start(): mod_Import.StartImport: End Sub
Public Sub Btn_Import_ConfirmSheet(): mod_Import.ConfirmSheetSelection: End Sub
Public Sub Btn_Import_ProceedConfirm(): mod_Import.ProceedToConfirm: End Sub
Public Sub Btn_Import_Commit(): mod_Import.CommitImport: End Sub
Public Sub Btn_Import_Cancel(): mod_Import.CancelImport: End Sub
Public Sub Btn_Import_SaveTemplate(): mod_Import.SaveCurrentMappingAsTemplate: End Sub

' ウィザードの表示段階を切り替える。0=未開始 1=シート選択 2=マッピング 3=確認・登録
Public Sub ShowImportStep(ByVal stepNo As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_IMPORT)

    SetShapeVisible ws, mod_Common.IMP_LISTBOX_SHEETS, (stepNo >= 1)
    SetShapeVisible ws, mod_Common.IMP_BTN_CONFIRMSHEET, (stepNo = 1)

    SetRowsHidden ws, mod_Common.IMP_MAPPING_HEADER_ROW, _
        mod_Common.IMP_MAPPING_START_ROW + mod_Common.IMP_MAPPING_MAX_ROWS - 1, (stepNo <> 2)
    SetShapeVisible ws, mod_Common.IMP_BTN_PROCEED_CONFIRM, (stepNo = 2)
    SetShapeVisible ws, mod_Common.IMP_BTN_SAVETEMPLATE, (stepNo = 2)

    Dim confirmRow As Long
    confirmRow = ws.Range(mod_Common.IMP_CELL_CONFIRM_SUMMARY).Row
    SetRowsHidden ws, confirmRow, confirmRow + mod_Common.IMP_CONFIRM_SUMMARY_ROWS, (stepNo <> 3)
    SetShapeVisible ws, mod_Common.IMP_CHK_OVERWRITE, (stepNo = 3)
    SetShapeVisible ws, mod_Common.IMP_BTN_COMMIT, (stepNo = 3)

    SetShapeVisible ws, mod_Common.IMP_BTN_CANCEL, (stepNo >= 1)
End Sub

Public Sub SetImportStatus(ByVal text As String)
    ThisWorkbook.Worksheets(mod_Common.SH_IMPORT).Range(mod_Common.IMP_CELL_STATUS).Value = text
End Sub

Public Sub SetImportConfirmSummary(ByVal text As String)
    ThisWorkbook.Worksheets(mod_Common.SH_IMPORT).Range(mod_Common.IMP_CELL_CONFIRM_SUMMARY).Value = text
End Sub

Public Sub ResetImportSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_IMPORT)

    ws.Range(mod_Common.IMP_CELL_FILEPATH).ClearContents
    ws.Range(mod_Common.IMP_CELL_YEAR_MANUAL).ClearContents
    ws.Range(mod_Common.IMP_CELL_CLASS_MANUAL).ClearContents
    PopulateFormListBox ws, mod_Common.IMP_LISTBOX_SHEETS, New Collection

    Dim r As Long
    For r = mod_Common.IMP_MAPPING_START_ROW To mod_Common.IMP_MAPPING_START_ROW + mod_Common.IMP_MAPPING_MAX_ROWS - 1
        ws.Cells(r, mod_Common.IMP_MAPPING_COL_HEADER).ClearContents
        ws.Cells(r, mod_Common.IMP_MAPPING_COL_SAMPLE).ClearContents
        With ws.Cells(r, mod_Common.IMP_MAPPING_COL_TARGET)
            .ClearContents
            .Validation.Delete
        End With
    Next r

    ws.Range(mod_Common.IMP_CELL_CONFIRM_SUMMARY).ClearContents
    ShowImportStep 0
    SetImportStatus "「データ取込を開始」ボタンを押してください。"
End Sub

' マッピング表（見出し・サンプル値・マッピング先ドロップダウン）を描画する。
Public Sub RenderImportMappingTable(ByRef headers() As String, ByRef samples() As String, _
                                     ByRef suggestions() As String, ByVal choices As Collection)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_IMPORT)

    ' 選択肢はヘルパー列（非表示）へ書き出し、Data Validation の参照範囲とする
    ' （リテラル文字列指定は255文字制限があるため範囲参照方式にしている）。
    Dim helperCol As Long
    helperCol = mod_Common.IMP_MAPPING_HELPER_COL
    ws.Columns(helperCol).ClearContents
    Dim i As Long
    For i = 1 To choices.Count
        ws.Cells(i, helperCol).Value = choices(i)
    Next i
    Dim listFormula As String
    listFormula = "=$" & ColLetter(helperCol) & "$1:$" & ColLetter(helperCol) & "$" & Application.Max(choices.Count, 1)

    ws.Cells(mod_Common.IMP_MAPPING_HEADER_ROW, mod_Common.IMP_MAPPING_COL_HEADER).Value = "元の見出し"
    ws.Cells(mod_Common.IMP_MAPPING_HEADER_ROW, mod_Common.IMP_MAPPING_COL_SAMPLE).Value = "サンプル値"
    ws.Cells(mod_Common.IMP_MAPPING_HEADER_ROW, mod_Common.IMP_MAPPING_COL_TARGET).Value = "マッピング先"
    ws.Range(ws.Cells(mod_Common.IMP_MAPPING_HEADER_ROW, mod_Common.IMP_MAPPING_COL_HEADER), _
             ws.Cells(mod_Common.IMP_MAPPING_HEADER_ROW, mod_Common.IMP_MAPPING_COL_TARGET)).Font.Bold = True

    Dim n As Long
    n = UBound(headers) - LBound(headers) + 1
    If n > mod_Common.IMP_MAPPING_MAX_ROWS Then n = mod_Common.IMP_MAPPING_MAX_ROWS

    For i = 1 To n
        Dim row As Long
        row = mod_Common.IMP_MAPPING_START_ROW + i - 1
        ws.Cells(row, mod_Common.IMP_MAPPING_COL_HEADER).Value = headers(LBound(headers) + i - 1)
        ws.Cells(row, mod_Common.IMP_MAPPING_COL_SAMPLE).Value = samples(LBound(samples) + i - 1)

        Dim targetCell As Range
        Set targetCell = ws.Cells(row, mod_Common.IMP_MAPPING_COL_TARGET)
        targetCell.Value = suggestions(LBound(suggestions) + i - 1)
        On Error Resume Next
        targetCell.Validation.Delete
        targetCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=listFormula
        On Error GoTo 0
    Next i

    For i = n + 1 To mod_Common.IMP_MAPPING_MAX_ROWS
        Dim row2 As Long
        row2 = mod_Common.IMP_MAPPING_START_ROW + i - 1
        ws.Cells(row2, mod_Common.IMP_MAPPING_COL_HEADER).ClearContents
        ws.Cells(row2, mod_Common.IMP_MAPPING_COL_SAMPLE).ClearContents
        With ws.Cells(row2, mod_Common.IMP_MAPPING_COL_TARGET)
            .ClearContents
            On Error Resume Next
            .Validation.Delete
            On Error GoTo 0
        End With
    Next i
End Sub

Public Function GetImportMappingSelections(ByVal colCount As Long) As Collection
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_IMPORT)
    Dim result As New Collection
    Dim i As Long
    For i = 1 To colCount
        Dim row As Long
        row = mod_Common.IMP_MAPPING_START_ROW + i - 1
        result.Add CStr(ws.Cells(row, mod_Common.IMP_MAPPING_COL_TARGET).Value)
    Next i
    Set GetImportMappingSelections = result
End Function

'==============================================================================
' 分析画面
'==============================================================================
Public Sub Btn_Analysis_Run(): mod_Analysis.RunAnalysis: End Sub

Public Sub RefreshAnalysisSelectionLists(ByVal ws As Worksheet)
    PopulateFormListBox ws, mod_Common.ANL_LISTBOX_YEARS, SortStrings(mod_Database.GetDistinctYears())
    PopulateFormListBox ws, mod_Common.ANL_LISTBOX_CLASSES, SortStrings(mod_Database.GetDistinctClasses())

    Dim itemCodes As Collection
    Set itemCodes = mod_Database.GetActiveItemCodes()
    Dim itemNames As New Collection
    Dim ic As Variant
    For Each ic In itemCodes
        itemNames.Add mod_Settings.GetItemDisplayName(CStr(ic))
    Next ic
    PopulateFormListBox ws, mod_Common.ANL_LISTBOX_ITEM, itemNames
End Sub

Public Sub GetAnalysisGraphOptions(ByRef binMode As String, ByRef binCount As Long, ByRef binWidth As Double, _
                                    ByRef showNormal As Boolean, ByRef overlayYears As Boolean)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_ANALYSIS)

    If GetOptionButtonValue(ws, mod_Common.ANL_OPT_BINWIDTH) Then
        binMode = "WIDTH"
    Else
        binMode = "COUNT"
    End If
    binCount = CLng(NzNum(ws.Range(mod_Common.ANL_CELL_BINCOUNT).Value, 10))
    binWidth = NzNum(ws.Range(mod_Common.ANL_CELL_BINWIDTH).Value, 10)
    showNormal = GetCheckBoxValue(ws, mod_Common.ANL_CHK_NORMAL)
    overlayYears = GetCheckBoxValue(ws, mod_Common.ANL_CHK_OVERLAY)
End Sub

Public Sub RenderStatsTable(ByRef labels() As String, ByRef values() As Variant)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_ANALYSIS)
    Dim startCell As Range
    Set startCell = ws.Range(mod_Common.ANL_STATS_START)

    ws.Range(startCell, startCell.Offset(20, 1)).ClearContents

    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        startCell.Offset(i - LBound(labels), 0).Value = labels(i)
        startCell.Offset(i - LBound(labels), 1).Value = values(i)
    Next i
End Sub

Public Sub RenderStudentDetailList(ByRef studentNos() As String, ByRef studentNames() As String, _
                                    ByRef scoreVals() As Double, ByRef devVals() As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_ANALYSIS)

    Dim startRow As Long, colBase As Long, maxRows As Long
    startRow = mod_Common.ANL_STUDENTLIST_START_ROW
    colBase = mod_Common.ANL_STUDENTLIST_COL
    maxRows = mod_Common.ANL_STUDENTLIST_MAX_ROWS

    ws.Range(ws.Cells(startRow, colBase), ws.Cells(startRow + maxRows - 1, colBase + 3)).ClearContents

    Dim n As Long
    On Error Resume Next
    n = UBound(studentNos) - LBound(studentNos) + 1
    On Error GoTo 0
    If n = 0 Then Exit Sub
    If n > maxRows Then n = maxRows   ' 大規模データでも画面が肥大化しないよう表示件数を制限

    Dim i As Long
    For i = 1 To n
        ws.Cells(startRow + i - 1, colBase).Value = studentNos(LBound(studentNos) + i - 1)
        ws.Cells(startRow + i - 1, colBase + 1).Value = studentNames(LBound(studentNames) + i - 1)
        ws.Cells(startRow + i - 1, colBase + 2).Value = scoreVals(LBound(scoreVals) + i - 1)
        ws.Cells(startRow + i - 1, colBase + 3).Value = Application.Round(devVals(LBound(devVals) + i - 1), 1)
    Next i
End Sub

'==============================================================================
' 設定画面
'==============================================================================
Public Sub Btn_Settings_Apply()
    mod_Settings.ApplyItemSettingsChanged
    MsgBox "総合評価を再計算しました。", vbInformation, mod_Common.APP_NAME
End Sub

Public Sub Btn_Settings_DeleteTemplate()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_SETTINGS)
    Dim tplName As String
    tplName = GetFormListBoxSelection(ws, mod_Common.SET_LISTBOX_TEMPLATES)
    If Len(tplName) = 0 Then
        MsgBox "削除するテンプレートを選択してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If
    If MsgBox("テンプレート「" & tplName & "」を削除します。よろしいですか？", vbYesNo + vbExclamation, mod_Common.APP_NAME) = vbYes Then
        mod_Settings.DeleteTemplate tplName
        RefreshTemplateList ws
    End If
End Sub

Public Sub RefreshTemplateList(ByVal ws As Worksheet)
    PopulateFormListBox ws, mod_Common.SET_LISTBOX_TEMPLATES, mod_Settings.GetTemplateNames()
End Sub

'==============================================================================
' データベース画面
'==============================================================================
Public Sub Btn_Database_Refresh()
    mod_Common.BeginBusy
    mod_Database.RecalcSummary
    mod_Common.EndBusy
    MsgBox "総合評価を再集計しました。", vbInformation, mod_Common.APP_NAME
End Sub

Public Sub Btn_Database_Delete()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_DATABASE)
    Dim y As String, c As String
    y = Trim$(CStr(ws.Range(mod_Common.DB_CELL_DELETE_YEAR).Value))
    c = mod_Common.FormatClassCode(ws.Range(mod_Common.DB_CELL_DELETE_CLASS).Value)

    If Len(y) = 0 Or Len(c) = 0 Then
        MsgBox "削除する年度とクラスコードを入力してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    If MsgBox("年度=" & y & " クラス=" & c & " のデータを完全に削除します。" & vbCrLf & _
              "この操作は元に戻せません。よろしいですか？", vbYesNo + vbExclamation, mod_Common.APP_NAME) = vbYes Then
        mod_Common.BeginBusy
        mod_Database.DeleteByYearClass y, c
        mod_Common.EndBusy
        MsgBox "削除しました。", vbInformation, mod_Common.APP_NAME
    End If
End Sub

'==============================================================================
' ログ画面
'==============================================================================
Public Sub Btn_Log_Clear()
    If MsgBox("ログを全件削除します。よろしいですか？", vbYesNo + vbExclamation, mod_Common.APP_NAME) = vbYes Then
        mod_Logging.ClearLog
        MsgBox "ログをクリアしました。", vbInformation, mod_Common.APP_NAME
    End If
End Sub

'==============================================================================
' フォームコントロール共通ヘルパー
'==============================================================================

Private Sub RemoveShapeIfExists(ByVal ws As Worksheet, ByVal name As String)
    On Error Resume Next
    ws.Shapes(name).Delete
    On Error GoTo 0
End Sub

Private Function AddControlByRange(ByVal ws As Worksheet, ByVal ctrlType As XlFormControl, _
                                    ByVal name As String, ByVal rangeAddr As String) As Shape
    RemoveShapeIfExists ws, name
    Dim r As Range
    Set r = ws.Range(rangeAddr)
    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(ctrlType, r.Left, r.Top, r.Width, r.Height)
    shp.Name = name
    Set AddControlByRange = shp
End Function

Public Function AddButtonControl(ByVal ws As Worksheet, ByVal name As String, ByVal rangeAddr As String, _
                                  ByVal caption As String, ByVal onAction As String) As Shape
    Dim shp As Shape
    Set shp = AddControlByRange(ws, xlButtonControl, name, rangeAddr)
    ws.Buttons(name).Caption = caption
    shp.OnAction = onAction
    Set AddButtonControl = shp
End Function

Public Function AddListBoxControl(ByVal ws As Worksheet, ByVal name As String, ByVal rangeAddr As String, _
                                   ByVal multiSelect As Boolean) As Shape
    Dim shp As Shape
    Set shp = AddControlByRange(ws, xlListBox, name, rangeAddr)
    shp.ControlFormat.MultiSelect = IIf(multiSelect, xlExtended, xlNone)
    Set AddListBoxControl = shp
End Function

Public Function AddCheckBoxControl(ByVal ws As Worksheet, ByVal name As String, ByVal rangeAddr As String, _
                                    ByVal caption As String, Optional ByVal initialValue As Boolean = False) As Shape
    Dim shp As Shape
    Set shp = AddControlByRange(ws, xlCheckBox, name, rangeAddr)
    ws.CheckBoxes(name).Caption = caption
    shp.ControlFormat.Value = IIf(initialValue, xlOn, xlOff)
    Set AddCheckBoxControl = shp
End Function

Public Function AddOptionButtonControl(ByVal ws As Worksheet, ByVal name As String, ByVal rangeAddr As String, _
                                        ByVal caption As String, Optional ByVal initialValue As Boolean = False) As Shape
    Dim shp As Shape
    Set shp = AddControlByRange(ws, xlOptionButton, name, rangeAddr)
    ws.OptionButtons(name).Caption = caption
    shp.ControlFormat.Value = IIf(initialValue, xlOn, xlOff)
    Set AddOptionButtonControl = shp
End Function

Private Sub EnsureChartObject(ByVal ws As Worksheet, ByVal name As String, ByVal rangeAddr As String)
    On Error Resume Next
    ws.ChartObjects(name).Delete
    On Error GoTo 0
    Dim r As Range
    Set r = ws.Range(rangeAddr)
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(r.Left, r.Top, r.Width, r.Height)
    co.Name = name
End Sub

Public Sub SetShapeVisible(ByVal ws As Worksheet, ByVal name As String, ByVal isVisible As Boolean)
    On Error Resume Next
    ws.Shapes(name).Visible = isVisible
    On Error GoTo 0
End Sub

Private Sub SetRowsHidden(ByVal ws As Worksheet, ByVal rowFrom As Long, ByVal rowTo As Long, ByVal isHidden As Boolean)
    ws.Rows(rowFrom & ":" & rowTo).Hidden = isHidden
End Sub

Public Function GetFormListBoxSelection(ByVal ws As Worksheet, ByVal name As String) As String
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(name)
    If Not shp Is Nothing Then
        With shp.ControlFormat
            If .ListIndex >= 1 Then GetFormListBoxSelection = CStr(.List(.ListIndex))
        End With
    End If
    On Error GoTo 0
End Function

Public Function GetSelectedListItems(ByVal ws As Worksheet, ByVal name As String) As Collection
    Dim result As New Collection
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(name)
    If Not shp Is Nothing Then
        Dim i As Long
        With shp.ControlFormat
            For i = 1 To .ListCount
                If .Selected(i) Then result.Add CStr(.List(i))
            Next i
        End With
    End If
    On Error GoTo 0
    Set GetSelectedListItems = result
End Function

' リストボックスへ選択肢を再設定する（既存項目は一旦すべて削除してから登録し直す）。
Public Sub PopulateFormListBox(ByVal ws As Worksheet, ByVal name As String, ByVal items As Collection, _
                                Optional ByVal selectedIndex As Long = 0)
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.Shapes(name)
    If shp Is Nothing Then Exit Sub

    shp.ControlFormat.RemoveAllItems
    Dim it As Variant
    For Each it In items
        shp.ControlFormat.AddItem CStr(it)
    Next it
    If selectedIndex >= 1 And selectedIndex <= items.Count Then
        shp.ControlFormat.ListIndex = selectedIndex
    End If
    On Error GoTo 0
End Sub

Public Function GetCheckBoxValue(ByVal ws As Worksheet, ByVal name As String) As Boolean
    On Error Resume Next
    GetCheckBoxValue = (ws.Shapes(name).ControlFormat.Value = xlOn)
    On Error GoTo 0
End Function

Public Function GetOptionButtonValue(ByVal ws As Worksheet, ByVal name As String) As Boolean
    On Error Resume Next
    GetOptionButtonValue = (ws.Shapes(name).ControlFormat.Value = xlOn)
    On Error GoTo 0
End Function

'==============================================================================
' 補助
'==============================================================================

Private Function ColLetter(ByVal colIndex As Long) As String
    ColLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, colIndex).Address(True, False), "$")(0)
End Function

Private Function NzNum(ByVal v As Variant, ByVal defaultVal As Double) As Double
    If IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = defaultVal
    End If
End Function

' 文字列 Collection を昇順ソートして返す（年度・クラスコードの選択肢を見やすくするため）。
Private Function SortStrings(ByVal src As Collection) As Collection
    Dim n As Long
    n = src.Count
    Dim arr() As String
    If n = 0 Then
        Set SortStrings = New Collection
        Exit Function
    End If
    ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n
        arr(i) = CStr(src(i))
    Next i

    Dim j As Long
    Dim tmp As String
    For i = 1 To n - 1
        For j = 1 To n - i
            If arr(j) > arr(j + 1) Then
                tmp = arr(j): arr(j) = arr(j + 1): arr(j + 1) = tmp
            End If
        Next j
    Next i

    Dim result As New Collection
    For i = 1 To n
        result.Add arr(i)
    Next i
    Set SortStrings = result
End Function
