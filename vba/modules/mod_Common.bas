Attribute VB_Name = "mod_Common"
Option Explicit
'==============================================================================
' mod_Common
' 目的 : アプリケーション全体で共有する定数・汎用ヘルパー関数を集約する。
'        シート名／テーブル名／列名をここに一元管理することで、
'        レイアウト変更や仕様変更が発生しても修正箇所を最小限にする。
' 注意 : 他モジュールは原則としてシート名・テーブル名を直接文字列で
'        記述せず、必ず本モジュールの定数を参照すること。
'==============================================================================

'------------------------------------------------------------------
' アプリケーション基本情報
'------------------------------------------------------------------
Public Const APP_NAME As String = "評価データ分析ツール"
Public Const APP_VERSION As String = "1.0.0"

'------------------------------------------------------------------
' 画面（可視シート）名
'------------------------------------------------------------------
Public Const SH_HOME As String = "ホーム"
Public Const SH_IMPORT As String = "データ取込"
Public Const SH_ANALYSIS As String = "分析"
Public Const SH_SETTINGS As String = "設定"
Public Const SH_DATABASE As String = "データベース"
Public Const SH_LOG As String = "ログ"

'------------------------------------------------------------------
' データ保持シート名（非表示：xlSheetVeryHidden）
'------------------------------------------------------------------
Public Const SH_D_SCORES As String = "D_Scores"       ' 評価点数（縦持ち）
Public Const SH_D_TEMPLATES As String = "D_Templates" ' 列マッピングテンプレート

'------------------------------------------------------------------
' テーブル（ListObject）名
' T_Items / T_Settings は「設定」シート、T_Log は「ログ」シート、
' T_Summary は「データベース」シート上に配置し、可視シート自体を
' 記憶領域として利用する（別途コピーを持たないことで容量増加を抑制）。
'------------------------------------------------------------------
Public Const TBL_SCORES As String = "T_Scores"
Public Const TBL_ITEMS As String = "T_Items"
Public Const TBL_SUMMARY As String = "T_Summary"
Public Const TBL_LOG As String = "T_Log"
Public Const TBL_SETTINGS As String = "T_Settings"
Public Const TBL_TEMPLATES As String = "T_Templates"

'------------------------------------------------------------------
' T_Scores 列名（縦持ち：1行 = 1学生 × 1評価項目）
'------------------------------------------------------------------
Public Const COL_YEAR As String = "年度"
Public Const COL_CLASSCODE As String = "クラスコード"
Public Const COL_STUDENTNO As String = "学生番号"
Public Const COL_NAME As String = "氏名"
Public Const COL_ITEMCODE As String = "項目コード"
Public Const COL_SCORE As String = "点数"

'------------------------------------------------------------------
' T_Items 列名（評価項目マスタ）
'------------------------------------------------------------------
Public Const COL_ITEM_CODE As String = "項目コード"
Public Const COL_ITEM_DISPNAME As String = "表示名"
Public Const COL_ITEM_WEIGHT As String = "重み(%)"
Public Const COL_ITEM_ORDER As String = "表示順"
Public Const COL_ITEM_ACTIVE As String = "有効"

'------------------------------------------------------------------
' T_Summary 列名（学生ごとの総合評価キャッシュ）
'------------------------------------------------------------------
Public Const COL_SUM_YEAR As String = "年度"
Public Const COL_SUM_CLASSCODE As String = "クラスコード"
Public Const COL_SUM_STUDENTNO As String = "学生番号"
Public Const COL_SUM_NAME As String = "氏名"
Public Const COL_SUM_TOTAL As String = "総合評価"
Public Const COL_SUM_UPDATED As String = "更新日時"

'------------------------------------------------------------------
' T_Log 列名
'------------------------------------------------------------------
Public Const COL_LOG_DATETIME As String = "日時"
Public Const COL_LOG_LEVEL As String = "種別"
Public Const COL_LOG_MODULE As String = "モジュール"
Public Const COL_LOG_PROC As String = "処理"
Public Const COL_LOG_MESSAGE As String = "メッセージ"
Public Const COL_LOG_DETAIL As String = "詳細"

'------------------------------------------------------------------
' T_Settings 列名（キー・バリュー形式の設定値）
'------------------------------------------------------------------
Public Const COL_SET_KEY As String = "設定キー"
Public Const COL_SET_VALUE As String = "設定値"

' 既定設定キー
Public Const SETKEY_HIST_BIN_MODE As String = "HIST_BIN_MODE"     ' "COUNT" or "WIDTH"
Public Const SETKEY_HIST_BIN_COUNT As String = "HIST_BIN_COUNT"
Public Const SETKEY_HIST_BIN_WIDTH As String = "HIST_BIN_WIDTH"
Public Const SETKEY_SHOW_NORMAL As String = "GRAPH_SHOW_NORMAL"
Public Const SETKEY_OVERLAY_YEARS As String = "GRAPH_OVERLAY_YEARS"
Public Const SETKEY_LOG_MAX_ROWS As String = "LOG_MAX_ROWS"

'------------------------------------------------------------------
' T_Templates 列名（データ取込の列マッピングテンプレート）
'------------------------------------------------------------------
Public Const COL_TPL_NAME As String = "テンプレート名"
Public Const COL_TPL_SHEETPATTERN As String = "総合成績シート名パターン"
Public Const COL_TPL_MAPPING As String = "マッピング"
Public Const COL_TPL_SAVEDAT As String = "保存日時"

'------------------------------------------------------------------
' 画面レイアウト定数 - 「データ取込」シート
' mod_UI（コントロール生成）と mod_Import（値の読み書き）が共通で参照する。
' セル位置を変更する場合は、この定数のみを修正すればよい。
'------------------------------------------------------------------
Public Const IMP_CELL_TITLE As String = "B2"
Public Const IMP_CELL_STATUS As String = "B3"
Public Const IMP_RANGE_BTN_SELECTFILE As String = "B5:C5"
Public Const IMP_BTN_SELECTFILE As String = "btnImpSelectFile"
Public Const IMP_CELL_FILEPATH As String = "D5"          ' 結合 D5:H5
Public Const IMP_RANGE_LISTBOX_SHEETS As String = "B8:E14"
Public Const IMP_LISTBOX_SHEETS As String = "lstImpSheets"
Public Const IMP_CELL_YEAR_MANUAL As String = "G9"        ' 結合 G9:H9
Public Const IMP_CELL_CLASS_MANUAL As String = "G11"      ' 結合 G11:H11
Public Const IMP_RANGE_BTN_CONFIRMSHEET As String = "B15:C15"
Public Const IMP_BTN_CONFIRMSHEET As String = "btnImpConfirmSheet"
Public Const IMP_CELL_MAPPING_HINT As String = "B17"
Public Const IMP_MAPPING_HEADER_ROW As Long = 19
Public Const IMP_MAPPING_START_ROW As Long = 20
Public Const IMP_MAPPING_MAX_ROWS As Long = 30      ' 元データの列数上限（超える場合はこの値を増やす）
Public Const IMP_MAPPING_COL_HEADER As Long = 2      ' B列：元の見出し
Public Const IMP_MAPPING_COL_SAMPLE As Long = 3      ' C列：サンプル値
Public Const IMP_MAPPING_COL_TARGET As Long = 4      ' D列：マッピング先（ドロップダウン）
Public Const IMP_RANGE_BTN_PROCEED_CONFIRM As String = "B51:C51"
Public Const IMP_BTN_PROCEED_CONFIRM As String = "btnImpProceedConfirm"
Public Const IMP_RANGE_BTN_SAVETEMPLATE As String = "E51:G51"
Public Const IMP_BTN_SAVETEMPLATE As String = "btnImpSaveTemplate"
Public Const IMP_CELL_CONFIRM_SUMMARY As String = "B53"   ' 結合 B53:H58 相当（複数行テキスト）
Public Const IMP_CONFIRM_SUMMARY_ROWS As Long = 6
Public Const IMP_RANGE_CHK_OVERWRITE As String = "B60:D60"
Public Const IMP_CHK_OVERWRITE As String = "chkImpOverwrite"
Public Const IMP_RANGE_BTN_COMMIT As String = "B62:C62"
Public Const IMP_BTN_COMMIT As String = "btnImpCommit"
Public Const IMP_RANGE_BTN_CANCEL As String = "E62:F62"
Public Const IMP_BTN_CANCEL As String = "btnImpCancel"
Public Const IMP_MAPPING_HELPER_COL As Long = 20     ' T列：マッピング選択肢のヘルパー領域（非表示列）

' マッピング先の固定選択肢（学生属性）。評価項目はここに動的追加される。
Public Const MAP_TARGET_IGNORE As String = "（取り込まない）"
Public Const MAP_TARGET_YEAR As String = "年度"
Public Const MAP_TARGET_CLASSCODE As String = "クラスコード"
Public Const MAP_TARGET_STUDENTNO As String = "学生番号"
Public Const MAP_TARGET_NAME As String = "氏名"
Public Const MAP_TARGET_NEWITEM_PREFIX As String = "新規評価項目："

'------------------------------------------------------------------
' 画面レイアウト定数 - 「分析」シート
'------------------------------------------------------------------
Public Const ANL_CELL_TITLE As String = "B2"
Public Const ANL_RANGE_LISTBOX_YEARS As String = "B5:C12"
Public Const ANL_LISTBOX_YEARS As String = "lstAnlYears"
Public Const ANL_RANGE_LISTBOX_CLASSES As String = "D5:E12"
Public Const ANL_LISTBOX_CLASSES As String = "lstAnlClasses"
Public Const ANL_RANGE_LISTBOX_ITEM As String = "F5:G12"
Public Const ANL_LISTBOX_ITEM As String = "lstAnlItem"
Public Const ANL_RANGE_CHK_NORMAL As String = "I5:K5"
Public Const ANL_CHK_NORMAL As String = "chkAnlNormal"
Public Const ANL_RANGE_CHK_OVERLAY As String = "I6:K6"
Public Const ANL_CHK_OVERLAY As String = "chkAnlOverlay"
Public Const ANL_RANGE_OPT_BINCOUNT As String = "I8:J8"
Public Const ANL_OPT_BINCOUNT As String = "optAnlBinCount"
Public Const ANL_CELL_BINCOUNT As String = "K8"
Public Const ANL_RANGE_OPT_BINWIDTH As String = "I9:J9"
Public Const ANL_OPT_BINWIDTH As String = "optAnlBinWidth"
Public Const ANL_CELL_BINWIDTH As String = "K9"
Public Const ANL_RANGE_BTN_RUN As String = "B14:C14"
Public Const ANL_BTN_RUN As String = "btnAnlRun"
Public Const ANL_RANGE_CHART_HIST As String = "B16:H26"
Public Const ANL_CHART_HIST As String = "chtAnlHistogram"
Public Const ANL_RANGE_CHART_CUM As String = "B27:H37"
Public Const ANL_CHART_CUM As String = "chtAnlCumulative"
Public Const ANL_RANGE_CHART_BOX As String = "B38:H48"
Public Const ANL_CHART_BOX As String = "chtAnlBoxplot"
Public Const ANL_STATS_START As String = "B50"
Public Const ANL_STUDENTLIST_TITLE As String = "E49"
Public Const ANL_STUDENTLIST_HEADER_ROW As Long = 50
Public Const ANL_STUDENTLIST_START_ROW As Long = 51
Public Const ANL_STUDENTLIST_COL As Long = 5   ' E列
Public Const ANL_STUDENTLIST_MAX_ROWS As Long = 260  ' 約200名クラスを想定した表示上限

'------------------------------------------------------------------
' 画面レイアウト定数 - 「設定」シート
'------------------------------------------------------------------
' T_Items テーブルは B4 起点（EnsureTableItems）。項目数増加に備え、
' 右側の制御群は十分な余白を空けた I列以降に配置する。
Public Const SET_RANGE_BTN_APPLY As String = "B30:D30"
Public Const SET_BTN_APPLY As String = "btnSetApply"
' T_Settings テーブルは I4 起点（EnsureTableSettings）。
Public Const SET_RANGE_LISTBOX_TEMPLATES As String = "I13:J20"
Public Const SET_LISTBOX_TEMPLATES As String = "lstSetTemplates"
Public Const SET_RANGE_BTN_DELETETEMPLATE As String = "I22:J22"
Public Const SET_BTN_DELETETEMPLATE As String = "btnSetDeleteTemplate"

'------------------------------------------------------------------
' 画面レイアウト定数 - 「データベース」シート
' T_Summary テーブルは B4 起点（EnsureTableSummary）。
'------------------------------------------------------------------
Public Const DB_CELL_DELETE_YEAR_LABEL As String = "I4"
Public Const DB_CELL_DELETE_YEAR As String = "J4"
Public Const DB_CELL_DELETE_CLASS_LABEL As String = "I5"
Public Const DB_CELL_DELETE_CLASS As String = "J5"
Public Const DB_RANGE_BTN_DELETE As String = "I6:J6"
Public Const DB_BTN_DELETE As String = "btnDbDeleteGroup"
Public Const DB_RANGE_BTN_REFRESH As String = "I8:J8"
Public Const DB_BTN_REFRESH As String = "btnDbRefresh"

'------------------------------------------------------------------
' 画面レイアウト定数 - 「ログ」シート
' T_Log テーブルは B4 起点（EnsureTableLog）。
'------------------------------------------------------------------
Public Const LOG_RANGE_BTN_CLEAR As String = "I4:J4"
Public Const LOG_BTN_CLEAR As String = "btnLogClear"

'------------------------------------------------------------------
' 画面レイアウト定数 - 「ホーム」シート
'------------------------------------------------------------------
Public Const HOME_RANGE_BTN_IMPORT As String = "B5:D6"
Public Const HOME_BTN_IMPORT As String = "btnHomeImport"
Public Const HOME_RANGE_BTN_ANALYSIS As String = "B8:D9"
Public Const HOME_BTN_ANALYSIS As String = "btnHomeAnalysis"
Public Const HOME_RANGE_BTN_SETTINGS As String = "B11:D12"
Public Const HOME_BTN_SETTINGS As String = "btnHomeSettings"
Public Const HOME_RANGE_BTN_DATABASE As String = "B14:D15"
Public Const HOME_BTN_DATABASE As String = "btnHomeDatabase"
Public Const HOME_RANGE_BTN_LOG As String = "B17:D18"
Public Const HOME_BTN_LOG As String = "btnHomeLog"
Public Const HOME_CELL_SUMMARY As String = "F5"   ' 結合 F5:J10 相当

'------------------------------------------------------------------
' ナビゲーションバー（全画面共通・1行目に配置。2行目がタイトル行）
'------------------------------------------------------------------
Public Const NAV_RANGE_HOME As String = "B1:C1"
Public Const NAV_BTN_HOME As String = "btnNavHome"
Public Const NAV_RANGE_IMPORT As String = "D1:E1"
Public Const NAV_BTN_IMPORT As String = "btnNavImport"
Public Const NAV_RANGE_ANALYSIS As String = "F1:G1"
Public Const NAV_BTN_ANALYSIS As String = "btnNavAnalysis"
Public Const NAV_RANGE_SETTINGS As String = "H1:I1"
Public Const NAV_BTN_SETTINGS As String = "btnNavSettings"
Public Const NAV_RANGE_DATABASE As String = "J1:K1"
Public Const NAV_BTN_DATABASE As String = "btnNavDatabase"
Public Const NAV_RANGE_LOG As String = "L1:M1"
Public Const NAV_BTN_LOG As String = "btnNavLog"

'==============================================================================
' 汎用ヘルパー関数
'==============================================================================

' 指定名のワークシートを安全に取得する。存在しない場合は Nothing を返す。
Public Function GetSheetSafe(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set GetSheetSafe = ws
End Function

' 指定シート上の指定テーブル（ListObject）を取得する。存在しない場合は Nothing。
Public Function GetTableSafe(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = GetSheetSafe(sheetName)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set GetTableSafe = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

' テーブルのデータ行数（見出しを除く）を返す。
Public Function TableRowCount(ByVal tbl As ListObject) As Long
    If tbl Is Nothing Then Exit Function
    If tbl.ListRows.Count = 0 Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    TableRowCount = tbl.ListRows.Count
End Function

' 文字列比較用の正規化処理。
' ・前後空白の除去
' ・全角英数記号を半角に変換（元データのレイアウト揺れを吸収する）
' ・大文字小文字を無視できるよう小文字化
Public Function NormalizeText(ByVal s As Variant) As String
    Dim t As String
    If IsError(s) Then
        NormalizeText = ""
        Exit Function
    End If
    t = CStr(s)
    t = Trim$(t)
    t = StrConv(t, vbNarrow)   ' 全角→半角（英数記号・カナ）
    t = LCase$(t)
    NormalizeText = t
End Function

' 2つの見出し文字列がどの程度類似しているかを簡易判定する。
' 完全一致 = 100、正規化後一致 = 90、部分一致 = 60、不一致 = 0
Public Function HeaderSimilarity(ByVal headerText As Variant, ByVal candidateWord As String) As Long
    Dim a As String, b As String
    a = NormalizeText(headerText)
    b = NormalizeText(candidateWord)
    If Len(a) = 0 Or Len(b) = 0 Then Exit Function
    If a = b Then
        HeaderSimilarity = 100
    ElseIf InStr(1, a, b, vbTextCompare) > 0 Or InStr(1, b, a, vbTextCompare) > 0 Then
        HeaderSimilarity = 60
    Else
        HeaderSimilarity = 0
    End If
End Function

' 指定列の最終行番号を返す（見出し行を含めた実際の行番号）。
Public Function LastRowInColumn(ByVal ws As Worksheet, ByVal col As Long) As Long
    LastRowInColumn = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' クラスコードを4桁ゼロ埋めの文字列として整形する。
Public Function FormatClassCode(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        FormatClassCode = ""
        Exit Function
    End If
    If IsNumeric(s) Then
        FormatClassCode = Format$(CLng(s), "0000")
    Else
        FormatClassCode = s
    End If
End Function

' 現在時刻を "yyyy/mm/dd hh:nn:ss" 形式で返す（ログ・更新日時記録用）。
Public Function NowText() As String
    NowText = Format$(Now, "yyyy/mm/dd hh:nn:ss")
End Function

' 集計キー（年度+クラスコード+学生番号+項目コード等）を生成する。
' Collection のキーとして使用し、重複判定を高速化するために用いる。
Public Function MakeKey(ParamArray parts() As Variant) As String
    Dim i As Long
    Dim s As String
    For i = LBound(parts) To UBound(parts)
        s = s & CStr(parts(i)) & vbTab
    Next i
    MakeKey = s
End Function

' Collection に指定キーが存在するかを判定する（存在確認専用の軽量関数）。
Public Function CollectionHasKey(ByVal col As Collection, ByVal key As String) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = col.Item(key)
    CollectionHasKey = (Err.Number = 0)
    On Error GoTo 0
End Function

' 共通エラーハンドラ。エラー内容をログへ記録し、必要に応じて利用者へ通知する。
' 各プロシージャは On Error GoTo ErrHandler の後、以下のように呼び出す。
'   ErrHandler:
'       mod_Common.HandleError "mod_Xxx", "ProcName"
Public Sub HandleError(ByVal moduleName As String, ByVal procName As String, _
                        Optional ByVal showMessage As Boolean = True)
    Dim msg As String
    Dim detail As String
    msg = Err.Description
    detail = "ErrNo=" & Err.Number
    mod_Logging.WriteLog "ERROR", moduleName, procName, msg, detail
    If showMessage Then
        MsgBox APP_NAME & " でエラーが発生しました。" & vbCrLf & _
               "処理: " & moduleName & "." & procName & vbCrLf & _
               "内容: " & msg, vbExclamation, APP_NAME
    End If
End Sub

' 画面更新・警告表示・イベントを一時停止する（重い処理の前後で使用）。
Public Sub BeginBusy()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Cursor = xlWait
    Application.StatusBar = APP_NAME & " 処理中..."
End Sub

Public Sub EndBusy()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
End Sub
