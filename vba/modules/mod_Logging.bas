Attribute VB_Name = "mod_Logging"
Option Explicit
'==============================================================================
' mod_Logging
' 目的 : 操作ログ・エラーログを T_Log テーブル（「ログ」シート）へ記録する。
'        長期間の運用で肥大化しないよう、既定の最大保持件数を超えた古い行を
'        自動的に間引く（容量対策）。
'==============================================================================

Private Const DEFAULT_LOG_MAX_ROWS As Long = 5000

' ログを1件追記する。
' level    : "INFO" / "WARN" / "ERROR"
' moduleName / procName : 発生元モジュール・処理名
' message  : 利用者にも分かる要約メッセージ
' detail   : 技術的な詳細（省略可）
Public Sub WriteLog(ByVal level As String, ByVal moduleName As String, _
                     ByVal procName As String, ByVal message As String, _
                     Optional ByVal detail As String = "")
    On Error GoTo ErrHandler

    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_LOG, mod_Common.TBL_LOG)
    If tbl Is Nothing Then Exit Sub   ' 初期化前（EnsureSchema未実行）は記録しない

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add

    newRow.Range(1, tbl.ListColumns(mod_Common.COL_LOG_DATETIME).Index).Value = Now
    newRow.Range(1, tbl.ListColumns(mod_Common.COL_LOG_LEVEL).Index).Value = level
    newRow.Range(1, tbl.ListColumns(mod_Common.COL_LOG_MODULE).Index).Value = moduleName
    newRow.Range(1, tbl.ListColumns(mod_Common.COL_LOG_PROC).Index).Value = procName
    newRow.Range(1, tbl.ListColumns(mod_Common.COL_LOG_MESSAGE).Index).Value = message
    newRow.Range(1, tbl.ListColumns(mod_Common.COL_LOG_DETAIL).Index).Value = detail

    TrimLogIfNeeded
    Exit Sub

ErrHandler:
    ' ログ機構自体のエラーは無限ループを避けるため画面表示のみに留める。
    Debug.Print "WriteLog failed: " & Err.Description
End Sub

' 保持上限（T_Settings!LOG_MAX_ROWS、既定5000件）を超えた場合、古い行から削除する。
Public Sub TrimLogIfNeeded()
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_LOG, mod_Common.TBL_LOG)
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    Dim maxRows As Long
    maxRows = mod_Settings.GetSettingLong(mod_Common.SETKEY_LOG_MAX_ROWS, DEFAULT_LOG_MAX_ROWS)
    If maxRows <= 0 Then maxRows = DEFAULT_LOG_MAX_ROWS

    Dim overflow As Long
    overflow = tbl.ListRows.Count - maxRows
    If overflow <= 0 Then Exit Sub

    Application.ScreenUpdating = False
    Dim i As Long
    For i = 1 To overflow
        tbl.ListRows(1).Delete   ' 先頭（最も古い）行を削除
    Next i
    Application.ScreenUpdating = True
End Sub

' ログを全件クリアする（「ログ」画面の「ログクリア」ボタンから呼び出す）。
' 呼び出し元（mod_UI）で利用者への確認ダイアログを表示した後に実行すること。
Public Sub ClearLog()
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_LOG, mod_Common.TBL_LOG)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
    WriteLog "INFO", "mod_Logging", "ClearLog", "ログを全件クリアしました。"
End Sub
