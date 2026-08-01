Attribute VB_Name = "mod_Database"
Option Explicit
'==============================================================================
' mod_Database
' 目的 : ワークブック内部データ（評価点数・集計・マスタ）に対する
'        スキーマ初期化（EnsureSchema）と CRUD 処理を提供する。
'
' 記憶領域の設計方針:
'   ・T_Scores    … 縦持ち（1行=1学生×1評価項目）。非表示シート D_Scores。
'                    項目数が増減してもテーブル構造の変更は不要。
'   ・T_Items     … 評価項目マスタ（内部名・表示名・重み等）。「設定」シート。
'   ・T_Summary   … 学生ごとの総合評価キャッシュ。「データベース」シート。
'                    T_Scores から重み付き平均を都度再計算して更新する。
'   ・T_Log       … 操作・エラーログ。「ログ」シート。
'   ・T_Settings  … キー・バリュー形式の各種設定値。「設定」シート。
'   ・T_Templates … 列マッピングテンプレート。非表示シート D_Templates。
'
' 総合評価の算出式:
'   総合評価 = Σ(点数i × 重みi) / Σ(重みi)   （その学生が保有する項目のみ対象）
'   重みは 0～100 の相対値として扱い、必ずしも合計100%である必要はない。
'==============================================================================

'==============================================================================
' スキーマ初期化
'==============================================================================

' 全シート・全テーブルの存在を確認し、不足していれば作成する（べき等）。
' ブック起動時（Workbook_Open）および保守用マクロから呼び出される。
Public Sub EnsureSchema()
    EnsureVisibleSheets
    EnsureHiddenSheets
    EnsureTableScores
    EnsureTableItems
    EnsureTableSummary
    EnsureTableLog
    EnsureTableSettings
    EnsureTableTemplates
    EnsureDefaultSettings
End Sub

Private Sub EnsureVisibleSheets()
    Dim names As Variant
    names = Array(mod_Common.SH_HOME, mod_Common.SH_IMPORT, mod_Common.SH_ANALYSIS, _
                   mod_Common.SH_SETTINGS, mod_Common.SH_DATABASE, mod_Common.SH_LOG)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        EnsureSheetExists CStr(names(i)), xlSheetVisible
    Next i
End Sub

Private Sub EnsureHiddenSheets()
    EnsureSheetExists mod_Common.SH_D_SCORES, xlSheetVeryHidden
    EnsureSheetExists mod_Common.SH_D_TEMPLATES, xlSheetVeryHidden
End Sub

Private Sub EnsureSheetExists(ByVal sheetName As String, ByVal visibility As XlSheetVisibility)
    Dim ws As Worksheet
    Set ws = mod_Common.GetSheetSafe(sheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    ws.Visible = visibility
End Sub

' 指定シート上にテーブルが無ければヘッダー行付きで新規作成する。
Private Function EnsureTable(ByVal sheetName As String, ByVal tableName As String, _
                              ByVal headers As Variant, ByVal startCellAddr As String) As ListObject
    Dim ws As Worksheet
    Set ws = mod_Common.GetSheetSafe(sheetName)
    If ws Is Nothing Then Exit Function

    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(sheetName, tableName)
    If Not tbl Is Nothing Then
        Set EnsureTable = tbl
        Exit Function
    End If

    Dim startCell As Range
    Set startCell = ws.Range(startCellAddr)
    Dim endCell As Range
    Set endCell = startCell.Offset(0, UBound(headers) - LBound(headers))

    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, endCell), , xlYes)
    tbl.Name = tableName

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        tbl.HeaderRowRange.Cells(1, i - LBound(headers) + 1).Value = headers(i)
    Next i

    Set EnsureTable = tbl
End Function

Private Sub EnsureTableScores()
    EnsureTable mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES, _
        Array(mod_Common.COL_YEAR, mod_Common.COL_CLASSCODE, mod_Common.COL_STUDENTNO, _
              mod_Common.COL_NAME, mod_Common.COL_ITEMCODE, mod_Common.COL_SCORE), "A1"
End Sub

Private Sub EnsureTableItems()
    EnsureTable mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS, _
        Array(mod_Common.COL_ITEM_CODE, mod_Common.COL_ITEM_DISPNAME, mod_Common.COL_ITEM_WEIGHT, _
              mod_Common.COL_ITEM_ORDER, mod_Common.COL_ITEM_ACTIVE), "B4"
End Sub

Private Sub EnsureTableSummary()
    EnsureTable mod_Common.SH_DATABASE, mod_Common.TBL_SUMMARY, _
        Array(mod_Common.COL_SUM_YEAR, mod_Common.COL_SUM_CLASSCODE, mod_Common.COL_SUM_STUDENTNO, _
              mod_Common.COL_SUM_NAME, mod_Common.COL_SUM_TOTAL, mod_Common.COL_SUM_UPDATED), "B4"
End Sub

Private Sub EnsureTableLog()
    EnsureTable mod_Common.SH_LOG, mod_Common.TBL_LOG, _
        Array(mod_Common.COL_LOG_DATETIME, mod_Common.COL_LOG_LEVEL, mod_Common.COL_LOG_MODULE, _
              mod_Common.COL_LOG_PROC, mod_Common.COL_LOG_MESSAGE, mod_Common.COL_LOG_DETAIL), "B4"
End Sub

Private Sub EnsureTableSettings()
    EnsureTable mod_Common.SH_SETTINGS, mod_Common.TBL_SETTINGS, _
        Array(mod_Common.COL_SET_KEY, mod_Common.COL_SET_VALUE), "I4"
End Sub

Private Sub EnsureTableTemplates()
    EnsureTable mod_Common.SH_D_TEMPLATES, mod_Common.TBL_TEMPLATES, _
        Array(mod_Common.COL_TPL_NAME, mod_Common.COL_TPL_SHEETPATTERN, _
              mod_Common.COL_TPL_MAPPING, mod_Common.COL_TPL_SAVEDAT), "A1"
End Sub

' 初回起動時のみ既定値を投入する（既に値があれば上書きしない）。
Private Sub EnsureDefaultSettings()
    SetIfMissing mod_Common.SETKEY_HIST_BIN_MODE, "COUNT"
    SetIfMissing mod_Common.SETKEY_HIST_BIN_COUNT, "10"
    SetIfMissing mod_Common.SETKEY_HIST_BIN_WIDTH, "10"
    SetIfMissing mod_Common.SETKEY_SHOW_NORMAL, "TRUE"
    SetIfMissing mod_Common.SETKEY_OVERLAY_YEARS, "TRUE"
    SetIfMissing mod_Common.SETKEY_LOG_MAX_ROWS, "5000"
End Sub

Private Sub SetIfMissing(ByVal key As String, ByVal value As String)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_SETTINGS)
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then
        Dim r As ListRow
        For Each r In tbl.ListRows
            If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SET_KEY).Index).Value) = key Then Exit Sub
        Next r
    End If
    mod_Settings.SetSetting key, value
End Sub

'==============================================================================
' T_Scores 操作
'==============================================================================

' 重複判定用のキー集合（年度|期別|学生番号|項目コード）を構築する。
' インポート時に1行ずつテーブル全体を走査すると件数増加で低速化するため、
' 事前に Collection へキーを読み込み O(1) 判定にする。
Public Function BuildScoreKeyIndex() As Collection
    Dim idx As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then GoTo Done
    If tbl.DataBodyRange Is Nothing Then GoTo Done

    Dim data As Variant
    data = tbl.DataBodyRange.Value  ' 1回の読み込みで配列化（高速化）
    Dim cYear As Long, cClass As Long, cStudent As Long, cItem As Long
    cYear = tbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_CLASSCODE).Index
    cStudent = tbl.ListColumns(mod_Common.COL_STUDENTNO).Index
    cItem = tbl.ListColumns(mod_Common.COL_ITEMCODE).Index

    Dim i As Long
    Dim key As String
    On Error Resume Next
    For i = 1 To UBound(data, 1)
        key = mod_Common.MakeKey(data(i, cYear), data(i, cClass), data(i, cStudent), data(i, cItem))
        idx.Add True, key
    Next i
    On Error GoTo 0

Done:
    Set BuildScoreKeyIndex = idx
End Function

' BuildScoreKeyIndex と同様だが、値の代わりに ListRow 参照を保持する。
' 上書き取込（既存データの点数を直接更新するモード）で使用する。
Public Function BuildScoreRowIndex() As Collection
    Dim idx As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then GoTo Done
    If tbl.DataBodyRange Is Nothing Then GoTo Done

    Dim cYear As Long, cClass As Long, cStudent As Long, cItem As Long
    cYear = tbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_CLASSCODE).Index
    cStudent = tbl.ListColumns(mod_Common.COL_STUDENTNO).Index
    cItem = tbl.ListColumns(mod_Common.COL_ITEMCODE).Index

    Dim i As Long, key As String
    On Error Resume Next
    For i = 1 To tbl.ListRows.Count
        Dim rw As ListRow
        Set rw = tbl.ListRows(i)
        key = mod_Common.MakeKey(rw.Range(1, cYear).Value, rw.Range(1, cClass).Value, _
                                  rw.Range(1, cStudent).Value, rw.Range(1, cItem).Value)
        idx.Add rw, key
    Next i
    On Error GoTo 0

Done:
    Set BuildScoreRowIndex = idx
End Function

' 既存の1行（ListRow）の点数を更新する（上書き取込で使用）。
Public Sub UpdateScoreValue(ByVal row As ListRow, ByVal newScore As Double)
    Dim tbl As ListObject
    Set tbl = row.Parent
    row.Range(1, tbl.ListColumns(mod_Common.COL_SCORE).Index).Value = newScore
End Sub

' 評価点数データを一括追加する。rows は 1行あたり
' (年度, 期別, 学生番号, 氏名, 項目コード, 点数) の配列。
' 事前にワークシート関数を使わず配列操作で書き込むことで大量データでも高速。
Public Sub AppendScoreRows(ByVal rows As Collection)
    If rows.Count = 0 Then Exit Sub

    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = tbl.Parent

    Dim startRow As Long
    If tbl.DataBodyRange Is Nothing Then
        startRow = tbl.HeaderRowRange.Row + 1
    Else
        startRow = tbl.DataBodyRange.Row + tbl.DataBodyRange.Rows.Count
    End If

    Dim outArr() As Variant
    ReDim outArr(1 To rows.Count, 1 To 6)
    Dim i As Long, r As Variant
    i = 1
    For Each r In rows
        outArr(i, 1) = r(0): outArr(i, 2) = r(1): outArr(i, 3) = r(2)
        outArr(i, 4) = r(3): outArr(i, 5) = r(4): outArr(i, 6) = r(5)
        i = i + 1
    Next r

    Dim destRange As Range
    Set destRange = ws.Range(ws.Cells(startRow, tbl.Range.Column), _
                              ws.Cells(startRow + rows.Count - 1, tbl.Range.Column + 5))
    destRange.Value = outArr
    tbl.Resize ws.Range(tbl.HeaderRowRange.Cells(1, 1), destRange.Cells(destRange.Rows.Count, destRange.Columns.Count))
End Sub

'==============================================================================
' T_Summary 操作（総合評価の再計算）
'==============================================================================

' T_Scores × T_Items(重み) から総合評価を再計算し、T_Summary を更新する。
' year / classCode を指定すると対象を絞り込み、省略時は全件再計算する。
Public Sub RecalcSummary(Optional ByVal year As Variant, Optional ByVal classCode As Variant)
    Dim scoresTbl As ListObject, itemsTbl As ListObject, sumTbl As ListObject
    Set scoresTbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    Set itemsTbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS)
    Set sumTbl = mod_Common.GetTableSafe(mod_Common.SH_DATABASE, mod_Common.TBL_SUMMARY)
    If scoresTbl Is Nothing Or sumTbl Is Nothing Then Exit Sub
    If scoresTbl.DataBodyRange Is Nothing Then Exit Sub

    ' 項目コード→重み の対応表を Collection で構築
    Dim weightIdx As New Collection
    If Not itemsTbl Is Nothing Then
        If Not itemsTbl.DataBodyRange Is Nothing Then
            Dim itemData As Variant
            itemData = itemsTbl.DataBodyRange.Value
            Dim ci As Long
            For ci = 1 To UBound(itemData, 1)
                weightIdx.Add CDbl(itemData(ci, itemsTbl.ListColumns(mod_Common.COL_ITEM_WEIGHT).Index)), _
                              CStr(itemData(ci, itemsTbl.ListColumns(mod_Common.COL_ITEM_CODE).Index))
            Next ci
        End If
    End If

    Dim data As Variant
    data = scoresTbl.DataBodyRange.Value
    Dim cYear As Long, cClass As Long, cStudent As Long, cName As Long, cItem As Long, cScore As Long
    cYear = scoresTbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = scoresTbl.ListColumns(mod_Common.COL_CLASSCODE).Index
    cStudent = scoresTbl.ListColumns(mod_Common.COL_STUDENTNO).Index
    cName = scoresTbl.ListColumns(mod_Common.COL_NAME).Index
    cItem = scoresTbl.ListColumns(mod_Common.COL_ITEMCODE).Index
    cScore = scoresTbl.ListColumns(mod_Common.COL_SCORE).Index

    ' 学生キー → (氏名, 加重合計, 重み合計) を集計
    Dim names As New Collection, wsum As New Collection, ssum As New Collection, keys As New Collection
    Dim i As Long, key As String, w As Double
    For i = 1 To UBound(data, 1)
        If Not IsMissing(year) Then
            If Len(CStr(year)) > 0 Then
                If CStr(data(i, cYear)) <> CStr(year) Then GoTo ContinueLoop
            End If
        End If
        If Not IsMissing(classCode) Then
            If Len(CStr(classCode)) > 0 Then
                If CStr(data(i, cClass)) <> CStr(classCode) Then GoTo ContinueLoop
            End If
        End If

        key = mod_Common.MakeKey(data(i, cYear), data(i, cClass), data(i, cStudent))
        w = 100
        On Error Resume Next
        w = weightIdx.Item(CStr(data(i, cItem)))
        On Error GoTo 0

        If Not mod_Common.CollectionHasKey(wsum, key) Then
            wsum.Add 0#, key
            ssum.Add 0#, key
            names.Add CStr(data(i, cName)), key
            keys.Add key, key
        End If
        ReplaceAccum wsum, key, CDbl(data(i, cScore)) * w
        ReplaceAccum ssum, key, w
ContinueLoop:
    Next i

    ' T_Summary へ Upsert
    Dim k As Variant
    For Each k In keys
        Dim parts() As String
        parts = Split(CStr(k), vbTab)
        Dim total As Double
        If ssum.Item(CStr(k)) > 0 Then
            total = wsum.Item(CStr(k)) / ssum.Item(CStr(k))
        Else
            total = 0
        End If
        UpsertSummaryRow parts(0), parts(1), parts(2), names.Item(CStr(k)), total
    Next k
End Sub

' Collection は値の直接更新ができないため、削除→再追加で累積加算を模倣する。
Private Sub ReplaceAccum(ByRef col As Collection, ByVal key As String, ByVal addValue As Double)
    Dim cur As Double
    cur = col.Item(key)
    col.Remove key
    col.Add cur + addValue, key
End Sub

' T_Summary に1件 Upsert する。
Private Sub UpsertSummaryRow(ByVal year As String, ByVal classCode As String, ByVal studentNo As String, _
                              ByVal studentName As String, ByVal total As Double)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_DATABASE, mod_Common.TBL_SUMMARY)
    If tbl Is Nothing Then Exit Sub

    Dim r As ListRow
    Dim found As Boolean
    If Not tbl.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To tbl.ListRows.Count
            Set r = tbl.ListRows(i)
            If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_YEAR).Index).Value) = year And _
               CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_CLASSCODE).Index).Value) = classCode And _
               CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_STUDENTNO).Index).Value) = studentNo Then
                found = True
                Exit For
            End If
        Next i
    End If
    If Not found Then Set r = tbl.ListRows.Add

    r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_YEAR).Index).Value = year
    r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_CLASSCODE).Index).Value = classCode
    r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_STUDENTNO).Index).Value = studentNo
    r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_NAME).Index).Value = studentName
    r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_TOTAL).Index).Value = Application.Round(total, 2)
    r.Range(1, tbl.ListColumns(mod_Common.COL_SUM_UPDATED).Index).Value = mod_Common.NowText()
End Sub

'==============================================================================
' 参照系（分析画面の選択肢生成に使用）
'==============================================================================

Public Function GetDistinctYears() As Collection
    Set GetDistinctYears = GetDistinctColumnValues(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES, mod_Common.COL_YEAR)
End Function

Public Function GetDistinctClasses(Optional ByVal year As String = "") As Collection
    Dim result As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then GoTo Done
    If tbl.DataBodyRange Is Nothing Then GoTo Done

    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim cYear As Long, cClass As Long
    cYear = tbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_CLASSCODE).Index

    Dim i As Long
    For i = 1 To UBound(data, 1)
        If Len(year) = 0 Or CStr(data(i, cYear)) = year Then
            Dim v As String
            v = CStr(data(i, cClass))
            If Not mod_Common.CollectionHasKey(result, v) Then result.Add v, v
        End If
    Next i
Done:
    Set GetDistinctClasses = result
End Function

Public Function GetActiveItemCodes() As Collection
    Dim result As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            Dim r As ListRow
            For Each r In tbl.ListRows
                If CBool(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_ACTIVE).Index).Value) Then
                    result.Add CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_CODE).Index).Value)
                End If
            Next r
        End If
    End If
    Set GetActiveItemCodes = result
End Function

Private Function GetDistinctColumnValues(ByVal sheetName As String, ByVal tableName As String, ByVal colName As String) As Collection
    Dim result As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(sheetName, tableName)
    If tbl Is Nothing Then GoTo Done
    If tbl.DataBodyRange Is Nothing Then GoTo Done

    Dim colIdx As Long
    colIdx = tbl.ListColumns(colName).Index
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim i As Long, v As String
    For i = 1 To UBound(data, 1)
        v = CStr(data(i, colIdx))
        If Not mod_Common.CollectionHasKey(result, v) Then result.Add v, v
    Next i
Done:
    Set GetDistinctColumnValues = result
End Function

'==============================================================================
' 容量対策（管理者向け）
'==============================================================================

' 指定した年度・期別のデータを完全に削除する（管理者機能）。
' 誤操作防止のため呼び出し元（mod_UI）で確認ダイアログを表示すること。
Public Sub DeleteByYearClass(ByVal year As String, ByVal classCode As String)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim cYear As Long, cClass As Long
    cYear = tbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_CLASSCODE).Index

    Dim i As Long
    Dim deletedCount As Long
    For i = tbl.ListRows.Count To 1 Step -1
        If CStr(tbl.ListRows(i).Range(1, cYear).Value) = year And _
           CStr(tbl.ListRows(i).Range(1, cClass).Value) = classCode Then
            tbl.ListRows(i).Delete
            deletedCount = deletedCount + 1
        End If
    Next i

    DeleteSummaryByYearClass year, classCode
    mod_Logging.WriteLog "INFO", "mod_Database", "DeleteByYearClass", _
        "データを削除しました。年度=" & year & " 期別=" & classCode & " 件数=" & deletedCount
End Sub

Private Sub DeleteSummaryByYearClass(ByVal year As String, ByVal classCode As String)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_DATABASE, mod_Common.TBL_SUMMARY)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim cYear As Long, cClass As Long
    cYear = tbl.ListColumns(mod_Common.COL_SUM_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_SUM_CLASSCODE).Index

    Dim i As Long
    For i = tbl.ListRows.Count To 1 Step -1
        If CStr(tbl.ListRows(i).Range(1, cYear).Value) = year And _
           CStr(tbl.ListRows(i).Range(1, cClass).Value) = classCode Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub
