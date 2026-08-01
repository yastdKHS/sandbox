Attribute VB_Name = "mod_Analysis"
Option Explicit
'==============================================================================
' mod_Analysis
' 目的 : 「分析」画面の中心的な処理を担当する。
'        利用者が選択した年度・期別・評価項目に基づき T_Scores から
'        該当データを抽出し、mod_Statistics で統計量を計算、mod_Graph で
'        グラフを更新し、mod_UI へ表示用データを渡す（司令塔の役割）。
'==============================================================================

' 「分析実行」ボタンから呼び出されるエントリポイント。
Public Sub RunAnalysis()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_ANALYSIS)

    Dim years As Collection, classes As Collection, itemDisp As String, itemCode As String
    Set years = mod_UI.GetSelectedListItems(ws, mod_Common.ANL_LISTBOX_YEARS)
    Set classes = mod_UI.GetSelectedListItems(ws, mod_Common.ANL_LISTBOX_CLASSES)
    itemDisp = mod_UI.GetFormListBoxSelection(ws, mod_Common.ANL_LISTBOX_ITEM)

    If years.Count = 0 Or classes.Count = 0 Or Len(itemDisp) = 0 Then
        MsgBox "年度・期別・評価項目をそれぞれ1つ以上選択してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    itemCode = mod_Settings.FindItemCodeByDisplayName(itemDisp)
    If Len(itemCode) = 0 Then itemCode = itemDisp

    mod_Common.BeginBusy

    Dim flat() As Double
    flat = BuildScoresFlat(years, classes, itemCode)

    If mod_Statistics.ArrLen(flat) = 0 Then
        mod_Common.EndBusy
        MsgBox "選択した条件に該当するデータがありません。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    Dim yearLabels As Collection, yearArrays As Collection
    BuildScoresGroupedByYear years, classes, itemCode, yearLabels, yearArrays

    ' 統計量の計算と表示
    Dim statLabels() As String, statValues() As Variant
    BuildStatsTable flat, statLabels, statValues
    mod_UI.RenderStatsTable statLabels, statValues

    ' 個人別一覧（偏差値付き）の計算と表示
    Dim studentNos() As String, studentNames() As String, scoreVals() As Double, devVals() As Double
    BuildStudentDetailList years, classes, itemCode, studentNos, studentNames, scoreVals, devVals
    mod_UI.RenderStudentDetailList studentNos, studentNames, scoreVals, devVals

    ' グラフの更新
    Dim binMode As String, binCount As Long, binWidth As Double
    Dim showNormal As Boolean, overlayYears As Boolean
    mod_UI.GetAnalysisGraphOptions binMode, binCount, binWidth, showNormal, overlayYears

    mod_Graph.UpdateHistogramChart yearLabels, yearArrays, overlayYears, binMode, binCount, binWidth, showNormal, itemDisp
    mod_Graph.UpdateCumulativeChart yearLabels, yearArrays, overlayYears, itemDisp
    mod_Graph.UpdateBoxPlotChart yearLabels, yearArrays, overlayYears, itemDisp

    mod_Common.EndBusy

    mod_Logging.WriteLog "INFO", "mod_Analysis", "RunAnalysis", _
        "分析を実行しました。項目=" & itemDisp & " 年度数=" & years.Count & " 期別数=" & classes.Count & _
        " データ件数=" & mod_Statistics.ArrLen(flat)
    Exit Sub

ErrHandler:
    mod_Common.EndBusy
    mod_Common.HandleError "mod_Analysis", "RunAnalysis"
End Sub

'==============================================================================
' データ抽出
'==============================================================================

' 選択条件（複数年度・複数期別・単一評価項目）に一致する点数を1本の配列として返す。
Public Function BuildScoresFlat(ByVal years As Collection, ByVal classes As Collection, ByVal itemCode As String) As Double()
    Dim result() As Double
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then GoTo Done
    If tbl.DataBodyRange Is Nothing Then GoTo Done

    Dim yearSet As Collection, classSet As Collection
    Set yearSet = ToLookupSet(years)
    Set classSet = ToLookupSet(classes)

    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim cYear As Long, cClass As Long, cItem As Long, cScore As Long
    cYear = tbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_CLASSCODE).Index
    cItem = tbl.ListColumns(mod_Common.COL_ITEMCODE).Index
    cScore = tbl.ListColumns(mod_Common.COL_SCORE).Index

    ReDim result(1 To UBound(data, 1))
    Dim n As Long, i As Long
    n = 0
    For i = 1 To UBound(data, 1)
        If CStr(data(i, cItem)) = itemCode And _
           mod_Common.CollectionHasKey(yearSet, CStr(data(i, cYear))) And _
           mod_Common.CollectionHasKey(classSet, CStr(data(i, cClass))) Then
            n = n + 1
            result(n) = CDbl(data(i, cScore))
        End If
    Next i

    If n = 0 Then
        Erase result
    Else
        ReDim Preserve result(1 To n)
    End If

Done:
    BuildScoresFlat = result
End Function

' 年度ごとに分けた点数配列群を返す（重畳表示用）。
' yearLabels(i) と yearArrays(i) が対になるよう、同じ並び順の2つの Collection として返す。
' （VBA の標準 Collection はキーの列挙ができないため、位置対応の2本立てとしている）
Public Sub BuildScoresGroupedByYear(ByVal years As Collection, ByVal classes As Collection, ByVal itemCode As String, _
                                     ByRef yearLabels As Collection, ByRef yearArrays As Collection)
    Set yearLabels = New Collection
    Set yearArrays = New Collection

    Dim y As Variant
    Dim singleYearSet As Collection
    For Each y In years
        Set singleYearSet = New Collection
        singleYearSet.Add CStr(y), CStr(y)
        Dim arr() As Double
        arr = BuildScoresFlat(singleYearSet, classes, itemCode)
        If mod_Statistics.ArrLen(arr) > 0 Then
            yearLabels.Add CStr(y)
            yearArrays.Add arr
        End If
    Next y
End Sub

' 個人別一覧（学生番号・氏名・点数・偏差値）を、選択条件全体を母集団として計算する。
Public Sub BuildStudentDetailList(ByVal years As Collection, ByVal classes As Collection, ByVal itemCode As String, _
                                   ByRef studentNos() As String, ByRef studentNames() As String, _
                                   ByRef scoreVals() As Double, ByRef devVals() As Double)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_SCORES, mod_Common.TBL_SCORES)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim yearSet As Collection, classSet As Collection
    Set yearSet = ToLookupSet(years)
    Set classSet = ToLookupSet(classes)

    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim cYear As Long, cClass As Long, cStudent As Long, cName As Long, cItem As Long, cScore As Long
    cYear = tbl.ListColumns(mod_Common.COL_YEAR).Index
    cClass = tbl.ListColumns(mod_Common.COL_CLASSCODE).Index
    cStudent = tbl.ListColumns(mod_Common.COL_STUDENTNO).Index
    cName = tbl.ListColumns(mod_Common.COL_NAME).Index
    cItem = tbl.ListColumns(mod_Common.COL_ITEMCODE).Index
    cScore = tbl.ListColumns(mod_Common.COL_SCORE).Index

    ReDim studentNos(1 To UBound(data, 1))
    ReDim studentNames(1 To UBound(data, 1))
    ReDim scoreVals(1 To UBound(data, 1))

    Dim n As Long, i As Long
    n = 0
    For i = 1 To UBound(data, 1)
        If CStr(data(i, cItem)) = itemCode And _
           mod_Common.CollectionHasKey(yearSet, CStr(data(i, cYear))) And _
           mod_Common.CollectionHasKey(classSet, CStr(data(i, cClass))) Then
            n = n + 1
            studentNos(n) = CStr(data(i, cStudent))
            studentNames(n) = CStr(data(i, cName))
            scoreVals(n) = CDbl(data(i, cScore))
        End If
    Next i

    If n = 0 Then
        Erase studentNos: Erase studentNames: Erase scoreVals
        Exit Sub
    End If

    ReDim Preserve studentNos(1 To n)
    ReDim Preserve studentNames(1 To n)
    ReDim Preserve scoreVals(1 To n)
    ReDim devVals(1 To n)

    Dim m As Double, sd As Double
    m = mod_Statistics.Mean(scoreVals)
    sd = mod_Statistics.StDevPopulation(scoreVals)
    For i = 1 To n
        devVals(i) = mod_Statistics.DeviationScore(scoreVals(i), m, sd)
    Next i
End Sub

' 基本統計表（平均・中央値・最頻値・標準偏差・最大値・最小値・四分位数・件数）を構築する。
Public Sub BuildStatsTable(ByRef arr() As Double, ByRef labels() As String, ByRef values() As Variant)
    labels = Split("件数,平均,中央値,最頻値,標準偏差,最小値,第1四分位数,第3四分位数,最大値", ",")
    ReDim values(LBound(labels) To UBound(labels))

    values(0) = mod_Statistics.ArrLen(arr)
    values(1) = Application.Round(mod_Statistics.Mean(arr), 2)
    values(2) = Application.Round(mod_Statistics.Median(arr), 2)
    values(3) = Application.Round(mod_Statistics.Mode(arr), 2)
    values(4) = Application.Round(mod_Statistics.StDevPopulation(arr), 2)
    values(5) = Application.Round(mod_Statistics.MinVal(arr), 2)
    values(6) = Application.Round(mod_Statistics.Quartile(arr, 1), 2)
    values(7) = Application.Round(mod_Statistics.Quartile(arr, 3), 2)
    values(8) = Application.Round(mod_Statistics.MaxVal(arr), 2)
End Sub

'==============================================================================
' 補助
'==============================================================================
Private Function ToLookupSet(ByVal items As Collection) As Collection
    Dim result As New Collection
    Dim v As Variant
    For Each v In items
        If Not mod_Common.CollectionHasKey(result, CStr(v)) Then result.Add True, CStr(v)
    Next v
    Set ToLookupSet = result
End Function
