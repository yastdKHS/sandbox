Attribute VB_Name = "mod_Graph"
Option Explicit
'==============================================================================
' mod_Graph
' 目的 : 「分析」シート上の3種類のグラフ（ヒストグラム／累積分布曲線／箱ひげ図）
'        の内容を更新する。グラフオブジェクト自体は mod_UI が初期構築時に
'        作成済みであることを前提とし、本モジュールはデータ（系列）の
'        差し替えのみを行う（責務分離：構築=mod_UI、内容更新=mod_Graph）。
'
' 系列データは Range を経由せず、配列を直接 Series.Values / .XValues へ
' 代入する。これにより補助セル領域を持つ必要がなく、保守対象が減る。
'==============================================================================

Private Const NORMAL_CURVE_POINTS As Long = 60   ' 正規分布曲線の描画点数

'==============================================================================
' ① ヒストグラム
'==============================================================================
' yearLabels(i) / yearArrays(i) は対になった年度ラベルと点数配列（BuildScoresGroupedByYearの出力）。
' overlayYears = True の場合、共通のビン境界を用いて年度別に複数系列で重畳表示する。
' overlayYears = False の場合、全年度のデータを1本のヒストグラムにまとめて表示する。
' showNormal = True の場合、系列が1本のときのみ理論正規分布曲線を重畳する
' （複数系列重畳時は視認性を優先し曲線は表示しない）。
Public Sub UpdateHistogramChart(ByVal yearLabels As Collection, ByVal yearArrays As Collection, _
                                 ByVal overlayYears As Boolean, ByVal binMode As String, _
                                 ByVal binCount As Long, ByVal binWidth As Double, _
                                 ByVal showNormal As Boolean, ByVal itemLabel As String)
    Dim cht As Chart
    Set cht = GetChart(mod_Common.ANL_CHART_HIST)
    If cht Is Nothing Then Exit Sub
    ClearChart cht
    cht.HasTitle = True
    cht.ChartTitle.Text = "ヒストグラム：" & itemLabel

    Dim combined() As Double
    combined = CombineAll(yearArrays)
    If mod_Statistics.ArrLen(combined) = 0 Then Exit Sub

    Dim edges() As Double
    edges = mod_Statistics.ComputeBinEdges(mod_Statistics.MinVal(combined), mod_Statistics.MaxVal(combined), _
                                            binMode, binCount, binWidth)
    Dim categories() As String
    categories = BuildBinLabels(edges)

    cht.ChartType = xlColumnClustered

    Dim seriesCount As Long
    If overlayYears Then
        Dim i As Long
        For i = 1 To yearLabels.Count
            Dim counts() As Long
            Dim arr() As Double
            arr = yearArrays(i)
            counts = mod_Statistics.CountIntoBins(arr, edges)
            AddColumnSeries cht, CStr(yearLabels(i)) & "年", categories, LongArrToDouble(counts)
            seriesCount = seriesCount + 1
        Next i
    Else
        Dim allCounts() As Long
        allCounts = mod_Statistics.CountIntoBins(combined, edges)
        AddColumnSeries cht, itemLabel, categories, LongArrToDouble(allCounts)
        seriesCount = 1
    End If

    If showNormal And seriesCount = 1 Then
        AddNormalCurveOverlay cht, combined, edges
    End If
End Sub

' 正規分布の理論曲線を、ヒストグラムと同じビン幅における期待度数に換算して重畳する。
Private Sub AddNormalCurveOverlay(ByVal cht As Chart, ByRef arr() As Double, ByRef edges() As Double)
    Dim m As Double, sd As Double, n As Long, w As Double
    m = mod_Statistics.Mean(arr)
    sd = mod_Statistics.StDevPopulation(arr)
    n = mod_Statistics.ArrLen(arr)
    w = edges(LBound(edges) + 1) - edges(LBound(edges))
    If sd <= 0 Then Exit Sub

    Dim xs() As Double, ys() As Double
    ReDim xs(1 To NORMAL_CURVE_POINTS)
    ReDim ys(1 To NORMAL_CURVE_POINTS)

    Dim minX As Double, maxX As Double, step_ As Double
    minX = edges(LBound(edges))
    maxX = edges(UBound(edges))
    step_ = (maxX - minX) / (NORMAL_CURVE_POINTS - 1)

    Dim i As Long
    For i = 1 To NORMAL_CURVE_POINTS
        xs(i) = minX + (i - 1) * step_
        ys(i) = mod_Statistics.NormalPDF(xs(i), m, sd) * n * w
    Next i

    Dim sc As Series
    Set sc = cht.SeriesCollection.NewSeries
    sc.Name = "正規分布（理論値）"
    sc.Values = ys
    sc.XValues = xs
    sc.ChartType = xlXYScatterSmoothNoMarkers
    On Error Resume Next
    sc.AxisGroup = xlSecondary
    On Error GoTo 0
End Sub

'==============================================================================
' ② 累積分布曲線
'==============================================================================
Public Sub UpdateCumulativeChart(ByVal yearLabels As Collection, ByVal yearArrays As Collection, _
                                  ByVal overlayYears As Boolean, ByVal itemLabel As String)
    Dim cht As Chart
    Set cht = GetChart(mod_Common.ANL_CHART_CUM)
    If cht Is Nothing Then Exit Sub
    ClearChart cht
    cht.HasTitle = True
    cht.ChartTitle.Text = "累積分布曲線：" & itemLabel
    cht.ChartType = xlXYScatterLines

    If overlayYears Then
        Dim i As Long
        For i = 1 To yearLabels.Count
            Dim arr() As Double
            arr = yearArrays(i)
            AddCumulativeSeries cht, CStr(yearLabels(i)) & "年", arr
        Next i
    Else
        Dim combined() As Double
        combined = CombineAll(yearArrays)
        AddCumulativeSeries cht, itemLabel, combined
    End If
End Sub

Private Sub AddCumulativeSeries(ByVal cht As Chart, ByVal seriesName As String, ByRef arr() As Double)
    If mod_Statistics.ArrLen(arr) = 0 Then Exit Sub
    Dim sortedVals() As Double, cumPct() As Double
    mod_Statistics.ComputeCumulative arr, sortedVals, cumPct

    Dim sc As Series
    Set sc = cht.SeriesCollection.NewSeries
    sc.Name = seriesName
    sc.XValues = sortedVals
    sc.Values = cumPct
    sc.ChartType = xlXYScatterLines
End Sub

'==============================================================================
' ③ 箱ひげ図（Excel 2016以降の標準チャート種類 xlBoxwhisker を使用）
'==============================================================================
Public Sub UpdateBoxPlotChart(ByVal yearLabels As Collection, ByVal yearArrays As Collection, _
                               ByVal overlayYears As Boolean, ByVal itemLabel As String)
    Dim cht As Chart
    Set cht = GetChart(mod_Common.ANL_CHART_BOX)
    If cht Is Nothing Then Exit Sub
    ClearChart cht
    cht.HasTitle = True
    cht.ChartTitle.Text = "箱ひげ図：" & itemLabel

    On Error Resume Next
    cht.ChartType = xlBoxwhisker
    On Error GoTo 0

    If overlayYears Then
        Dim i As Long
        For i = 1 To yearLabels.Count
            Dim arr() As Double
            arr = yearArrays(i)
            If mod_Statistics.ArrLen(arr) > 0 Then
                Dim sc As Series
                Set sc = cht.SeriesCollection.NewSeries
                sc.Name = CStr(yearLabels(i)) & "年"
                sc.Values = arr
            End If
        Next i
    Else
        Dim combined() As Double
        combined = CombineAll(yearArrays)
        If mod_Statistics.ArrLen(combined) > 0 Then
            Dim sc2 As Series
            Set sc2 = cht.SeriesCollection.NewSeries
            sc2.Name = itemLabel
            sc2.Values = combined
        End If
    End If
End Sub

'==============================================================================
' 補助
'==============================================================================

Private Function GetChart(ByVal chartObjName As String) As Chart
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mod_Common.SH_ANALYSIS)
    On Error Resume Next
    Set GetChart = ws.ChartObjects(chartObjName).Chart
    On Error GoTo 0
End Function

Private Sub ClearChart(ByVal cht As Chart)
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
End Sub

Private Sub AddColumnSeries(ByVal cht As Chart, ByVal seriesName As String, _
                             ByRef categories() As String, ByRef values() As Double)
    Dim sc As Series
    Set sc = cht.SeriesCollection.NewSeries
    sc.Name = seriesName
    sc.XValues = categories
    sc.Values = values
    sc.ChartType = xlColumnClustered
End Sub

Private Function CombineAll(ByVal arrays As Collection) As Double()
    Dim total As Long, i As Long
    Dim a() As Double
    For i = 1 To arrays.Count
        a = arrays(i)
        total = total + mod_Statistics.ArrLen(a)
    Next i

    Dim result() As Double
    If total = 0 Then
        CombineAll = result
        Exit Function
    End If
    ReDim result(1 To total)

    Dim pos As Long, j As Long
    pos = 0
    For i = 1 To arrays.Count
        a = arrays(i)
        For j = LBound(a) To UBound(a)
            pos = pos + 1
            result(pos) = a(j)
        Next j
    Next i
    CombineAll = result
End Function

Private Function BuildBinLabels(ByRef edges() As Double) As String()
    Dim n As Long
    n = UBound(edges) - LBound(edges)
    Dim labels() As String
    ReDim labels(1 To n)
    Dim i As Long
    For i = 1 To n
        labels(i) = Format$(edges(LBound(edges) + i - 1), "0.#") & "-" & Format$(edges(LBound(edges) + i), "0.#")
    Next i
    BuildBinLabels = labels
End Function

Private Function LongArrToDouble(ByRef src() As Long) As Double()
    Dim result() As Double
    Dim n As Long
    n = UBound(src) - LBound(src) + 1
    ReDim result(LBound(src) To UBound(src))
    Dim i As Long
    For i = LBound(src) To UBound(src)
        result(i) = src(i)
    Next i
    LongArrToDouble = result
End Function
