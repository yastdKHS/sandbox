Attribute VB_Name = "mod_Statistics"
Option Explicit
'==============================================================================
' mod_Statistics
' 目的 : 数値配列（Double型、1-based）に対する統計計算を提供する純粋関数群。
'        ワークシート関数や外部ライブラリに依存せず、VBAのみで実装する
'        （将来 Excel のバージョンが変わっても計算結果が変わらないようにするため）。
'
' 標準偏差・偏差値は「母集団（分析対象グループそのもの）」を前提とし、
' 母標準偏差（nで除算）を用いる。これは学校教育における偏差値算出の
' 一般的な慣習に合わせたものである。
'==============================================================================

' ---- 基本統計 ----

Public Function Mean(ByRef arr() As Double) As Double
    Dim n As Long
    n = ArrLen(arr)
    If n = 0 Then Exit Function
    Dim s As Double, i As Long
    For i = LBound(arr) To UBound(arr)
        s = s + arr(i)
    Next i
    Mean = s / n
End Function

Public Function StDevPopulation(ByRef arr() As Double) As Double
    Dim n As Long
    n = ArrLen(arr)
    If n = 0 Then Exit Function
    Dim m As Double, s As Double, i As Long
    m = Mean(arr)
    For i = LBound(arr) To UBound(arr)
        s = s + (arr(i) - m) ^ 2
    Next i
    StDevPopulation = Sqr(s / n)
End Function

Public Function MinVal(ByRef arr() As Double) As Double
    Dim n As Long
    n = ArrLen(arr)
    If n = 0 Then Exit Function
    Dim v As Double, i As Long
    v = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If arr(i) < v Then v = arr(i)
    Next i
    MinVal = v
End Function

Public Function MaxVal(ByRef arr() As Double) As Double
    Dim n As Long
    n = ArrLen(arr)
    If n = 0 Then Exit Function
    Dim v As Double, i As Long
    v = arr(LBound(arr))
    For i = LBound(arr) To UBound(arr)
        If arr(i) > v Then v = arr(i)
    Next i
    MaxVal = v
End Function

Public Function Median(ByRef arr() As Double) As Double
    Dim s() As Double
    s = SortedCopy(arr)
    Dim n As Long
    n = ArrLen(s)
    If n = 0 Then Exit Function
    If n Mod 2 = 1 Then
        Median = s(LBound(s) + (n - 1) \ 2)
    Else
        Dim lo As Long
        lo = LBound(s) + n \ 2 - 1
        Median = (s(lo) + s(lo + 1)) / 2
    End If
End Function

' 最頻値。複数存在する場合は最小の値を返す（結果を一意にするため）。
' 度数分布上の最頻値であり、連続量の推定モードではない点に留意。
Public Function Mode(ByRef arr() As Double) As Double
    Dim n As Long
    n = ArrLen(arr)
    If n = 0 Then Exit Function

    Dim s() As Double
    s = SortedCopy(arr)

    Dim bestVal As Double, bestCount As Long
    Dim curVal As Double, curCount As Long
    Dim i As Long
    curVal = s(LBound(s))
    curCount = 1
    bestVal = curVal
    bestCount = 1

    For i = LBound(s) + 1 To UBound(s)
        If s(i) = curVal Then
            curCount = curCount + 1
        Else
            curVal = s(i)
            curCount = 1
        End If
        If curCount > bestCount Then
            bestCount = curCount
            bestVal = curVal
        End If
    Next i
    Mode = bestVal
End Function

' Excel の QUARTILE.INC と同じ線形補間方式で分位点を求める。
' q = 0(最小値) / 1(第1四分位) / 2(中央値) / 3(第3四分位) / 4(最大値)
Public Function Quartile(ByRef arr() As Double, ByVal q As Long) As Double
    Dim s() As Double
    s = SortedCopy(arr)
    Dim n As Long
    n = ArrLen(s)
    If n = 0 Then Exit Function
    If n = 1 Then
        Quartile = s(LBound(s))
        Exit Function
    End If

    Dim rank As Double
    rank = q / 4# * (n - 1)   ' 0-based rank
    Dim lo As Long, hi As Long, frac As Double
    lo = Int(rank)
    hi = lo + 1
    frac = rank - lo
    If hi > n - 1 Then hi = n - 1

    Quartile = s(LBound(s) + lo) + frac * (s(LBound(s) + hi) - s(LBound(s) + lo))
End Function

' 偏差値 = 50 + 10 × (x - 平均) / 標準偏差。標準偏差が0の場合は50を返す。
Public Function DeviationScore(ByVal x As Double, ByVal meanVal As Double, ByVal sdVal As Double) As Double
    If sdVal <= 0 Then
        DeviationScore = 50
    Else
        DeviationScore = 50 + 10 * (x - meanVal) / sdVal
    End If
End Function

' 正規分布の確率密度関数（理論分布のヒストグラム重畳表示に使用）。
Public Function NormalPDF(ByVal x As Double, ByVal meanVal As Double, ByVal sdVal As Double) As Double
    If sdVal <= 0 Then Exit Function
    Const PI As Double = 3.14159265358979
    NormalPDF = (1# / (sdVal * Sqr(2 * PI))) * Exp(-((x - meanVal) ^ 2) / (2 * sdVal ^ 2))
End Function

' ---- ヒストグラム ----

' ビンの区切り（境界値）を、指定範囲(minV～maxV)に対して計算する。
' mode = "COUNT" のとき binCount 個の等幅ビンを作る。
' mode = "WIDTH" のとき binWidth 幅で、データ範囲を覆うだけのビンを作る。
' 複数年度を同一ビン幅で重畳比較する場合は、全年度合算の範囲から
' 求めた境界値を各年度のデータへ共通適用する（CountIntoBins を使用）。
Public Function ComputeBinEdges(ByVal minV As Double, ByVal maxV As Double, _
                                 ByVal mode As String, ByVal binCount As Long, ByVal binWidth As Double) As Double()
    If maxV = minV Then maxV = minV + 1   ' 全データが同一値の場合の縮退対応

    Dim nBins As Long, w As Double
    If UCase$(mode) = "WIDTH" And binWidth > 0 Then
        w = binWidth
        nBins = Application.WorksheetFunction.Ceiling_Precise((maxV - minV) / w)
        If nBins < 1 Then nBins = 1
    Else
        nBins = binCount
        If nBins < 1 Then nBins = 1
        w = (maxV - minV) / nBins
    End If

    Dim edges() As Double
    ReDim edges(0 To nBins)
    Dim i As Long
    For i = 0 To nBins
        edges(i) = minV + i * w
    Next i
    ComputeBinEdges = edges
End Function

' 指定されたビン境界値に従って、配列の値を各ビンへ集計する。
Public Function CountIntoBins(ByRef arr() As Double, ByRef edges() As Double) As Long()
    Dim nBins As Long
    nBins = ArrLen(edges) - 1
    Dim counts() As Long
    ReDim counts(1 To nBins)
    If ArrLen(arr) = 0 Then
        CountIntoBins = counts
        Exit Function
    End If

    Dim minV As Double, w As Double
    minV = edges(LBound(edges))
    w = edges(LBound(edges) + 1) - edges(LBound(edges))

    Dim i As Long, idx As Long
    For i = LBound(arr) To UBound(arr)
        If w = 0 Then
            idx = 1
        Else
            idx = Int((arr(i) - minV) / w) + 1
            If idx > nBins Then idx = nBins
            If idx < 1 Then idx = 1
        End If
        counts(idx) = counts(idx) + 1
    Next i
    CountIntoBins = counts
End Function

' 単一データ配列に対するビン境界・度数を一括計算する（従来互換の簡易版）。
Public Sub ComputeHistogram(ByRef arr() As Double, ByVal mode As String, _
                             ByVal binCount As Long, ByVal binWidth As Double, _
                             ByRef binEdges() As Double, ByRef binCounts() As Long)
    If ArrLen(arr) = 0 Then Exit Sub
    binEdges = ComputeBinEdges(MinVal(arr), MaxVal(arr), mode, binCount, binWidth)
    binCounts = CountIntoBins(arr, binEdges)
End Sub

' ---- 累積分布 ----

' 昇順ソート済みの値と、その累積相対度数（%）を返す（折れ線グラフ用）。
Public Sub ComputeCumulative(ByRef arr() As Double, ByRef sortedVals() As Double, ByRef cumPct() As Double)
    Dim n As Long
    n = ArrLen(arr)
    If n = 0 Then Exit Sub
    sortedVals = SortedCopy(arr)
    ReDim cumPct(LBound(sortedVals) To UBound(sortedVals))
    Dim i As Long
    For i = LBound(sortedVals) To UBound(sortedVals)
        cumPct(i) = (i - LBound(sortedVals) + 1) / n * 100
    Next i
End Sub

' ---- 補助 ----

Public Function ArrLen(ByRef arr() As Double) As Long
    On Error Resume Next
    ArrLen = UBound(arr) - LBound(arr) + 1
    On Error GoTo 0
End Function

' 単純な昇順クイックソートのコピーを返す（元配列は変更しない）。
Public Function SortedCopy(ByRef arr() As Double) As Double()
    Dim n As Long
    n = ArrLen(arr)
    Dim result() As Double
    If n = 0 Then
        SortedCopy = result
        Exit Function
    End If
    ReDim result(LBound(arr) To UBound(arr))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        result(i) = arr(i)
    Next i
    QuickSort result, LBound(result), UBound(result)
    SortedCopy = result
End Function

Private Sub QuickSort(ByRef a() As Double, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, pivot As Double, tmp As Double
    i = lo: j = hi
    pivot = a((lo + hi) \ 2)
    Do While i <= j
        Do While a(i) < pivot: i = i + 1: Loop
        Do While a(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSort a, lo, j
    If i < hi Then QuickSort a, i, hi
End Sub
