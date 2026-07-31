Attribute VB_Name = "mod_Import"
Option Explicit
'==============================================================================
' mod_Import
' 目的 : 「データ取込」画面のウィザード処理（①～⑨）を担当する。
'
'   ① 外部Excelを選択        → StartImport
'   ② シート一覧を取得        → StartImport 内
'   ③ 総合成績シートを自動候補表示 → DetectSummarySheet
'   ④ 認識できない場合のみ利用者が選択 → シート一覧リストボックス（常時表示）
'   ⑤ 見出しを自動解析        → DetectHeaderRow / ParseHeaders
'   ⑥ 評価項目を候補表示       → BuildMappingSuggestions / RenderMappingTable
'   ⑦ 必要に応じて利用者が修正   → マッピング表のドロップダウンを直接編集
'   ⑧ インポート内容確認       → ProceedToConfirm
'   ⑨ 登録                  → CommitImport
'
' 元データ（外部ブック）はインポート完了後、保存せずに閉じる。
' ワークシートの内容はセルへ直接コピーせず、必要な値のみを配列として
' メモリ上に取り出してから内部テーブルへ書き込む（元データを保持しない）。
'==============================================================================

' --- ウィザード内部状態（Private：本モジュール内のみで完結させる） ---
Private mSrcWorkbook As Workbook
Private mSrcFilePath As String
Private mSrcSheetName As String
Private mHeaderRow As Long
Private mLastDataRow As Long
Private mLastCol As Long
Private mHeaders() As String
Private mSamples() As String
Private mColItemCode() As String   ' 列ごとに確定した項目コード（マッピング解決後にキャッシュ）

'==============================================================================
' ① 外部Excelを選択 ～ ③ 総合成績シート候補表示
'==============================================================================
Public Sub StartImport()
    On Error GoTo ErrHandler

    ResetWizardState
    mod_UI.ResetImportSheet

    Dim fPath As Variant
    fPath = Application.GetOpenFilename( _
        "Excel ファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
        "分析対象の外部Excelファイルを選択してください")

    If fPath = False Then
        mod_Logging.WriteLog "INFO", "mod_Import", "StartImport", "ファイル選択がキャンセルされました。"
        Exit Sub
    End If
    mSrcFilePath = CStr(fPath)

    mod_Common.BeginBusy
    On Error Resume Next
    Set mSrcWorkbook = Workbooks.Open(mSrcFilePath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error GoTo 0
    mod_Common.EndBusy

    If mSrcWorkbook Is Nothing Then
        mod_Logging.WriteLog "ERROR", "mod_Import", "StartImport", _
            "ファイルを開けませんでした（破損または非対応形式の可能性）。", mSrcFilePath
        MsgBox "ファイルを開けませんでした。破損しているか、対応していない形式の可能性があります。", _
               vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    ThisWorkbook.Worksheets(mod_Common.SH_IMPORT).Range(mod_Common.IMP_CELL_FILEPATH).Value = mSrcFilePath

    ' ② シート一覧を取得し、③ 総合成績シートの候補を自動選定する
    Dim sheetNames As New Collection
    Dim ws As Worksheet
    For Each ws In mSrcWorkbook.Worksheets
        sheetNames.Add ws.Name
    Next ws

    Dim bestSheet As String
    bestSheet = DetectSummarySheet(mSrcWorkbook)

    Dim bestIndex As Long
    bestIndex = 0
    Dim i As Long
    For i = 1 To sheetNames.Count
        If sheetNames(i) = bestSheet Then bestIndex = i
    Next i

    mod_UI.PopulateFormListBox ThisWorkbook.Worksheets(mod_Common.SH_IMPORT), mod_Common.IMP_LISTBOX_SHEETS, _
        sheetNames, bestIndex
    mod_UI.SetImportStatus "シート一覧を取得しました。総合成績シートを確認し「シート決定」を押してください。"
    mod_UI.ShowImportStep 1

    mod_Logging.WriteLog "INFO", "mod_Import", "StartImport", "ファイルを読み込みました: " & mSrcFilePath
    Exit Sub

ErrHandler:
    mod_Common.EndBusy
    CleanupSourceWorkbook
    mod_Common.HandleError "mod_Import", "StartImport"
End Sub

' シート名・見出し内容の両面から「総合成績」シートらしさを判定する。
' シート名変更・レイアウト変更があっても、内容（学生番号・氏名等の見出し）から
' 推定できるようにするための保険的ロジック。
Private Function DetectSummarySheet(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim bestName As String
    Dim bestScore As Long
    bestScore = -1

    For Each ws In wb.Worksheets
        Dim score As Long
        score = 0

        Dim nm As String
        nm = mod_Common.NormalizeText(ws.Name)
        If InStr(nm, mod_Common.NormalizeText("総合成績")) > 0 Then
            score = score + 100
        ElseIf InStr(nm, mod_Common.NormalizeText("総合")) > 0 Then
            score = score + 70
        ElseIf InStr(nm, mod_Common.NormalizeText("成績")) > 0 Then
            score = score + 40
        End If

        score = score + ContentHeaderScore(ws)

        If score > bestScore Then
            bestScore = score
            bestName = ws.Name
        End If
    Next ws

    DetectSummarySheet = bestName
End Function

' シート先頭付近を走査し、学生番号・氏名等の見出しがどれだけ含まれるかを採点する。
Private Function ContentHeaderScore(ByVal ws As Worksheet) As Long
    Dim r As Long, c As Long, score As Long
    Dim maxRow As Long, maxCol As Long
    maxRow = Application.WorksheetFunction.Min(15, ws.UsedRange.Rows.Count)
    maxCol = Application.WorksheetFunction.Min(40, ws.UsedRange.Columns.Count)
    If maxRow < 1 Or maxCol < 1 Then Exit Function

    Dim keywords As Variant
    keywords = Array("学生番号", "生徒番号", "氏名", "クラス", "組", "総合評価")

    For r = 1 To maxRow
        For c = 1 To maxCol
            Dim v As String
            v = mod_Common.NormalizeText(ws.Cells(r, c).Value)
            If Len(v) > 0 Then
                Dim k As Variant
                For Each k In keywords
                    If InStr(v, mod_Common.NormalizeText(CStr(k))) > 0 Then
                        score = score + 10
                        Exit For
                    End If
                Next k
            End If
        Next c
    Next r
    ContentHeaderScore = score
End Function

'==============================================================================
' ④～⑤ シート確定・見出し解析
'==============================================================================
Public Sub ConfirmSheetSelection()
    On Error GoTo ErrHandler

    Dim chosen As String
    chosen = mod_UI.GetFormListBoxSelection(ThisWorkbook.Worksheets(mod_Common.SH_IMPORT), mod_Common.IMP_LISTBOX_SHEETS)
    If Len(chosen) = 0 Then
        MsgBox "シートを選択してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If
    If mSrcWorkbook Is Nothing Then
        MsgBox "先にファイルを選択してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    mSrcSheetName = chosen
    Dim ws As Worksheet
    Set ws = mSrcWorkbook.Worksheets(mSrcSheetName)

    mHeaderRow = DetectHeaderRow(ws)
    ParseHeaders ws
    ResolveColumnItemCodesReset
    BuildMappingSuggestions
    RenderMappingTable

    mod_UI.SetImportStatus "見出しを解析しました。マッピング内容を確認し、必要に応じて修正してください。"
    mod_UI.ShowImportStep 2

    mod_Logging.WriteLog "INFO", "mod_Import", "ConfirmSheetSelection", _
        "シートを確定しました: " & mSrcSheetName & " (見出し行=" & mHeaderRow & ")"
    Exit Sub

ErrHandler:
    mod_Common.HandleError "mod_Import", "ConfirmSheetSelection"
End Sub

' 先頭15行を走査し、既知の見出しキーワードを最も多く含む行を見出し行と推定する。
Private Function DetectHeaderRow(ByVal ws As Worksheet) As Long
    Dim keywords As Variant
    keywords = Array("学生番号", "生徒番号", "氏名", "クラス", "組", "番号", "得点", "点数", "評価")

    Dim r As Long, bestRow As Long, bestScore As Long
    Dim maxRow As Long
    maxRow = Application.WorksheetFunction.Min(15, ws.UsedRange.Rows.Count)
    bestRow = 1
    bestScore = -1

    For r = 1 To maxRow
        Dim c As Long, maxCol As Long, score As Long
        maxCol = Application.WorksheetFunction.Min(60, ws.UsedRange.Columns.Count)
        score = 0
        For c = 1 To maxCol
            Dim v As String
            v = mod_Common.NormalizeText(ws.Cells(r, c).Value)
            If Len(v) > 0 Then
                Dim k As Variant
                For Each k In keywords
                    If InStr(v, mod_Common.NormalizeText(CStr(k))) > 0 Then
                        score = score + 1
                        Exit For
                    End If
                Next k
            End If
        Next c
        If score > bestScore Then
            bestScore = score
            bestRow = r
        End If
    Next r

    DetectHeaderRow = bestRow
End Function

' 見出し行・データ範囲・サンプル値を取得する（元データはこの配列以外に保持しない）。
Private Sub ParseHeaders(ByVal ws As Worksheet)
    mLastCol = ws.Cells(mHeaderRow, ws.Columns.Count).End(xlToLeft).Column
    If mLastCol > mod_Common.IMP_MAPPING_MAX_ROWS Then
        mod_Logging.WriteLog "WARN", "mod_Import", "ParseHeaders", _
            "列数が上限(" & mod_Common.IMP_MAPPING_MAX_ROWS & ")を超えたため、超過分は取り込み対象外とします。", _
            "検出列数=" & mLastCol
        mLastCol = mod_Common.IMP_MAPPING_MAX_ROWS
    End If

    Dim c As Long, tmp As Long
    mLastDataRow = mHeaderRow
    For c = 1 To mLastCol
        tmp = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If tmp > mLastDataRow Then mLastDataRow = tmp
    Next c

    ReDim mHeaders(1 To mLastCol)
    ReDim mSamples(1 To mLastCol)
    ReDim mColItemCode(1 To mLastCol)

    For c = 1 To mLastCol
        mHeaders(c) = Trim$(CStr(ws.Cells(mHeaderRow, c).Value))
        Dim r As Long
        For r = mHeaderRow + 1 To mLastDataRow
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
                mSamples(c) = CStr(ws.Cells(r, c).Value)
                Exit For
            End If
        Next r
    Next c
End Sub

Private Sub ResolveColumnItemCodesReset()
    Dim c As Long
    For c = LBound(mColItemCode) To UBound(mColItemCode)
        mColItemCode(c) = ""
    Next c
End Sub

'==============================================================================
' ⑥ 評価項目候補表示
'==============================================================================
' 各列見出しから、学生属性（年度／クラス／学生番号／氏名）・既存評価項目・
' 無視のいずれに割り当てるべきかを推定し、mod_UI 経由でマッピング表を描画する。
Private mSuggestions() As String

Private Sub BuildMappingSuggestions()
    Dim existingItems As Collection
    Set existingItems = GetActiveItemDisplayNames()

    ReDim mSuggestions(1 To mLastCol)
    Dim c As Long
    For c = 1 To mLastCol
        mSuggestions(c) = GuessTarget(mHeaders(c), mSamples(c), existingItems)
    Next c
End Sub

Private Function GuessTarget(ByVal header As String, ByVal sample As String, ByVal existingItems As Collection) As String
    If mod_Common.HeaderSimilarity(header, "学生番号") >= 60 Or _
       mod_Common.HeaderSimilarity(header, "生徒番号") >= 60 Or _
       mod_Common.HeaderSimilarity(header, "出席番号") >= 60 Or _
       mod_Common.HeaderSimilarity(header, "学籍番号") >= 60 Then
        GuessTarget = mod_Common.MAP_TARGET_STUDENTNO
        Exit Function
    End If
    If mod_Common.HeaderSimilarity(header, "氏名") >= 60 Or mod_Common.HeaderSimilarity(header, "名前") >= 60 Then
        GuessTarget = mod_Common.MAP_TARGET_NAME
        Exit Function
    End If
    If mod_Common.HeaderSimilarity(header, "クラスコード") >= 60 Or _
       mod_Common.HeaderSimilarity(header, "クラス") >= 60 Or _
       mod_Common.HeaderSimilarity(header, "組") >= 60 Or _
       mod_Common.HeaderSimilarity(header, "学級") >= 60 Then
        GuessTarget = mod_Common.MAP_TARGET_CLASSCODE
        Exit Function
    End If
    If mod_Common.HeaderSimilarity(header, "年度") >= 60 Then
        GuessTarget = mod_Common.MAP_TARGET_YEAR
        Exit Function
    End If

    ' 既存の評価項目名と一致すれば再利用を提案する
    Dim it As Variant
    For Each it In existingItems
        If mod_Common.HeaderSimilarity(header, CStr(it)) >= 90 Then
            GuessTarget = CStr(it)
            Exit Function
        End If
    Next it

    ' 数値サンプルであれば新規評価項目として提案、それ以外は無視を提案
    If IsNumeric(sample) Then
        GuessTarget = mod_Common.MAP_TARGET_NEWITEM_PREFIX & header
    Else
        GuessTarget = mod_Common.MAP_TARGET_IGNORE
    End If
End Function

Private Function GetActiveItemDisplayNames() As Collection
    Dim result As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            Dim r As ListRow
            For Each r In tbl.ListRows
                result.Add CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_DISPNAME).Index).Value)
            Next r
        End If
    End If
    Set GetActiveItemDisplayNames = result
End Function

' マッピング表（見出し・サンプル・マッピング先ドロップダウン）を描画する。
' 実際の描画（セル書き込み・データ入力規則の設定）は mod_UI に委譲する。
Private Sub RenderMappingTable()
    Dim choices As New Collection
    choices.Add mod_Common.MAP_TARGET_IGNORE
    choices.Add mod_Common.MAP_TARGET_YEAR
    choices.Add mod_Common.MAP_TARGET_CLASSCODE
    choices.Add mod_Common.MAP_TARGET_STUDENTNO
    choices.Add mod_Common.MAP_TARGET_NAME

    Dim it As Variant
    For Each it In GetActiveItemDisplayNames()
        choices.Add CStr(it)
    Next it

    mod_UI.RenderImportMappingTable mHeaders, mSamples, mSuggestions, choices
End Sub

' 現在のマッピング内容をテンプレートとして保存する（画面の「テンプレート保存」ボタンから使用）。
' 同じレイアウトの外部ファイルを翌年度以降も取り込む際、見出しとマッピング先の対応を
' 再利用できるようにする（自動適用は将来拡張。保守マニュアル参照）。
Public Sub SaveCurrentMappingAsTemplate()
    On Error GoTo ErrHandler
    If mLastCol = 0 Then
        MsgBox "保存するマッピングがありません。マッピング表を表示した状態で実行してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    Dim tplName As Variant
    tplName = InputBox("テンプレート名を入力してください。", mod_Common.APP_NAME)
    If Len(Trim$(CStr(tplName))) = 0 Then Exit Sub

    Dim mapVals As Collection
    Set mapVals = mod_UI.GetImportMappingSelections(mLastCol)

    Dim parts As String, c As Long
    For c = 1 To mLastCol
        If c > 1 Then parts = parts & ";"
        parts = parts & Replace(mHeaders(c), "=", "") & "=" & Replace(CStr(mapVals(c)), "=", "")
    Next c

    mod_Settings.SaveTemplate CStr(tplName), mSrcSheetName, parts
    MsgBox "テンプレートを保存しました。", vbInformation, mod_Common.APP_NAME
    Exit Sub

ErrHandler:
    mod_Common.HandleError "mod_Import", "SaveCurrentMappingAsTemplate"
End Sub

' 現在のマッピング表（画面上のドロップダウン値）を読み取り、列ごとの役割へ変換する。
' colRole(c)     : "IGNORE" / "YEAR" / "CLASSCODE" / "STUDENTNO" / "NAME" / "ITEM"
' colItemName(c) : colRole(c)="ITEM" の場合のみ有効。既存項目コード、または "NEW:表示名"
' ok = False の場合は、呼び出し元は処理を中断すること（エラーメッセージは本関数内で表示済み）。
Private Sub ResolveMapping(ByRef colRole() As String, ByRef colItemName() As String, _
                            ByRef yearColIdx As Long, ByRef classColIdx As Long, _
                            ByRef studentColIdx As Long, ByRef nameColIdx As Long, _
                            ByRef manualYear As String, ByRef manualClass As String, _
                            ByRef ok As Boolean)
    ok = False
    ReDim colRole(1 To mLastCol)
    ReDim colItemName(1 To mLastCol)

    Dim mapVals As Collection
    Set mapVals = mod_UI.GetImportMappingSelections(mLastCol)

    Dim c As Long
    For c = 1 To mLastCol
        Dim v As String
        v = CStr(mapVals(c))
        Select Case v
            Case mod_Common.MAP_TARGET_IGNORE
                colRole(c) = "IGNORE"
            Case mod_Common.MAP_TARGET_YEAR
                colRole(c) = "YEAR": yearColIdx = c
            Case mod_Common.MAP_TARGET_CLASSCODE
                colRole(c) = "CLASSCODE": classColIdx = c
            Case mod_Common.MAP_TARGET_STUDENTNO
                colRole(c) = "STUDENTNO": studentColIdx = c
            Case mod_Common.MAP_TARGET_NAME
                colRole(c) = "NAME": nameColIdx = c
            Case Else
                If Len(v) = 0 Then
                    colRole(c) = "IGNORE"
                Else
                    colRole(c) = "ITEM"
                    If Left$(v, Len(mod_Common.MAP_TARGET_NEWITEM_PREFIX)) = mod_Common.MAP_TARGET_NEWITEM_PREFIX Then
                        colItemName(c) = "NEW:" & Mid$(v, Len(mod_Common.MAP_TARGET_NEWITEM_PREFIX) + 1)
                    Else
                        Dim existingCode As String
                        existingCode = mod_Settings.FindItemCodeByDisplayName(v)
                        If Len(existingCode) > 0 Then
                            colItemName(c) = existingCode
                        Else
                            colItemName(c) = "NEW:" & v
                        End If
                    End If
                End If
        End Select
    Next c

    If studentColIdx = 0 Or nameColIdx = 0 Then
        MsgBox "「学生番号」と「氏名」は必ずマッピングしてください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    manualYear = Trim$(CStr(ThisWorkbook.Worksheets(mod_Common.SH_IMPORT).Range(mod_Common.IMP_CELL_YEAR_MANUAL).Value))
    manualClass = Trim$(CStr(ThisWorkbook.Worksheets(mod_Common.SH_IMPORT).Range(mod_Common.IMP_CELL_CLASS_MANUAL).Value))

    If yearColIdx = 0 And Len(manualYear) = 0 Then
        MsgBox "年度が元データの列に無い場合は、画面右上の「年度（手入力）」欄に入力してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If
    If classColIdx = 0 And Len(manualClass) = 0 Then
        MsgBox "クラスコードが元データの列に無い場合は、画面右上の「クラスコード（手入力）」欄に入力してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    Dim hasItem As Boolean
    For c = 1 To mLastCol
        If colRole(c) = "ITEM" Then hasItem = True
    Next c
    If Not hasItem Then
        MsgBox "評価項目として取り込む列が1つもありません。マッピングを確認してください。", vbExclamation, mod_Common.APP_NAME
        Exit Sub
    End If

    ok = True
End Sub

'==============================================================================
' ⑧ インポート内容確認
'==============================================================================
Public Sub ProceedToConfirm()
    On Error GoTo ErrHandler
    If mSrcWorkbook Is Nothing Then Exit Sub

    Dim colRole() As String, colItemName() As String
    Dim yearColIdx As Long, classColIdx As Long, studentColIdx As Long, nameColIdx As Long
    Dim manualYear As String, manualClass As String, ok As Boolean
    ResolveMapping colRole, colItemName, yearColIdx, classColIdx, studentColIdx, nameColIdx, _
                   manualYear, manualClass, ok
    If Not ok Then Exit Sub

    Dim ws As Worksheet
    Set ws = mSrcWorkbook.Worksheets(mSrcSheetName)
    Dim existingKeys As Collection
    Set existingKeys = mod_Database.BuildScoreKeyIndex()

    Dim totalRows As Long, blankSkip As Long, dupCount As Long, newItemCols As Long, validScoreCells As Long
    Dim c As Long
    For c = 1 To mLastCol
        If colRole(c) = "ITEM" Then
            If Left$(colItemName(c), 4) = "NEW:" Then newItemCols = newItemCols + 1
        End If
    Next c

    Dim r As Long
    For r = mHeaderRow + 1 To mLastDataRow
        totalRows = totalRows + 1
        Dim sno As String, snm As String
        sno = Trim$(CStr(ws.Cells(r, studentColIdx).Value))
        snm = Trim$(CStr(ws.Cells(r, nameColIdx).Value))
        If Len(sno) = 0 Or Len(snm) = 0 Then
            blankSkip = blankSkip + 1
        Else
            Dim yv As String, cv As String
            If yearColIdx > 0 Then
                yv = Trim$(CStr(ws.Cells(r, yearColIdx).Value))
            Else
                yv = manualYear
            End If
            If classColIdx > 0 Then
                cv = mod_Common.FormatClassCode(ws.Cells(r, classColIdx).Value)
            Else
                cv = mod_Common.FormatClassCode(manualClass)
            End If
            For c = 1 To mLastCol
                If colRole(c) = "ITEM" Then
                    Dim sc As Variant
                    sc = ws.Cells(r, c).Value
                    If Len(Trim$(CStr(sc))) > 0 And IsNumeric(sc) Then
                        validScoreCells = validScoreCells + 1
                        If Left$(colItemName(c), 4) <> "NEW:" Then
                            Dim key As String
                            key = mod_Common.MakeKey(yv, cv, sno, colItemName(c))
                            If mod_Common.CollectionHasKey(existingKeys, key) Then dupCount = dupCount + 1
                        End If
                    End If
                End If
            Next c
        End If
    Next r

    Dim summary As String
    summary = "対象データ行数: " & totalRows & vbCrLf & _
              "氏名/学生番号が空欄のためスキップ: " & blankSkip & " 行" & vbCrLf & _
              "登録対象の得点セル数: " & validScoreCells & vbCrLf & _
              "既存データと重複（既定では登録をスキップします）: " & dupCount & " 件" & vbCrLf & _
              "新規に追加される評価項目数: " & newItemCols

    mod_UI.SetImportConfirmSummary summary
    mod_UI.ShowImportStep 3
    mod_UI.SetImportStatus "内容を確認し、問題なければ「登録」を押してください。"

    mod_Logging.WriteLog "INFO", "mod_Import", "ProceedToConfirm", "取込内容を確認しました。" & summary
    Exit Sub

ErrHandler:
    mod_Common.HandleError "mod_Import", "ProceedToConfirm"
End Sub

'==============================================================================
' ⑨ 登録
'==============================================================================
Public Sub CommitImport()
    On Error GoTo ErrHandler
    If mSrcWorkbook Is Nothing Then Exit Sub

    Dim colRole() As String, colItemName() As String
    Dim yearColIdx As Long, classColIdx As Long, studentColIdx As Long, nameColIdx As Long
    Dim manualYear As String, manualClass As String, ok As Boolean
    ResolveMapping colRole, colItemName, yearColIdx, classColIdx, studentColIdx, nameColIdx, _
                   manualYear, manualClass, ok
    If Not ok Then Exit Sub

    mod_Common.BeginBusy

    ' 新規評価項目は列単位で一度だけ作成する（行ループ内での重複作成を防止）
    Dim c As Long
    For c = 1 To mLastCol
        If colRole(c) = "ITEM" Then
            If Left$(colItemName(c), 4) = "NEW:" Then
                colItemName(c) = mod_Settings.GetOrCreateItem(Mid$(colItemName(c), 5), "")
            End If
        End If
    Next c

    Dim overwrite As Boolean
    overwrite = mod_UI.GetCheckBoxValue(ThisWorkbook.Worksheets(mod_Common.SH_IMPORT), mod_Common.IMP_CHK_OVERWRITE)

    Dim existingKeys As Collection, rowIndex As Collection
    Set existingKeys = mod_Database.BuildScoreKeyIndex()
    If overwrite Then Set rowIndex = mod_Database.BuildScoreRowIndex()

    Dim ws As Worksheet
    Set ws = mSrcWorkbook.Worksheets(mSrcSheetName)

    Dim newRows As New Collection
    Dim touchedPairs As New Collection
    Dim importedCount As Long, updatedCount As Long, blankSkip As Long, dupSkip As Long, invalidScore As Long

    Dim r As Long
    For r = mHeaderRow + 1 To mLastDataRow
        Dim sno As String, snm As String
        sno = Trim$(CStr(ws.Cells(r, studentColIdx).Value))
        snm = Trim$(CStr(ws.Cells(r, nameColIdx).Value))
        If Len(sno) = 0 Or Len(snm) = 0 Then
            blankSkip = blankSkip + 1
            GoTo NextRow
        End If

        Dim yv As String, cv As String
        If yearColIdx > 0 Then
            yv = Trim$(CStr(ws.Cells(r, yearColIdx).Value))
        Else
            yv = manualYear
        End If
        If classColIdx > 0 Then
            cv = mod_Common.FormatClassCode(ws.Cells(r, classColIdx).Value)
        Else
            cv = mod_Common.FormatClassCode(manualClass)
        End If

        Dim pairKey As String
        pairKey = yv & "|" & cv
        If Not mod_Common.CollectionHasKey(touchedPairs, pairKey) Then touchedPairs.Add pairKey, pairKey

        For c = 1 To mLastCol
            If colRole(c) = "ITEM" Then
                Dim sc As Variant
                sc = ws.Cells(r, c).Value
                If Len(Trim$(CStr(sc))) = 0 Then
                    ' 未回答（空欄）は無視（エラーではない）
                ElseIf Not IsNumeric(sc) Then
                    invalidScore = invalidScore + 1
                Else
                    Dim key As String
                    key = mod_Common.MakeKey(yv, cv, sno, colItemName(c))
                    If overwrite And mod_Common.CollectionHasKey(rowIndex, key) Then
                        mod_Database.UpdateScoreValue rowIndex.Item(key), CDbl(sc)
                        updatedCount = updatedCount + 1
                    ElseIf mod_Common.CollectionHasKey(existingKeys, key) Then
                        dupSkip = dupSkip + 1
                    Else
                        Dim rowArr(0 To 5) As Variant
                        rowArr(0) = yv: rowArr(1) = cv: rowArr(2) = sno: rowArr(3) = snm
                        rowArr(4) = colItemName(c): rowArr(5) = CDbl(sc)
                        newRows.Add rowArr
                        existingKeys.Add True, key
                        importedCount = importedCount + 1
                    End If
                End If
            End If
        Next c
NextRow:
    Next r

    mod_Database.AppendScoreRows newRows

    Dim pk As Variant
    For Each pk In touchedPairs
        Dim parts() As String
        parts = Split(CStr(pk), "|")
        mod_Database.RecalcSummary parts(0), parts(1)
    Next pk

    CleanupSourceWorkbook
    mod_UI.ResetImportSheet
    mod_Common.EndBusy

    Dim msg As String
    msg = "取込が完了しました。新規登録:" & importedCount & "件 / 上書き更新:" & updatedCount & "件 / " & _
          "空欄スキップ:" & blankSkip & "行 / 重複スキップ:" & dupSkip & "件 / 数値変換不可:" & invalidScore & "件"
    MsgBox msg, vbInformation, mod_Common.APP_NAME
    mod_Logging.WriteLog "INFO", "mod_Import", "CommitImport", msg
    Exit Sub

ErrHandler:
    mod_Common.EndBusy
    CleanupSourceWorkbook
    mod_Common.HandleError "mod_Import", "CommitImport"
End Sub

'==============================================================================
' キャンセル・後始末
'==============================================================================

' ウィザードの途中で「キャンセル」が押された場合に呼び出す。
' 登録は一切行わず、開いていた外部ブックを保存せずに閉じるだけで安全に中断できる
' （T_Scores への書き込みは CommitImport の最後で一括実行されるため、
'   途中キャンセルでも中途半端なデータが残ることはない）。
Public Sub CancelImport()
    CleanupSourceWorkbook
    mod_UI.ResetImportSheet
    mod_UI.SetImportStatus "取込をキャンセルしました。"
    mod_Logging.WriteLog "INFO", "mod_Import", "CancelImport", "ユーザーにより取込がキャンセルされました。"
End Sub

Private Sub ResetWizardState()
    CleanupSourceWorkbook
    mSrcFilePath = ""
    mSrcSheetName = ""
    mHeaderRow = 0
    mLastDataRow = 0
    mLastCol = 0
    Erase mHeaders
    Erase mSamples
    Erase mColItemCode
End Sub

' 外部ブックを保存せずに閉じる（元データは一切保存しない方針の徹底）。
Private Sub CleanupSourceWorkbook()
    On Error Resume Next
    If Not mSrcWorkbook Is Nothing Then
        mSrcWorkbook.Close SaveChanges:=False
    End If
    Set mSrcWorkbook = Nothing
    On Error GoTo 0
End Sub
