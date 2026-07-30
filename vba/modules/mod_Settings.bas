Attribute VB_Name = "mod_Settings"
Option Explicit
'==============================================================================
' mod_Settings
' 目的 : 「設定」シート上の T_Items（評価項目マスタ）・T_Settings（各種設定値）
'        ・T_Templates（列マッピングテンプレート）へのアクセスを一元管理する。
'        表示名・重みは利用者が「設定」シートのセルを直接編集する運用とし、
'        本モジュールはその読み書きロジックと再計算トリガのみを担当する。
'==============================================================================

'------------------------------------------------------------------
' T_Settings（キー・バリュー設定） アクセサ
'------------------------------------------------------------------

' 設定値を文字列として取得する。未設定の場合は既定値を返す。
Public Function GetSettingText(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_SETTINGS)
    GetSettingText = defaultValue
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    Dim r As ListRow
    For Each r In tbl.ListRows
        If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SET_KEY).Index).Value) = key Then
            GetSettingText = CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SET_VALUE).Index).Value)
            Exit Function
        End If
    Next r
End Function

Public Function GetSettingLong(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
    Dim s As String
    s = GetSettingText(key, CStr(defaultValue))
    If IsNumeric(s) Then
        GetSettingLong = CLng(s)
    Else
        GetSettingLong = defaultValue
    End If
End Function

Public Function GetSettingBool(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim s As String
    s = GetSettingText(key, IIf(defaultValue, "TRUE", "FALSE"))
    GetSettingBool = (UCase$(s) = "TRUE")
End Function

' 設定値を保存する（キーが存在すれば更新、無ければ新規追加）。
Public Sub SetSetting(ByVal key As String, ByVal value As String)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_SETTINGS)
    If tbl Is Nothing Then Exit Sub

    Dim r As ListRow
    If Not tbl.DataBodyRange Is Nothing Then
        For Each r In tbl.ListRows
            If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_SET_KEY).Index).Value) = key Then
                r.Range(1, tbl.ListColumns(mod_Common.COL_SET_VALUE).Index).Value = value
                Exit Sub
            End If
        Next r
    End If

    Set r = tbl.ListRows.Add
    r.Range(1, tbl.ListColumns(mod_Common.COL_SET_KEY).Index).Value = key
    r.Range(1, tbl.ListColumns(mod_Common.COL_SET_VALUE).Index).Value = value
End Sub

'------------------------------------------------------------------
' T_Items（評価項目マスタ） アクセサ
'------------------------------------------------------------------

' 表示名から項目コードを逆引きする（正規化して比較）。見つからなければ空文字。
Public Function FindItemCodeByDisplayName(ByVal dispName As String) As String
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    Dim r As ListRow
    For Each r In tbl.ListRows
        If mod_Common.NormalizeText(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_DISPNAME).Index).Value) _
           = mod_Common.NormalizeText(dispName) Then
            FindItemCodeByDisplayName = CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_CODE).Index).Value)
            Exit Function
        End If
    Next r
End Function

' 項目コードから ListRow を検索する。見つからない場合は Nothing。
Public Function FindItemRow(ByVal itemCode As String) As ListRow
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    Dim r As ListRow
    For Each r In tbl.ListRows
        If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_CODE).Index).Value) = itemCode Then
            Set FindItemRow = r
            Exit Function
        End If
    Next r
End Function

' 項目コードが存在すれば取得し、無ければ新規作成して返す（内部名は自動採番 ScoreN）。
' displayNameHint : 元データの見出し文字列（新規作成時の初期表示名として使用）
Public Function GetOrCreateItem(ByVal displayNameHint As String, ByVal existingItemCode As String) As String
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_SETTINGS, mod_Common.TBL_ITEMS)
    If tbl Is Nothing Then Exit Function

    ' 既存コード指定がある場合はそのまま使用（インポート時の再マッピング用）
    If Len(existingItemCode) > 0 Then
        If Not FindItemRow(existingItemCode) Is Nothing Then
            GetOrCreateItem = existingItemCode
            Exit Function
        End If
    End If

    ' 表示名が既存項目と一致すれば再利用する（年度をまたいだ同一項目の統合）
    Dim r As ListRow
    If Not tbl.DataBodyRange Is Nothing Then
        For Each r In tbl.ListRows
            If mod_Common.NormalizeText(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_DISPNAME).Index).Value) _
               = mod_Common.NormalizeText(displayNameHint) Then
                GetOrCreateItem = CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_CODE).Index).Value)
                Exit Function
            End If
        Next r
    End If

    ' 新規項目コードを採番（ScoreN、既存の最大N+1）
    Dim nextNo As Long
    nextNo = 1
    If Not tbl.DataBodyRange Is Nothing Then
        For Each r In tbl.ListRows
            Dim code As String
            code = CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_CODE).Index).Value)
            If Left$(code, 5) = "Score" And IsNumeric(Mid$(code, 6)) Then
                If CLng(Mid$(code, 6)) >= nextNo Then nextNo = CLng(Mid$(code, 6)) + 1
            End If
        Next r
    End If

    Dim newCode As String
    newCode = "Score" & nextNo

    Set r = tbl.ListRows.Add
    r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_CODE).Index).Value = newCode
    r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_DISPNAME).Index).Value = _
        IIf(Len(displayNameHint) > 0, displayNameHint, newCode)
    r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_WEIGHT).Index).Value = 100
    r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_ORDER).Index).Value = nextNo
    r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_ACTIVE).Index).Value = True

    GetOrCreateItem = newCode

    mod_Logging.WriteLog "INFO", "mod_Settings", "GetOrCreateItem", _
        "新規評価項目を作成しました: " & newCode & " (" & displayNameHint & ")"
End Function

' 項目コードに対応する表示名を返す（未登録ならコードをそのまま返す）。
Public Function GetItemDisplayName(ByVal itemCode As String) As String
    Dim r As ListRow
    Set r = FindItemRow(itemCode)
    If r Is Nothing Then
        GetItemDisplayName = itemCode
    Else
        Dim tbl As ListObject
        Set tbl = r.Parent
        GetItemDisplayName = CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_DISPNAME).Index).Value)
    End If
End Function

' 項目コードに対応する重み(%)を返す（未登録は100を既定とする）。
Public Function GetItemWeight(ByVal itemCode As String) As Double
    Dim r As ListRow
    Set r = FindItemRow(itemCode)
    If r Is Nothing Then
        GetItemWeight = 100
    Else
        Dim tbl As ListObject
        Set tbl = r.Parent
        Dim v As Variant
        v = r.Range(1, tbl.ListColumns(mod_Common.COL_ITEM_WEIGHT).Index).Value
        If IsNumeric(v) Then
            GetItemWeight = CDbl(v)
        Else
            GetItemWeight = 100
        End If
    End If
End Function

' 「設定」シートの重み・表示名が編集された際に呼び出す。
' 総合評価（T_Summary）を全件再計算し、ログに記録する。
' ボタン操作または Worksheet_Change から呼び出される想定。
Public Sub ApplyItemSettingsChanged()
    On Error GoTo ErrHandler
    mod_Common.BeginBusy
    mod_Database.RecalcSummary   ' 引数省略時は全件再計算
    mod_Common.EndBusy
    mod_Logging.WriteLog "INFO", "mod_Settings", "ApplyItemSettingsChanged", _
        "評価項目の設定変更に伴い総合評価を再計算しました。"
    Exit Sub
ErrHandler:
    mod_Common.EndBusy
    mod_Common.HandleError "mod_Settings", "ApplyItemSettingsChanged"
End Sub

'------------------------------------------------------------------
' T_Templates（列マッピングテンプレート） アクセサ
' マッピングは "元見出し1=項目コード1;元見出し2=項目コード2;..." の
' 単純な区切り文字列として保存する（JSON等の外部パーサに依存しない）。
'------------------------------------------------------------------

Public Sub SaveTemplate(ByVal templateName As String, ByVal sheetPattern As String, ByVal mappingText As String)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_TEMPLATES, mod_Common.TBL_TEMPLATES)
    If tbl Is Nothing Then Exit Sub

    Dim r As ListRow
    Dim found As Boolean
    If Not tbl.DataBodyRange Is Nothing Then
        For Each r In tbl.ListRows
            If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_NAME).Index).Value) = templateName Then
                found = True
                Exit For
            End If
        Next r
    End If
    If Not found Then Set r = tbl.ListRows.Add

    r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_NAME).Index).Value = templateName
    r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_SHEETPATTERN).Index).Value = sheetPattern
    r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_MAPPING).Index).Value = mappingText
    r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_SAVEDAT).Index).Value = mod_Common.NowText()

    mod_Logging.WriteLog "INFO", "mod_Settings", "SaveTemplate", "テンプレートを保存しました: " & templateName
End Sub

' テンプレート名一覧を返す。
Public Function GetTemplateNames() As Collection
    Dim result As New Collection
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_TEMPLATES, mod_Common.TBL_TEMPLATES)
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            Dim r As ListRow
            For Each r In tbl.ListRows
                result.Add CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_NAME).Index).Value)
            Next r
        End If
    End If
    Set GetTemplateNames = result
End Function

' テンプレートを削除する（「設定」画面のテンプレート管理から使用）。
Public Sub DeleteTemplate(ByVal templateName As String)
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_TEMPLATES, mod_Common.TBL_TEMPLATES)
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    Dim i As Long
    For i = tbl.ListRows.Count To 1 Step -1
        If CStr(tbl.ListRows(i).Range(1, tbl.ListColumns(mod_Common.COL_TPL_NAME).Index).Value) = templateName Then
            tbl.ListRows(i).Delete
            mod_Logging.WriteLog "INFO", "mod_Settings", "DeleteTemplate", "テンプレートを削除しました: " & templateName
            Exit Sub
        End If
    Next i
End Sub

' テンプレート名からマッピング文字列を取得する。存在しなければ空文字。
Public Function LoadTemplateMapping(ByVal templateName As String) As String
    Dim tbl As ListObject
    Set tbl = mod_Common.GetTableSafe(mod_Common.SH_D_TEMPLATES, mod_Common.TBL_TEMPLATES)
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function

    Dim r As ListRow
    For Each r In tbl.ListRows
        If CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_NAME).Index).Value) = templateName Then
            LoadTemplateMapping = CStr(r.Range(1, tbl.ListColumns(mod_Common.COL_TPL_MAPPING).Index).Value)
            Exit Function
        End If
    Next r
End Function
