Attribute VB_Name = "ModStep13"
Option Explicit

' ============================================================
' ステップ⑬: V8/V9集計表作成
'
' BHPlan日程表からMODEL=V8/V9の行を抽出し、
' 光真出荷日(N列) × BH型式TYPE(S列) のピボット集計を作成して
' V8 Production Scheduleの「0110集計」シートのS列以降を更新する
' ============================================================
Public Sub Step13_集計表作成(targetWs As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim model As String
    Dim bhType As String
    Dim shukkaDate As Variant
    Dim suryo As Variant
    
    ' --- Phase 1: BHPlanからデータ収集 ---
    ' Dictionary: key="YYYY/MM/DD|BHType", value=数量合計
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' BH型式の一覧（列ヘッダー用）
    Dim bhTypes As Object
    Set bhTypes = CreateObject("Scripting.Dictionary")
    
    ' 日付一覧
    Dim dates As Object
    Set dates = CreateObject("Scripting.Dictionary")
    
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    For i = g_DataStartRow To lastRow
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        ' V8/V9のみ（メンテは除外）
        If model <> "V8" And model <> "V9" Then GoTo NextRow
        
        bhType = Trim(CStr(targetWs.Cells(i, g_ColBHType).Value))
        If bhType = "" Then GoTo NextRow
        
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextRow
        
        ' 当月以降のみ
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextRow
        
        suryo = targetWs.Cells(i, g_ColSuryo).Value
        If IsEmpty(suryo) Or Not IsNumeric(suryo) Then GoTo NextRow
        
        Dim dateKey As String
        dateKey = Format(CDate(shukkaDate), "YYYY/MM/DD")
        
        Dim dictKey As String
        dictKey = dateKey & "|" & bhType
        
        If dict.Exists(dictKey) Then
            dict(dictKey) = dict(dictKey) + CLng(suryo)
        Else
            dict.Add dictKey, CLng(suryo)
        End If
        
        ' BH型式一覧に追加
        If Not bhTypes.Exists(bhType) Then bhTypes.Add bhType, 0
        
        ' 日付一覧に追加
        If Not dates.Exists(dateKey) Then dates.Add dateKey, CDate(shukkaDate)
NextRow:
    Next i
    
    If dict.Count = 0 Then
        Call ログ書込("Step13_集計表作成", "完了", "集計対象データなし")
        Exit Sub
    End If
    
    ' --- Phase 2: 0110集計シートを更新 ---
    Dim psWb As Workbook
    Dim psPath As String
    psPath = g_V8ProdSchedulePath
    
    If psPath = "" Then
        Call ログ書込("Step13_集計表作成", "警告", "V8_ProductionScheduleパスが未設定です")
        Exit Sub
    End If
    
    Set psWb = Workbooks.Open(psPath)
    Dim shukeiWs As Worksheet
    Set shukeiWs = シート検索(psWb, g_SheetV8Shukei)
    ' フォールバック: シート名末尾が"集計"のシートを動的検索
    If shukeiWs Is Nothing Then
        Set shukeiWs = 集計シート動的検索(psWb)
        If Not shukeiWs Is Nothing Then
            Call ログ書込("Step13_集計表作成", "情報", _
                "集計シートを動的検索で発見: " & shukeiWs.Name)
        End If
    End If
    If shukeiWs Is Nothing Then
        Call ログ書込("Step13_集計表作成", "エラー", _
            "集計シートが見つかりません（設定: " & g_SheetV8Shukei & "、末尾""集計""のシートも無し）")
        psWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' --- S列以降のBH型式ヘッダーを読み取り/追加 ---
    Dim headerRow As Long
    headerRow = 3  ' ヘッダー行
    Dim startCol As Long
    startCol = 19  ' S列 = 19列目
    
    ' 既存ヘッダーを読み取り: col -> bhType
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")  ' bhType -> col
    Dim c As Long
    c = startCol
    Do While Trim(CStr(shukeiWs.Cells(headerRow, c).Value)) <> ""
        Dim existingType As String
        existingType = Trim(CStr(shukeiWs.Cells(headerRow, c).Value))
        colMap.Add existingType, c
        c = c + 1
    Loop
    Dim nextCol As Long
    nextCol = c  ' 次の空き列
    
    ' 新しいBH型式があれば列を追加
    Dim bt As Variant
    For Each bt In bhTypes.Keys
        If Not colMap.Exists(CStr(bt)) Then
            shukeiWs.Cells(headerRow, nextCol).Value = CStr(bt)
            colMap.Add CStr(bt), nextCol
            nextCol = nextCol + 1
        End If
    Next bt
    
    ' --- 日付行を探してデータ書き込み ---
    Dim shukeiLastRow As Long
    shukeiLastRow = shukeiWs.Cells(shukeiWs.Rows.Count, 2).End(xlUp).Row
    
    ' 既存の日付 -> 行番号マップを作成
    Dim rowMap As Object
    Set rowMap = CreateObject("Scripting.Dictionary")  ' dateKey -> row
    Dim r As Long
    For r = 4 To shukeiLastRow
        Dim cellVal As Variant
        cellVal = shukeiWs.Cells(r, 2).Value
        If IsDate(cellVal) Then
            Dim existDateKey As String
            existDateKey = Format(CDate(cellVal), "YYYY/MM/DD")
            If Not rowMap.Exists(existDateKey) Then
                rowMap.Add existDateKey, r
            End If
        End If
    Next r

    ' 当月以降の既存集計セル(S列以降)をクリア
    ' 数量1→0等の減少更新を反映させるため、書き込み前に対象範囲をリセット
    Dim clearedCells As Long
    clearedCells = 0
    For r = 4 To shukeiLastRow
        Dim rowDate As Variant
        rowDate = shukeiWs.Cells(r, 2).Value
        If IsDate(rowDate) Then
            If CDate(rowDate) >= g_BaseDate Then
                Dim cc As Long
                For cc = startCol To nextCol - 1
                    If Not IsEmpty(shukeiWs.Cells(r, cc).Value) Then
                        shukeiWs.Cells(r, cc).Value = Empty
                        clearedCells = clearedCells + 1
                    End If
                Next cc
            End If
        End If
    Next r
    If clearedCells > 0 Then
        Call ログ書込("Step13_集計表作成", "情報", _
            "当月以降の既存セル " & clearedCells & "箇所をクリアしました(減少更新対応)")
    End If
    
    ' 集計データを書き込み
    Dim writtenCount As Long
    writtenCount = 0
    Dim dk As Variant
    For Each dk In dict.Keys
        Dim parts() As String
        parts = Split(CStr(dk), "|")
        Dim dkDate As String
        Dim dkType As String
        dkDate = parts(0)
        dkType = parts(1)
        
        ' 行を特定（なければ末尾に追加）
        Dim targetRow As Long
        If rowMap.Exists(dkDate) Then
            targetRow = rowMap(dkDate)
        Else
            ' 新しい日付行を総計行の上に挿入
            ' 総計行を探す
            Dim insertRow As Long
            insertRow = shukeiLastRow + 1  ' デフォルトは末尾
            Dim sr As Long
            For sr = shukeiLastRow To 4 Step -1
                Dim bVal As String
                bVal = Trim(CStr(shukeiWs.Cells(sr, 2).Value))
                If bVal = "総計" Or InStr(bVal, "総計") > 0 Then
                    insertRow = sr
                    shukeiWs.Rows(sr).Insert Shift:=xlDown
                    shukeiLastRow = shukeiLastRow + 1
                    Exit For
                End If
            Next sr
            targetRow = insertRow
            shukeiWs.Cells(targetRow, 2).Value = CDate(dkDate)
            shukeiWs.Cells(targetRow, 1).Value = Format(CDate(dkDate), "YY/MM")
            rowMap.Add dkDate, targetRow
        End If
        
        ' 列を特定
        If colMap.Exists(dkType) Then
            Dim targetCol As Long
            targetCol = colMap(dkType)
            shukeiWs.Cells(targetRow, targetCol).Value = dict(dk)
            writtenCount = writtenCount + 1
        End If
        
        ' 罫線: 同じ月=hair(点線)、月末=double(二重線)
        Dim nextDateKey As String
        Dim isMonthEnd As Boolean
        isMonthEnd = False
        ' 次の行の日付と月を比較
        If targetRow < shukeiLastRow Then
            Dim nextB As Variant
            nextB = shukeiWs.Cells(targetRow + 1, 2).Value
            If IsDate(CDate(dkDate)) And IsDate(nextB) Then
                If Month(CDate(dkDate)) <> Month(CDate(nextB)) Then
                    isMonthEnd = True
                End If
            End If
        Else
            isMonthEnd = True
        End If
        
        Dim bc2 As Long
        For bc2 = 1 To nextCol - 1
            With shukeiWs.Cells(targetRow, bc2).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                If isMonthEnd Then
                    .Weight = xlMedium
                    .LineStyle = xlDouble
                Else
                    .Weight = xlHairline
                End If
            End With
        Next bc2
    Next dk
    
    ' 保存して閉じる
    psWb.Save
    psWb.Close SaveChanges:=False
    
    Call ログ書込("Step13_集計表作成", "完了", _
        writtenCount & "件の集計データを0110集計シートに書き込みました" & _
        "（BH型式: " & bhTypes.Count & "種類、日付: " & dates.Count & "件）")
End Sub


' ============================================================
' シート名末尾が"集計"のシートを検索(0110集計, 0412集計など対応)
' g_SheetV8Shukei で見つからない場合のフォールバック
' ============================================================
Private Function 集計シート動的検索(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Sheets
        Dim n As String
        n = Trim(ws.Name)
        If Right(n, 2) = "集計" And Len(n) >= 3 Then
            ' "XXXX集計" のような4桁数値+集計を優先
            If Len(n) = 6 Then
                If IsNumeric(Left(n, 4)) Then
                    Set 集計シート動的検索 = ws
                    Exit Function
                End If
            End If
        End If
    Next ws
    ' 末尾"集計"なら何でも(2回目走査)
    For Each ws In wb.Sheets
        If Right(Trim(ws.Name), 2) = "集計" Then
            Set 集計シート動的検索 = ws
            Exit Function
        End If
    Next ws
    Set 集計シート動的検索 = Nothing
End Function
