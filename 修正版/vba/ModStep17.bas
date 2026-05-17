Attribute VB_Name = "ModStep17"
Option Explicit

' ============================================================
' ステップ⑰: V8星取日程マスター更新
'
' BHPlan日程表からMODEL=V8のデータを抽出し、
' 星取表計算マスターの「Ｖ８星取日程マスター」シートに転記する
'
' 列マッピング（BHPlan → マスター）:
'   W列(T-No) → A列, AA～AG列(号機) → B～H列,
'   S列(BH型式) → I列, M列(順序確定日) → J列,
'   N列(光真出荷日) → K列
'   AI列～BO列の情報も転記
' ============================================================
Public Sub Step17_V8マスター更新(targetWs As Worksheet)
    If g_HoshitoriMasterPath = "" Then
        Call ログ書込("Step17_V8マスター更新", "警告", "星取表計算マスターパスが未設定です")
        Exit Sub
    End If
    
    Dim masterWb As Workbook
    On Error Resume Next
    Set masterWb = Workbooks.Open(g_HoshitoriMasterPath)
    On Error GoTo 0
    If masterWb Is Nothing Then
        Call ログ書込("Step17_V8マスター更新", "エラー", "ファイルを開けません: " & g_HoshitoriMasterPath)
        Exit Sub
    End If
    
    Dim masterWs As Worksheet
    Set masterWs = シート検索(masterWb, g_SheetV8Master)
    If masterWs Is Nothing Then
        Call ログ書込("Step17_V8マスター更新", "エラー", "V8星取日程マスターシートが見つかりません（設定: " & g_SheetV8Master & "）")
        masterWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' BHPlanからV8データを収集
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    ' マスターの既存データ末尾を取得
    Dim masterLastRow As Long
    masterLastRow = masterWs.Cells(masterWs.Rows.Count, 1).End(xlUp).Row
    If masterLastRow < 4 Then masterLastRow = 4  ' ヘッダー行の後
    
    ' 既存のT-Noを収集（重複防止）
    Dim existingTNos As Object
    Set existingTNos = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 5 To masterLastRow
        Dim tno As Variant
        tno = masterWs.Cells(r, 1).Value
        If Not IsEmpty(tno) And CStr(tno) <> "" Then
            existingTNos(CStr(tno)) = r
        End If
    Next r
    
    Dim addedCount As Long
    addedCount = 0
    Dim nextRow As Long
    nextRow = masterLastRow + 1
    
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        If model <> "V8" Then GoTo NextRow
        
        ' T-No取得
        Dim tNoVal As String
        tNoVal = Trim(CStr(targetWs.Cells(i, g_ColTNo).Value))
        
        ' 出荷日が当月以降のみ
        Dim shukkaDate As Variant
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextRow
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextRow
        
        ' 既存のT-Noは更新、新規は追加
        Dim writeRow As Long
        Dim isNewRow As Boolean
        If tNoVal <> "" And existingTNos.Exists(tNoVal) Then
            writeRow = existingTNos(tNoVal)
            isNewRow = False
        Else
            writeRow = nextRow
            nextRow = nextRow + 1
            isNewRow = True
        End If

        ' 新規行は先にN列以降の関数を行5からコピー（M列までは下でValue上書き）
        If isNewRow Then
            Dim templateRow As Long
            templateRow = 5  ' 最上データ行をテンプレートとして利用
            Dim mLastCol As Long
            mLastCol = masterWs.Cells(templateRow, masterWs.Columns.Count).End(xlToLeft).Column
            If mLastCol >= 14 Then  ' N列以降が存在する場合のみ
                masterWs.Range(masterWs.Cells(templateRow, 14), masterWs.Cells(templateRow, mLastCol)).Copy
                masterWs.Range(masterWs.Cells(writeRow, 14), masterWs.Cells(writeRow, mLastCol)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
                Application.CutCopyMode = False
            End If
        End If

        ' 列マッピングで転記
        masterWs.Cells(writeRow, 1).Value = targetWs.Cells(i, g_ColTNo).Value      ' A: T-No
        ' B～H: 号機 (AA～AG = col 27～33)
        Dim j As Long
        For j = 0 To 6
            masterWs.Cells(writeRow, 2 + j).Value = targetWs.Cells(i, g_ColZ002 + j).Value
        Next j
        masterWs.Cells(writeRow, 9).Value = targetWs.Cells(i, g_ColBHType).Value    ' I: BH型式
        masterWs.Cells(writeRow, 10).Value = targetWs.Cells(i, g_ColJunjoHakkoDate).Value  ' J: 順序確定日
        masterWs.Cells(writeRow, 11).Value = targetWs.Cells(i, g_ColShukkaDate).Value  ' K: 光真出荷日
        masterWs.Cells(writeRow, 13).Value = "・"  ' M列: 全て「・」
        
        addedCount = addedCount + 1
NextRow:
    Next i
    
    masterWb.Save
    masterWb.Close SaveChanges:=False
    
    Call ログ書込("Step17_V8マスター更新", "完了", addedCount & "行をV8星取日程マスターに転記しました")
End Sub
