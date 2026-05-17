Attribute VB_Name = "ModStep18"
Option Explicit

' ============================================================
' ステップ⑱: V9星取表日程マスター更新
'
' BHPlan日程表からMODEL=V9のデータを抽出し、
' 星取表計算マスターの「V9星取表日程マスター」シートに転記する
'
' V9の号機: Z001,Z002,Y001,Y002 (AA～AD列 = col 27～30)
' ============================================================
Public Sub Step18_V9マスター更新(targetWs As Worksheet)
    If g_HoshitoriMasterPath = "" Then
        Call ログ書込("Step18_V9マスター更新", "警告", "星取表計算マスターパスが未設定です")
        Exit Sub
    End If
    
    Dim masterWb As Workbook
    On Error Resume Next
    Set masterWb = Workbooks.Open(g_HoshitoriMasterPath)
    On Error GoTo 0
    If masterWb Is Nothing Then
        Call ログ書込("Step18_V9マスター更新", "エラー", "ファイルを開けません: " & g_HoshitoriMasterPath)
        Exit Sub
    End If
    
    Dim masterWs As Worksheet
    Set masterWs = シート検索(masterWb, g_SheetV9Master)
    If masterWs Is Nothing Then
        Call ログ書込("Step18_V9マスター更新", "エラー", "V9星取表日程マスターシートが見つかりません（設定: " & g_SheetV9Master & "）")
        masterWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    Dim masterLastRow As Long
    masterLastRow = masterWs.Cells(masterWs.Rows.Count, 1).End(xlUp).Row
    If masterLastRow < 4 Then masterLastRow = 4
    
    ' 既存T-Noを収集
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
        If model <> "V9" Then GoTo NextRow
        
        Dim tNoVal As String
        tNoVal = Trim(CStr(targetWs.Cells(i, g_ColTNo).Value))
        
        Dim shukkaDate As Variant
        shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
        If IsEmpty(shukkaDate) Or Not IsDate(shukkaDate) Then GoTo NextRow
        If CDate(shukkaDate) < g_BaseDate Then GoTo NextRow
        
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

        ' 新規行は先にN列以降の関数を行5からコピー
        If isNewRow Then
            Dim templateRow As Long
            templateRow = 5
            Dim mLastCol As Long
            mLastCol = masterWs.Cells(templateRow, masterWs.Columns.Count).End(xlToLeft).Column
            If mLastCol >= 14 Then
                masterWs.Range(masterWs.Cells(templateRow, 14), masterWs.Cells(templateRow, mLastCol)).Copy
                masterWs.Range(masterWs.Cells(writeRow, 14), masterWs.Cells(writeRow, mLastCol)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
                Application.CutCopyMode = False
            End If
        End If

        ' V9用列マッピング
        masterWs.Cells(writeRow, 1).Value = targetWs.Cells(i, g_ColTNo).Value      ' A: T-No
        ' B～E: 号機 (V9はZ001,Z002,Y001,Y002 = AA～AD = col 27～30)
        Dim j As Long
        For j = 0 To 3
            masterWs.Cells(writeRow, 2 + j).Value = targetWs.Cells(i, g_ColZ002 + j).Value
        Next j
        masterWs.Cells(writeRow, 6).Value = targetWs.Cells(i, g_ColBHType).Value    ' F: BH型式
        masterWs.Cells(writeRow, 7).Value = targetWs.Cells(i, g_ColJunjoHakkoDate).Value  ' G: 順序確定日
        masterWs.Cells(writeRow, 8).Value = targetWs.Cells(i, g_ColShukkaDate).Value  ' H: 光真出荷日
        masterWs.Cells(writeRow, 13).Value = "・"  ' M列: 全て「・」
        
        addedCount = addedCount + 1
NextRow:
    Next i
    
    masterWb.Save
    masterWb.Close SaveChanges:=False
    
    Call ログ書込("Step18_V9マスター更新", "完了", addedCount & "行をV9星取表日程マスターに転記しました")
End Sub
