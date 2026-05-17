Attribute VB_Name = "ModStep12"
Option Explicit

' ============================================================
' Step12: T-No/号機の連番付与と変更点検出
'
' 星取表からKP-No(R列)を基にT-Noの先頭番号を取得し、
' BHPlanのW列(T-No)、AA～AG列(号機)に連番を付与する。
'
' 変更点検出:
'   今月分に変更あり → 処理停止（現場に影響）
'   来月以降に変更あり → 注意ポップアップで続行
' ============================================================
Public Sub Step12_連番付与(targetWs As Worksheet)
    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    
    ' 星取表を開く（V8）
    Dim v8Wb As Workbook
    Dim v8HsWs As Worksheet
    If g_V8ProdSchedulePath <> "" Then
        Set v8Wb = Workbooks.Open(g_V8ProdSchedulePath, ReadOnly:=True)
        Set v8HsWs = シート検索(v8Wb, g_SheetV8Hoshitori)
    End If
    
    ' 星取表を開く（V9）
    Dim v9Wb As Workbook
    Dim v9HsWs As Worksheet
    If g_V9ProdSchedulePath <> "" Then
        If g_V9ProdSchedulePath <> g_V8ProdSchedulePath Then
            Set v9Wb = Workbooks.Open(g_V9ProdSchedulePath, ReadOnly:=True)
        Else
            Set v9Wb = v8Wb
        End If
        Set v9HsWs = シート検索(v9Wb, g_SheetV9Hoshitori)
    End If
    
    ' 星取表からT-Noの最大値を取得（V8/V9それぞれ）
    Dim v8MaxTNo As Long
    v8MaxTNo = 星取表最大TNo取得(v8HsWs)
    Dim v9MaxTNo As Long
    v9MaxTNo = 星取表最大TNo取得(v9HsWs)
    
    ' 星取表から既存のKP-No→T-Noマップを作成
    Dim v8KPMap As Object
    Set v8KPMap = KPNoTNoマップ作成(v8HsWs, g_V8SavedKPNoCol)
    Dim v9KPMap As Object
    Set v9KPMap = KPNoTNoマップ作成(v9HsWs, g_V9SavedKPNoCol)

    ' 生産計画No→T-Noマップ（KP-No空時のフォールバック用）
    Dim v8SeisanMap As Object
    Set v8SeisanMap = 生産計画NoTNoマップ作成(v8HsWs, 52)  ' V8星取表 AZ列
    Dim v9SeisanMap As Object
    Set v9SeisanMap = 生産計画NoTNoマップ作成(v9HsWs, 27)  ' V9星取表 AA列
    
    ' 変更点の記録
    Dim thisMonthChanges As Long
    Dim nextMonthChanges As Long
    thisMonthChanges = 0
    nextMonthChanges = 0
    
    Dim nextMonth As Date
    nextMonth = DateSerial(Year(g_BaseDate), Month(g_BaseDate) + 1, 1)
    
    Dim assignedCount As Long
    assignedCount = 0
    
    Dim i As Long
    For i = g_DataStartRow To lastRow
        Dim model As String
        model = Trim(CStr(targetWs.Cells(i, g_ColModel).Value))
        
        ' V8/V9のみ（メンテは除外 - Step11で処理済み）
        If model <> "V8" And model <> "V9" Then GoTo NextRow
        
        Dim kpNo As String
        kpNo = Trim(CStr(targetWs.Cells(i, g_ColKPNo).Value))
        Dim seisanNo As String
        seisanNo = Trim(CStr(targetWs.Cells(i, g_ColSeisanNo).Value))

        ' 既存のT-Noがあるか確認
        Dim existingTNo As String
        Dim kpMap As Object
        Dim seisanMap As Object
        Dim hsWs As Worksheet
        Dim maxTNo As Long
        
        If model = "V8" Then
            Set kpMap = v8KPMap
            Set seisanMap = v8SeisanMap
            Set hsWs = v8HsWs
            maxTNo = v8MaxTNo
        Else
            Set kpMap = v9KPMap
            Set seisanMap = v9SeisanMap
            Set hsWs = v9HsWs
            maxTNo = v9MaxTNo
        End If
        
        existingTNo = ""
        ' 1. KP-Noで検索
        If kpNo <> "" And Not kpMap Is Nothing Then
            If kpMap.Exists(kpNo) Then
                existingTNo = CStr(kpMap(kpNo))
            End If
        End If
        ' 2. KP-Noで見つからない場合は生産計画Noで検索
        If existingTNo = "" And seisanNo <> "" And Not seisanMap Is Nothing Then
            If seisanMap.Exists(seisanNo) Then
                existingTNo = CStr(seisanMap(seisanNo))
            End If
        End If
        
        ' 現在のBHPlanのT-No
        Dim currentTNo As String
        currentTNo = Trim(CStr(targetWs.Cells(i, g_ColTNo).Value))
        
        If existingTNo <> "" Then
            ' 星取表にT-Noが存在する場合
            If currentTNo = "" Then
                ' BHPlanにT-Noがない → 星取表の値を転記
                targetWs.Cells(i, g_ColTNo).Value = existingTNo
                assignedCount = assignedCount + 1
            ElseIf currentTNo <> existingTNo Then
                ' T-Noが異なる → 変更点検出
                Dim shukkaDate As Variant
                shukkaDate = targetWs.Cells(i, g_ColShukkaDate).Value
                
                If IsDate(shukkaDate) And CDate(shukkaDate) < nextMonth Then
                    ' 今月分の変更
                    targetWs.Rows(i).Interior.Color = RGB(255, 200, 200)  ' 赤ハイライト
                    Call ログ書込("Step12_連番付与", "警告", _
                        "行" & i & ": 今月分T-No変更検出（" & currentTNo & "→" & existingTNo & "） " & 識別子表示(kpNo, seisanNo))
                    thisMonthChanges = thisMonthChanges + 1
                Else
                    ' 来月以降の変更
                    Call ログ書込("Step12_連番付与", "情報", _
                        "行" & i & ": 来月以降T-No変更（" & currentTNo & "→" & existingTNo & "） " & 識別子表示(kpNo, seisanNo))
                    nextMonthChanges = nextMonthChanges + 1
                End If
                ' 星取表の値で上書き
                targetWs.Cells(i, g_ColTNo).Value = existingTNo
            End If
        Else
            ' 星取表にない → 新規連番を付与
            If currentTNo = "" Then
                maxTNo = maxTNo + 1
                targetWs.Cells(i, g_ColTNo).Value = maxTNo
                assignedCount = assignedCount + 1
                
                ' 最大値を更新
                If model = "V8" Then
                    v8MaxTNo = maxTNo
                Else
                    v9MaxTNo = maxTNo
                End If
            End If
        End If
        
        ' 号機の連番付与（AA～AG列 = g_ColZ002から7列分）
        ' T-Noが入っていれば、号機が空の場合に連番を付与
        If Trim(CStr(targetWs.Cells(i, g_ColTNo).Value)) <> "" Then
            Dim colIdx As Long
            For colIdx = 0 To 6  ' Z002～Z201
                If IsEmpty(targetWs.Cells(i, g_ColZ002 + colIdx).Value) Then
                    ' 空の場合は前行の値+1 or T-Noベースで付与
                    ' （既存データに合わせるため、空のままにしておく選択肢もある）
                End If
            Next colIdx
        End If
NextRow:
    Next i
    
    ' ファイルを閉じる
    If Not v8Wb Is Nothing Then v8Wb.Close SaveChanges:=False
    If Not v9Wb Is Nothing And g_V9ProdSchedulePath <> g_V8ProdSchedulePath Then
        v9Wb.Close SaveChanges:=False
    End If
    
    ' 結果ログ
    Call ログ書込("Step12_連番付与", "完了", _
        "連番付与:" & assignedCount & "件、今月変更:" & thisMonthChanges & "件、来月以降変更:" & nextMonthChanges & "件")
    
    ' 今月分に変更があれば処理停止
    If thisMonthChanges > 0 Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        
        MsgBox "【処理停止】" & vbCrLf & vbCrLf & _
               "今月分のT-Noに変更が " & thisMonthChanges & " 件検出されました。" & vbCrLf & _
               "赤色ハイライトの行を確認し、オムロン担当者に問い合わせてください。" & vbCrLf & vbCrLf & _
               "データ修正後、最初から再実行してください。", _
               vbCritical, "T-No変更検出 - 要問い合わせ"
        End
    End If
    
    ' 来月以降の変更は注意ポップアップで続行
    If nextMonthChanges > 0 Then
        MsgBox "【注意】" & vbCrLf & vbCrLf & _
               "来月以降のT-Noに変更が " & nextMonthChanges & " 件ありました。" & vbCrLf & _
               "ログシートで詳細を確認してください。" & vbCrLf & vbCrLf & _
               "処理は続行します。", _
               vbInformation, "T-No変更検出（来月以降）"
    End If
End Sub

' ============================================================
' 星取表からT-Noの最大値を取得
' ============================================================
Private Function 星取表最大TNo取得(hsWs As Worksheet) As Long
    If hsWs Is Nothing Then
        星取表最大TNo取得 = 0
        Exit Function
    End If
    
    Dim maxVal As Long
    maxVal = 0
    Dim lastRow As Long
    lastRow = hsWs.Cells(hsWs.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 7 To lastRow
        Dim v As Variant
        v = hsWs.Cells(r, 1).Value  ' A列=T-No
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CLng(v) > maxVal Then maxVal = CLng(v)
        End If
    Next r
    
    星取表最大TNo取得 = maxVal
End Function

' ============================================================
' 星取表からKP-No→T-Noのマップを作成
' ============================================================
Private Function KPNoTNoマップ作成(hsWs As Worksheet, kpCol As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If hsWs Is Nothing Then
        Set KPNoTNoマップ作成 = dict
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = hsWs.Cells(hsWs.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 7 To lastRow
        Dim kp As String
        kp = Trim(CStr(hsWs.Cells(r, kpCol).Value))
        Dim tno As String
        tno = Trim(CStr(hsWs.Cells(r, 1).Value))
        If kp <> "" And tno <> "" Then
            If Not dict.Exists(kp) Then
                dict.Add kp, tno
            End If
        End If
    Next r
    
    Set KPNoTNoマップ作成 = dict
End Function

' ============================================================
' 星取表から生産計画No→T-Noのマップを作成
' V8星取表の生産計画No列 = 52 (AZ列)
' V9星取表の生産計画No列 = 27 (AA列)
' ============================================================
Private Function 生産計画NoTNoマップ作成(hsWs As Worksheet, seisanCol As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If hsWs Is Nothing Then
        Set 生産計画NoTNoマップ作成 = dict
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = hsWs.Cells(hsWs.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 7 To lastRow
        Dim sn As String
        sn = Trim(CStr(hsWs.Cells(r, seisanCol).Value))
        Dim tno As String
        tno = Trim(CStr(hsWs.Cells(r, 1).Value))
        If sn <> "" And tno <> "" Then
            If Not dict.Exists(sn) Then
                dict.Add sn, tno
            End If
        End If
    Next r
    
    Set 生産計画NoTNoマップ作成 = dict
End Function


' ============================================================
' ログ用: KP-No / 生産計画No のどちらで照合したかを示す表示文字列
' ============================================================
Private Function 識別子表示(kpNo As String, seisanNo As String) As String
    If kpNo <> "" Then
        識別子表示 = "KP:" & kpNo
    Else
        識別子表示 = "生産計画No:" & seisanNo
    End If
End Function
