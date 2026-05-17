Attribute VB_Name = "ModStep06"
Option Explicit

' ============================================================
' ステップ⑥: 出荷済みデータ削除
' N列（光真ss出荷日）が当月より前の行で、
'   - R列（KP-No）が BH計画保存版（V8/V9）のKP-No列に存在する
'   - もしくは KP-Noが空の場合は B列（生産計画No）が保存版の
'     生産計画No列(V8=col52, V9=col27)に存在する
' のいずれかを満たす行を出荷済みとして削除する
' ============================================================
Public Sub Step06_出荷済みデータ削除(ws As Worksheet)
    Dim savedKPNos As Collection
    Dim savedSeisanNos As Collection
    Call 保存版識別子読み込み(savedKPNos, savedSeisanNos)

    Dim lastRow As Long
    Dim i As Long
    Dim kpNo As String
    Dim seisanNo As String
    Dim shukkaDate As Variant
    Dim deletedCount As Long
    Dim deletedByKP As Long
    Dim deletedBySeisan As Long

    deletedCount = 0
    deletedByKP = 0
    deletedBySeisan = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 下から上に向かって処理（行削除時のインデックスズレを防ぐ）
    For i = lastRow To g_DataStartRow Step -1
        shukkaDate = ws.Cells(i, g_ColShukkaDate).Value

        If IsEmpty(shukkaDate) Or CStr(shukkaDate) = "" Then GoTo NextRow
        If Not IsDate(shukkaDate) Then GoTo NextRow  ' ヘッダー行等の非日付値をスキップ

        ' 出荷日が当月より前のもの（過去分）のみチェック対象
        If CDate(shukkaDate) < g_BaseDate Then
            kpNo = 識別子を正規化(ws.Cells(i, g_ColKPNo).Value)
            seisanNo = 識別子を正規化(ws.Cells(i, g_ColSeisanNo).Value)

            If kpNo <> "" Then
                If 識別子Exists(savedKPNos, kpNo) Then
                    ws.Rows(i).Delete
                    deletedCount = deletedCount + 1
                    deletedByKP = deletedByKP + 1
                    GoTo NextRow
                End If
            End If

            ' KP-No空 or KP-No不一致 → 生産計画No でフォールバック照合
            If seisanNo <> "" Then
                If 識別子Exists(savedSeisanNos, seisanNo) Then
                    ws.Rows(i).Delete
                    deletedCount = deletedCount + 1
                    deletedBySeisan = deletedBySeisan + 1
                End If
            End If
        End If
NextRow:
    Next i

    Call ログ書込("Step06_出荷済みデータ削除", "成功", _
        deletedCount & "行を削除しました(KP-No一致:" & deletedByKP & "件、生産計画No一致:" & deletedBySeisan & "件)")
End Sub

' ============================================================
' BH計画保存版（V8/V9）からKP-Noと生産計画Noを読み込んで返す
' V8: KP-No列はg_V8SavedKPNoCol(13)、生産計画No列は52(AZ列)
' V9: KP-No列はg_V9SavedKPNoCol(10)、生産計画No列は27(AA列)
' ============================================================
Private Sub 保存版識別子読み込み(ByRef kpCol As Collection, ByRef seisanCol As Collection)
    Set kpCol = New Collection
    Set seisanCol = New Collection

    ' pathInfo: (パス, KP-No列, 生産計画No列)
    Dim pathInfo(1, 2) As Variant
    pathInfo(0, 0) = g_V8SavedPath
    pathInfo(0, 1) = g_V8SavedKPNoCol
    pathInfo(0, 2) = 52  ' V8保存版 生産計画No列 = AZ列
    pathInfo(1, 0) = g_V9SavedPath
    pathInfo(1, 1) = g_V9SavedKPNoCol
    pathInfo(1, 2) = 27  ' V9保存版 生産計画No列 = AA列

    Dim idx As Long
    Dim filePath As String
    Dim kpNoCol As Long
    Dim seisanNoCol As Long
    Dim fileExists As Boolean
    Dim dirErrNum As Long
    For idx = 0 To 1
        filePath = CStr(pathInfo(idx, 0))
        kpNoCol = CLng(pathInfo(idx, 1))
        seisanNoCol = CLng(pathInfo(idx, 2))

        If filePath = "" Then GoTo NextFile
        fileExists = False
        dirErrNum = 0
        On Error Resume Next
        fileExists = (Dir(filePath) <> "")
        dirErrNum = Err.Number
        On Error GoTo 0
        If dirErrNum <> 0 And dirErrNum <> 52 Then
            Call ログ書込("Step06", "警告", "ファイル確認中にエラーが発生しました(Error " & dirErrNum & "): " & filePath)
        End If
        If Not fileExists Then
            Call ログ書込("Step06", "警告", "保存版ファイルが見つかりません: " & filePath)
            GoTo NextFile
        End If

        Dim wb As Workbook
        Set wb = Workbooks.Open(filePath, ReadOnly:=True)

        Dim ws As Worksheet
        For Each ws In wb.Sheets
            Dim lastRow As Long
            lastRow = ws.Cells(ws.Rows.Count, kpNoCol).End(xlUp).Row
            Dim i As Long
            For i = 2 To lastRow
                Dim kpNo As String
                kpNo = 識別子を正規化(ws.Cells(i, kpNoCol).Value)
                If kpNo <> "" Then
                    On Error Resume Next
                    kpCol.Add kpNo, kpNo
                    On Error GoTo 0
                End If

                Dim seisanNo As String
                seisanNo = 識別子を正規化(ws.Cells(i, seisanNoCol).Value)
                If seisanNo <> "" Then
                    On Error Resume Next
                    seisanCol.Add seisanNo, seisanNo
                    On Error GoTo 0
                End If
            Next i
        Next ws

        wb.Close SaveChanges:=False
NextFile:
    Next idx
End Sub

' ============================================================
' 識別子(KP-No / 生産計画No)を型に関わらず統一した文字列に正規化
' 数値型(Double/Long等)は小数点・カンマなしの整数文字列に変換
' 文字列型はTrimのみ
' ============================================================
Private Function 識別子を正規化(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        識別子を正規化 = ""
    ElseIf IsNumeric(v) Then
        識別子を正規化 = CStr(CLng(v))
    Else
        識別子を正規化 = Trim(CStr(v))
    End If
End Function

' ============================================================
' Collectionに識別子が存在するか確認
' ============================================================
Private Function 識別子Exists(col As Collection, id As String) As Boolean
    On Error Resume Next
    Dim dummy As String
    dummy = col(id)
    識別子Exists = (Err.Number = 0)
    On Error GoTo 0
End Function
