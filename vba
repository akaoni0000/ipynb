重複削除------------------------------------------------------------------------

Sub GetUniqueAColumn_Array2D()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- A列の最終行取得 ---
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' --- A列を配列で取得 ---
    Dim data As Variant
    data = ws.Range("A2:A" & lastRow).Value

    ' --- Dictionary作成 ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' --- 重複排除 ---
    Dim i As Long
    For i = 1 To UBound(data, 1)
        If data(i, 1) <> "" Then
            If Not dict.Exists(data(i, 1)) Then
                dict.Add data(i, 1), 1
            End If
        End If
    Next i

    ' --- ★重複なしデータを2次元配列に格納 ---
    Dim uniqueArr2D() As Variant
    ReDim uniqueArr2D(1 To dict.Count, 1 To 1)

    For i = 1 To dict.Count
        uniqueArr2D(i, 1) = dict.Keys()(i - 1)
    Next i

    ' --- 確認用（必要なければ削除） ---
    ws.Range("B2").Resize(UBound(uniqueArr2D, 1), 1).Value = uniqueArr2D

End Sub


辞書で検索------------------------------------------------------------------

Sub DictLookupValue()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    '--- キー：型番 / 値：分類 ---
    dict.Add "FC-AAA", "ディスク"
    dict.Add "NH-123", "ディスクパック"
    dict.Add "SC999", "バッテリ"

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim src As Variant
    src = ws.Range("A2:A" & lastRow).Value

    Dim result() As Variant
    ReDim result(1 To UBound(src, 1), 1 To 1)

    Dim i As Long
    For i = 1 To UBound(src, 1)

        If dict.Exists(src(i, 1)) Then
            result(i, 1) = dict(src(i, 1))  ' ← 取り出した「値」
        Else
            result(i, 1) = ""               ' ← ヒットしない場合
        End If

    Next i

    ws.Range("B2").Resize(UBound(result, 1), 1).Value = result

End Sub




抽出--------------------


Sub 抽出_Dictionary()

    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsOut As Worksheet
    Dim wsKey As Worksheet

    Set wbSrc = Workbooks("対象ブック.xlsx")
    Set wsSrc = wbSrc.Sheets("Sheet1")     ' A列に伝票番号
    Set wsKey = ThisWorkbook.Sheets("Key") ' 抽出したい60万件
    Set wsOut = ThisWorkbook.Sheets("結果")

    Dim lastRowSrc As Long, lastRowKey As Long
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    lastRowKey = wsKey.Cells(wsKey.Rows.Count, "A").End(xlUp).Row

    '--- 辞書作成 ---
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim keys As Variant
    keys = wsKey.Range("A1:A" & lastRowKey).Value

    Dim i As Long
    For i = 1 To UBound(keys, 1)
        If Not dict.Exists(keys(i, 1)) Then
            dict.Add keys(i, 1), True
        End If
    Next i

    '--- 元データを配列で処理 ---
    Dim src As Variant
    src = wsSrc.Range("A1").CurrentRegion.Value

    Dim result()
    ReDim result(1 To UBound(src, 1), 1 To UBound(src, 2))

    Dim r As Long, outRow As Long
    outRow = 0

    For r = 2 To UBound(src, 1)
        If dict.Exists(src(r, 1)) Then
            outRow = outRow + 1
            For i = 1 To UBound(src, 2)
                result(outRow, i) = src(r, i)
            Next i
        End If
    Next r

    '--- 出力 ---
    wsOut.Cells.Clear
    wsOut.Range("A1").Resize(outRow, UBound(src, 2)).Value = result

End Sub

列の並び替え------------------------------------------------------------------
Sub ReorderByHeader_WithMissing_FillHeader()

    Dim headers As Variant
    headers = Array("機器", "売上", "列C", "伝票番号", "列A")

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim i As Long
    Dim f As Range

    '① headers の数だけ先頭に空列を作る
    ws.Columns(1).Resize(, UBound(headers) + 1).Insert Shift:=xlToRight

    '② 左から順に配置
    For i = LBound(headers) To UBound(headers)
        Set f = ws.Rows(1).Find(headers(i), LookAt:=xlWhole)

        If Not f Is Nothing Then
            ' 見つかった列 → コピーして配置
            f.EntireColumn.Copy ws.Columns(i + 1)
            f.EntireColumn.Delete
        Else
            ' 見つからなかった列 → 見出しだけ書く
            ws.Cells(1, i + 1).Value = headers(i)
        End If
    Next i

End Sub



