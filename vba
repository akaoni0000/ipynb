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

