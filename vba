Sub GetFirstPostedMonth_Array()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' --- データを一気に配列へ読み込み ---
    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).Value   ' 2次元配列（1始まり）

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim ym As Long
    Dim slipNo As String

    ' --- 初回計上年月を作る ---
    For i = 1 To UBound(data, 1)
        ym = CLng(data(i, 1))     ' A列：計上年月
        slipNo = CStr(data(i, 2)) ' B列：伝票番号

        If Not dict.Exists(slipNo) Then
            dict.Add slipNo, ym
        ElseIf ym < dict(slipNo) Then
            dict(slipNo) = ym
        End If
    Next i

    ' --- 出力用配列 ---
    Dim result() As Variant
    ReDim result(1 To UBound(data, 1), 1 To 1)

    For i = 1 To UBound(data, 1)
        result(i, 1) = dict(data(i, 2))
    Next i

    ' --- 一気に書き戻し ---
    ws.Cells(1, 3).Value = "初回計上年月"
    ws.Range("C2").Resize(UBound(result, 1), 1).Value = result

End Sub
