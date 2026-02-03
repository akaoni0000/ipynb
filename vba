For r = 1 To UBound(src, 1)

    model = src(r, 1)
    result(r, 1) = ""   ' デフォルト

    For i = 1 To UBound(rules, 1)
        If model Like rules(i, 1) Then
            result(r, 1) = rules(i, 2)
            Exit For   ' ★最初にヒットしたルールを採用
        End If
    Next i

Next r

' BR列へ一気に書き戻し
ws.Range("BR2").Resize(UBound(result, 1), 1).Value = result
