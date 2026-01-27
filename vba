Dim summaryWs As Worksheet
Dim pasteRow As Long
Dim wb As Workbook
Dim ws As Worksheet
Dim i As Long
Dim copyRange As Range

Set summaryWs = ThisWorkbook.Sheets("Summary")
pasteRow = 1   ' 最初は1行目

For i = LBound(files) To UBound(files)
    Set wb = Workbooks.Open(files(i))
    Set ws = wb.Sheets(1)

    ws.Range("A1").CurrentRegion.AutoFilter _
        Field:=3, Criteria1:="東京"

    If ws.AutoFilter.Range.Columns(1) _
        .SpecialCells(xlCellTypeVisible).Count > 1 Then

        If pasteRow = 1 Then
            ' 最初の1回目：ヘッダ込み
            Set copyRange = ws.Range("A1").CurrentRegion _
                .SpecialCells(xlCellTypeVisible)
        Else
            ' 2回目以降：ヘッダ除外
            Set copyRange = ws.Range("A1").CurrentRegion _
                .Offset(1) _
                .SpecialCells(xlCellTypeVisible)
        End If

        copyRange.Copy summaryWs.Cells(pasteRow, 1)

        pasteRow = summaryWs.Cells(summaryWs.Rows.Count, 1) _
                        .End(xlUp).Row + 1
    End If

    ws.AutoFilterMode = False
    wb.Close SaveChanges:=False
Next i
