Sub CreatePivot()

    Dim srcWs As Worksheet
    Dim ptWs As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim srcRange As Range

    ' 元データ
    Set srcWs = Worksheets("Sheet1")
    Set srcRange = srcWs.Range("A1").CurrentRegion

    ' ピボット用シート
    Set ptWs = Worksheets("Sheet2")
    ptWs.Cells.Clear

    ' ピボットキャッシュ作成
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange _
    )

    ' ピボットテーブル作成
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ptWs.Range("A1"), _
        TableName:="SamplePivot" _
    )

    ' 行フィールド
    pt.PivotFields("商品名").Orientation = xlRowField

    ' 列フィールド
    pt.PivotFields("月").Orientation = xlColumnField

    ' 値フィールド
    pt.AddDataField _
        pt.PivotFields("売上"), _
        "売上合計", _
        xlSum

End Sub
