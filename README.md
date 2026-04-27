Sub CellValueIfSample()

    Dim v As Variant
    
    v = Worksheets("Sheet1").Range("C2").Value
    
    If v = "" Then
        MsgBox "空白です"
        
    ElseIf v = 0 Then
        MsgBox "0です"
        
    Else
        MsgBox "0でも空白でもありません"
    End If

End Sub




Sub FilterAndCopy_ABF()

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim lastRow As Long
    Dim filterCol As String
    Dim i As Long
    Dim pasteRow As Long
    
    '=== 設定ここから ===
    Set wsSrc = Worksheets("Sheet1")   '元データのシート名
    Set wsDst = Worksheets("Sheet2")   '転記先のシート名
    
    filterCol = "C"                    '0・空白を除外したい列
    '=== 設定ここまで ===
    
    '最終行を取得
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    
    '転記先をクリア
    wsDst.Cells.Clear
    
    '見出しを転記
    wsDst.Range("A1").Value = wsSrc.Range("A1").Value
    wsDst.Range("B1").Value = wsSrc.Range("B1").Value
    wsDst.Range("C1").Value = wsSrc.Range("F1").Value
    
    pasteRow = 2
    
    '2行目から最終行まで確認
    For i = 2 To lastRow
        
        '指定列が 0 ではなく、空白でもない場合
        If wsSrc.Cells(i, filterCol).Value <> 0 _
           And wsSrc.Cells(i, filterCol).Value <> "" Then
            
            wsDst.Cells(pasteRow, "A").Value = wsSrc.Cells(i, "A").Value
            wsDst.Cells(pasteRow, "B").Value = wsSrc.Cells(i, "B").Value
            wsDst.Cells(pasteRow, "C").Value = wsSrc.Cells(i, "F").Value
            
            pasteRow = pasteRow + 1
            
        End If
        
    Next i
    
    MsgBox "転記が完了しました。", vbInformation

End Sub



Sub SampleCase()

    Dim status As String
    status = Range("A1").Value

    Select Case status
        Case "完了"
            MsgBox "処理済みです"

        Case "未対応"
            MsgBox "これから対応します"

        Case "保留"
            MsgBox "確認が必要です"

        Case Else
            MsgBox "想定外の値です"
    End Select

End Sub
