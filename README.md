Sub OpenMyBoxFolder()
    Dim targetPath As String

    ' ユーザー名以降だけ自分の環境に合わせて変更
    targetPath = "C:\Users\" & Environ("USERNAME") & "\Box\共有フォルダ\対象フォルダ"

    If Dir(targetPath, vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません:" & vbCrLf & targetPath, vbExclamation
        Exit Sub
    End If

    Shell "explorer.exe """ & targetPath & """", vbNormalFocus
End Sub

Sub 更新()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim sql As String

    ' ── 対応表：ここだけメンテすればよい ──
    ' Array(製品IDパターン, サービス名, 合致条件)
    Dim r As Variant
    r = Array( _
        Array("ABC001", "個別対応", "条件X"), _
        Array("ABC0*",  "グループ0", "条件Y"), _
        Array("ABC*",   "ABC全般",  "条件Z") _
    )   ' ← ここに50行。狭い→広いの順に並べる

    sql = "UPDATE 作業テーブル SET サービス名 = Switch(" & _
          BuildSwitch(r, 1) & "True,'未分類'), " & _
          "合致条件 = Switch(" & _
          BuildSwitch(r, 2) & "True,'未分類')"

    db.Execute sql, dbFailOnError
    MsgBox "完了"
End Sub




Sub 更新()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim sql As String

    ' ── 対応表：ここだけメンテすればよい ──
    ' Array(製品IDパターン, サービス名, 合致条件)
    Dim r As Variant
    r = Array( _
        Array("ABC001", "個別対応", "条件X"), _
        Array("ABC0*",  "グループ0", "条件Y"), _
        Array("ABC*",   "ABC全般",  "条件Z") _
    )   ' ← ここに50行。狭い→広いの順に並べる

    sql = "UPDATE 作業テーブル SET サービス名 = Switch(" & _
          BuildSwitch(r, 1) & "True,'未分類'), " & _
          "合致条件 = Switch(" & _
          BuildSwitch(r, 2) & "True,'未分類')"

    db.Execute sql, dbFailOnError
    MsgBox "完了"
End Sub

' 対応表から Switch の中身を組み立てる（valIdx=1:サービス名, 2:合致条件）
Function BuildSwitch(r As Variant, valIdx As Integer) As String
    Dim i As Integer, s As String, op As String
    For i = LBound(r) To UBound(r)
        ' パターンに * か ? があれば Like、なければ = で比較
        If InStr(r(i)(0), "*") > 0 Or InStr(r(i)(0), "?") > 0 Then
            op = " Like '"
        Else
            op = " = '"
        End If
        s = s & "製品ID" & op & r(i)(0) & "'," & "'" & r(i)(valIdx) & "', "
    Next i
    BuildSwitch = s
End Function


Sub 初回計上月を更新()
    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim sql As String
    Dim i As Integer
    Dim cnt As Integer
    Dim 列名 As String
    Dim 年月 As String
    
    Set db = CurrentDb
    Set td = db.TableDefs("テーブル名")
    
    sql = "UPDATE テーブル名 SET 初回計上月 = "
    cnt = 0
    
    ' 100列目から170列目を順に見る（実・予の両方が対象）
    For i = 100 To 170
        列名 = td.Fields(i).Name
        年月 = Left(列名, 6)                  ' "202504実" / "202505予" → "202504" / "202505"
        sql = sql & "IIf([" & 列名 & "]>0,'" & 年月 & "',"
        cnt = cnt + 1
    Next i
    
    sql = sql & "Null" & String(cnt, ")")     ' 開いたIIfの数だけ閉じる
    
    db.Execute sql, dbFailOnError             ' ←書き込みはこの1回だけ
    
    Set td = Nothing
    Set db = Nothing
    MsgBox "完了"
End Sub





























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
