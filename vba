' 2. B列を基準に降順 (xlDescending) でソート
    ' Header:=xlYes は「1行目は見出しなので動かさない」という指定
    ws.Range("A1").CurrentRegion.Sort _
        Key1:=ws.Range("B1"), _
        Order1:=xlDescending, _
        Header:=xlYes


Sub SplitSheetIntoCSV()
    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim lastRow As Long
    Dim i As Long, fileCount As Long
    Dim folderPath As String
    Dim fileName As String
    Dim headerRange As Range
    
    ' --- 1. 初期設定 ---
    Set wsSource = ActiveSheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    Set headerRange = wsSource.Rows(1) ' 1行目を項目名として保持
    
    ' 保存先フォルダの作成（現在のブックと同じ場所）
    folderPath = ThisWorkbook.Path & "\SplitCSV\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    ' 画面更新を停止して高速化
    Application.ScreenUpdating = False
    
    fileCount = 1
    
    ' --- 2. 5行ずつループ処理（データ開始の2行目からスタート） ---
    For i = 2 To lastRow Step 5
        ' 新しいブックを作成
        Set wbNew = Workbooks.Add
        
        ' 項目名を新しいブックの1行目にコピー
        headerRange.Copy Destination:=wbNew.Sheets(1).Rows(1)
        
        ' データを5行分コピーして2行目以降に貼り付け
        ' (i + 4 が最終行を超えないように調整)
        wsSource.Rows(i & ":" & Application.Min(i + 4, lastRow)).Copy _
            Destination:=wbNew.Sheets(1).Rows(2)
        
        ' CSVとして保存して閉じる
        fileName = "SplitData_" & fileCount & ".csv"
        
        ' アラートを一時停止して上書き確認などをスキップ
        Application.DisplayAlerts = False
        wbNew.SaveAs Filename:=folderPath & fileName, FileFormat:=xlCSV
        wbNew.Close SaveChanges:=False
        Application.DisplayAlerts = True
        
        fileCount = fileCount + 1
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "分割が完了しました！" & vbCrLf & "保存先: " & folderPath
End Sub





# パワポの画像をsvg、文字起こし------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

import vtracer
import os

# 読み込むファイルのパス（ここを書き換えてください）
input_path = "予防保守.png"
output_path = "b.svg"
output_pptx = "output.pptx"



# svgファイル出力--------------------------------------------------------------------------------------------------------------
# 最新の関数名 'py_convert_image_to_svg' を使用
vtracer.convert_image_to_svg_py(
    input_path, 
    output_path,
    colormode="color",
    hierarchical="stacked",   # "stacked"（デフォルト）にすることで穴のない重なった図形を生成しパスを減らす stacked
    mode="polygon",            # なめらかな曲線で出力 spline
    filter_speckle=45,        # 【重要】デフォルト4。20px以下の細かいオブジェクトを無視する 45
    color_precision=6,        # 【重要】デフォルト6。色の境界を大まかにする（下げるほど色がまとまる）
    layer_difference=16       # 【重要】デフォルト16。グラデーションの階層をまとめる
)

print(f"変換が完了しました: {output_path}")

# svgファイル出力--------------------------------------------------------------------------------------------------------------



# pptxファイル出力--------------------------------------------------------------------------------------------------------------
# 1. 画像のサイズを取得
img = cv2.imread(input_path)
if img is None:
    raise FileNotFoundError(f"画像が見つかりません: {input_path}")
height_px, width_px, _ = img.shape

# 2. EasyOCRで画像から文字と座標を抜き出す
print("文字を抽出しています...")
reader = easyocr.Reader(['ja', 'en'])
results = reader.readtext(input_path)

# 3. PowerPointプレゼンテーションの作成
prs = Presentation()

# スライドのサイズを画像のピクセルサイズに合わせて調整
prs.slide_width = Pt(width_px)
prs.slide_height = Pt(height_px)  # 【修正】高さを設定（スライド外へのはみ出しを防止） [1]

# 白紙のスライドレイアウトを追加
blank_slide_layout = prs.slide_layouts[6] 
slide = prs.slides.add_slide(blank_slide_layout)

# 4. 抽出した文字をテキストボックスとして配置
for (bbox, text, prob) in results:
    # バウンディングボックスの座標を取得
    (tl, tr, br, bl) = bbox
    
    # x, y座標と幅・高さを計算
    x = int(tl[0])
    y = int(tl[1])
    w = int(tr[0] - tl[0])
    h = int(bl[1] - tl[1])
    
    # pptxのテキストボックスを追加（位置とサイズを指定）
    txBox = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    
    # テキストボックスのフォーマット設定を変数に格納
    tf = txBox.text_frame
    
    # テキストボックス内に文字をセット
    tf.text = text
    
    # 【修正】文字が枠からはみ出ないようにする設定
    # ① 自動で改行（折り返し）されるのを防ぐ [2]
    tf.word_wrap = False
    
    # ② 文字が枠より大きい場合、枠に合わせて自動縮小する [4]
    # tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    # ② 文字が枠より大きい場合、枠に合わせて自動縮小しない [4]
    tf.auto_size = None
    
    # ③ フォントサイズを枠の高さに合わせて調整する [3, 5]
    # （※最初の段落(0番目)のフォントサイズとして指定するのが正しい文法です）
    pt_size = h * 0.75  # ピクセル高さをフォントサイズ(pt)に概算変換
    tf.paragraphs[0].font.size = Pt(pt_size)

# 5. PowerPointファイルとして保存
prs.save(output_pptx)
print(f"完了しました！ {output_pptx} を保存しました。")

# pptxファイル出力--------------------------------------------------------------------------------------------------------------



# パワポの画像をsvg、文字起こし ここまで------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------



# outlookのメール取得----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub GetMail()

    Dim oApp As New Outlook.Application
    
    Dim oNs As Outlook.Namespace
    Set oNs = oApp.GetNamespace("MAPI")
    
    Dim oF As Folder
    Set oF = oNs.Folders("test@it-yobi.com").Folders("受信トレイ")
    
    Dim mailLists As Items
    Set mailLists = oF.Items
    # Set mailLists = oF.Items.Restrict("[Subject]='[重要]テストメール'")
    
    mailLists.Sort "[ReceivedTime]", False


    Dim i As Long
    For i = 1 To 3 'mailLists.Count
        On Error Resume Next
        Cells(i + 1, "A").Value = mailLists.Item(i).ReceivedTime
        Cells(i + 1, "B").Value = mailLists.Item(i).SenderEmailAddress
        Cells(i + 1, "C").Value = mailLists.Item(i).Subject
        Cells(i + 1, "D").Value = mailLists.Item(i).Body
    Next i

End Sub

# outlookのメール取得----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------























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


列名を抽出----------------------------------------------
Sub GetHeaderArray()

    Dim ws As Worksheet
    Set ws = ActiveSheet   ' 必要なら Sheets("Sheet1") に変更

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim headers2() As String
    Dim i As Long

    ReDim headers2(1 To lastCol)

    For i = 1 To lastCol
        headers2(i) = CStr(ws.Cells(1, i).Value)
    Next i

    ' 動作確認用（イミディエイトウィンドウに出力）
    For i = 1 To UBound(headers2)
        Debug.Print i & " : " & headers2(i)
    Next i

End Sub



