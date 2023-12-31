' 関数のサンプル
Function functionSample()
  Sample3 = Range("A1") * 2
End Function

' 引数のサンプル
Sub Sample(A As Long)
End Sub

Sub Sample(ByRef A As Long)
End Sub

Sub Sample(ByVal A As Long)
End Sub

' 配列のサンプル
Sub arraySample()
  Dim a() As String
  ReDim a(1)
  a(0) = "佐藤"
  a(1) = "山本"
  ReDim Preserve a(2)
  a(2) = "おとか"
  MsgBox "最小は" & LBound(a) & "、最大は" & UBound(a) & "。"
End Sub

' 添字指定バージョン
Sub arraySample()
  Dim a() As String
  ReDim a(3 To 4)
  a(3) = "佐藤"
  a(4) = "山本"
  ReDim Preserve a(3 To 5)
  a(5) = "おとか"
  MsgBox "最小は" & LBound(a) & "、最大は" & UBound(a) & "。"
End Sub


' オブジェクトセットのサンプル
Sub setObjectSample()
  Dim a As Range
  Set a = Range("A1")
  a.Font.ColorIndex = 3
End Sub

' ワークシートを追加したり削除したり
Sub addWorkSheet()
  Dim WS1 As Worksheet, WS2 As Worksheet
  Set WS1 = ActiveSheet
  Set WS2 = Worksheets.Add
  WS1.Activate
  WS2.Name = "合計"
  Set WS1 = Nothing
  Set WS2 = Nothing
End Sub

Sub deleteWorkSheet()
  For Each sheet In Worksheets
    If sheet.Name = "合計" Then
      Application.DisplayAlerts = False
      sheet.delete
    End If
  Next
End Sub

' 適当なループのサンプル
Sub ForSample()
  Dim i As Long
  For i = 1 To 100
    MsgBox i
    If Cells(i, 1) = "" Then
      MsgBox "Empty!"
      Exit Sub
    End If
  Next i
  MsgBox "すべて何らかの値がありました。"
End Sub

' スイッチ文のサンプル
Sub selectCaseSample()
  Select Case Range("a1").Value
    Case "月曜"
      MsgBox 1
    Case "火曜"
      MsgBox 2
    Case "水曜"
      MsgBox 3
  End Select
End Sub

' do-whileのサンプル

Sub doWhileSample()
  Dim i As Long
  Dim a As Long
  i = 1
  Do While Cells(i, 1) <> ""
    a = a + Cells(i, 1)
    i = i + 1
  Loop
  MsgBox a
End Sub

' for-eachのサンプル

Sub forEachSample()
  Dim wb As Workbook
  For Each wb In Workbooks
    MsgBox wb.Name
    If wb.Name = "合計.xlsx" Then
      MsgBox "存在します"
    End If
  Next wb
End Sub

' ifのサンプル

Sub ifSample()
  Dim i As Long
  For i = 2 To 10
    If Cells(i, 1) = "広瀬" Or Cells(i, 1) = "西野" Or Cells(i, 1) = "桜井" Then
      Cells(i, 3) = Cells(i, 2) * 2
    End If
  Next i
End Sub

' ブックを開くときのサンプル

Sub openBook()
  Workbooks.Open "C:\Users\user\Downloads\合計.xlsx"
End Sub

' ブックの保存

Sub saveBook()
  ActiveWorkbook.SaveAs "C:\Users\user\Downloads\Book1.xlsm"
End Sub

' ファイルのコピー

Sub sample()
  FileCopy "C:\Users\user\Downloads\合計.xlsx", "C:\Users\user\Downloads\Book2.xlsx"
End Sub

' フォルダ作成
sub sample()
  filecopy "C:\work\売上.xlsx", "C:\Work\Sub\売上.xlsx"
end sub

' sumif関数
Sub sample()
  Range("a6") = WorksheetFunction.Sum(Range("a1:a5"))
End Sub

' countif/sumif関数
Sub sample()
  Dim name As String
  name = Range("e1").Value
  With WorksheetFunction
    Range("e3") = .CountIf(Range("a1:a6"), name)
    Range("e4") = .SumIf(Range("a1:a6"), name, Range("b1:b6"))
    Range("e5") = Range("e4") / Range("e3")
  End With
End Sub

' countifとnothingを使った関数
Sub sample()
  Dim a As Range
  Dim name As String
  name = Range("d1").Value
  Set a = Range("a1:a6").Find(what:=name)
  If a Is Nothing Then
    MsgBox "存在しません"
  Else
    MsgBox "存在します"
  End If
End Sub

' large/small関数
Sub sample()
  With WorksheetFunction
    Range("d1") = .Large(Range("a2:a6"), 1)
    Range("d2") = .Large(Range("a2:a6"), 2)
    Range("d3") = .Large(Range("a2:a6"), 3)
  End With
End Sub

Sub sample()
  With WorksheetFunction
    Range("d1") = .Small(Range("a2:a6"), 1)
    Range("d2") = .Small(Range("a2:a6"), 2)
    Range("d3") = .Small(Range("a2:a6"), 3)
  End With
End Sub

' vlookup
Sub sample()
  Range("e1") = WorksheetFunction.VLookup(Range("d1"), Range("a2:b7"), 2, False)
End Sub

' indexとmatchの複合
Sub sample()
  Dim n As Long
  With WorksheetFunction
    n = .Match(Range("d1"), Range("b2:b7"), 0)
    Range("e1") = .Index(Range("a2:a7"), n)
  End With
End Sub

' eomonth
Sub sample()
  Range("b1") = WorksheetFunction.EoMonth(Range("a1"), 0)
End Sub

' lookat
Sub sample()
  Dim a As Range
  Set a = Range("a1:a8").Find(what:="佐々木", lookat:=xlWhole)
  a.Offset(0, 1) = 100
End Sub

' nothingの例
Sub sample()
  Dim a As Range
  Set a = Range("a:a").Find(what:="田中")
  If Not a Is Nothing Then
    a.Offset(0, 1) = 100
  Else
    MsgBox "見つかりません"
  End If
End Sub

' deleteで行の削除
Sub sample()
  Dim a As Range
  Set a = Range("a:a").Find(what:="佐々木")
  If a Is Nothing Then
    MsgBox "見つかりません"
  Else
    a.EntireRow.delete
  End If
End Sub

' offsetの例
Sub sample()
  Dim a As Range
  Set a = Range("a:a").Find(what:="石橋")
  If a Is Nothing Then
    MsgBox "見つかりません"
  Else
    a.Offset(0, 1) = 1000
  End If
End Sub

' copyの使用例
Sub sample()
  Dim a As Range
  Set a = Range("a:a").Find(what:="佐々木")
  If a Is Nothing Then
    MsgBox "見つかりません"
  Else
    Range(a, a.End(xlToRight)).Copy Range("E2")
  End If
End Sub

' resizeの使用例
Sub sample()
  Dim a As Range
  Set a = Range("a:a").Find(what:="佐々木")
  If a Is Nothing Then
    MsgBox "見つかりません"
  Else
    a.Resize(1, 3).Copy Range("e2")
  End If
End Sub

' auto filterの使用例
Sub sample()
  Range("a1").AutoFilter field:=1, Criteria1:="佐々木"
End Sub

' 省略記法
Sub sample()
  Range("a1").AutoFilter 1, "佐々木", xlOr, "桜井"
End Sub

' 3つ以上ある場合
Sub sample()
  Dim a(2) As String
  a(0) = "佐々木"
  a(1) = "桜井"
  a(2) = "松本"
  Range("a1").AutoFilter 1, a, xlFilterValues
End Sub

' 絞り込んだ結果をコピーする
Sub sample()
  Range("a1").AutoFilter 1, "佐々木"
  Range("a1").CurrentRegion.Copy Sheets("sheet2").Range("a1")
End Sub

' 絞り込んだ結果をカウントする
Sub sample()
  Dim n As Long
  Range("a1").AutoFilter 1, "佐々木"
  n = WorksheetFunction.Subtotal(3, Range("a:a"))
  MsgBox "佐々木は、" & n - 1 & "件あります"
End Sub

' 絞り込んだ結果を編集するサンプル
Sub sample()
  Range("a1").AutoFilter 1, "佐々木"
  Range(Cells(2, 4), Cells(Rows.Count, 4).End(xlUp)) = 1000
  Range("a1").AutoFilter
End Sub

' データ並び替えの例
Sub sample()
  ActiveWorkbook.Worksheets("sheet1").Sort.SortFields.Clear
  ActiveWorkbook.Worksheets("sheet1").Sort.SortFields.Add2 Key:=Range("d2"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  With ActiveWorkbook.Worksheets("sheet1").Sort
    .SetRange Range("a2:d9")
    .Header = xlNo
    .MatchCase = False
    .SortMethod = xlPinYin
    .Apply
  End With
End Sub

' 別解
Sub sample()
  With Sheets("sheet1").Sort
    With .SortFields
      .Clear
      .Add2 _
        Key := Range("d2")
    End With
    .SetRange Range("a2:d9")
    .Header = xlNo
    .Apply
  End With
End Sub

' 従来の簡易的な並べ替え方法
Sub sample()
  Range("a1").Sort _
    key1 := Range("d1"), _
    order1 := xlAscending, _
    Header := xlYes
End Sub

' テーブル全体を選択する
Sub sample()
  Range("a1").ListObject.Range.Select
  Range("a1").ListObject.DataBodyRange.Select
  Range("a1").ListObject.ListColumns(3).Range.Select
  Range("a1").ListObject.ListColumns(3).DataBodyRange.Select
  Range("a1").ListObject.ListRows(3).Range.Select
  ' ListRows(行選択)はDataBodyRangeがない。
End Sub

' 構造化参照
Sub sample()
  Range("テーブル1[#All]").Select
End Sub


' テーブル内のデータを探す
Sub sample()
  Range("a1").ListObject.Range.AutoFilter 2, "新垣"
End Sub

' 絞り込んだ結果をコピーする
Sub sample()
  Range("a1").ListObject.Range.AutoFilter 2, "新垣"
  Range("a1").ListObject.Range.Copy Sheets("sheet2").Range("a1")
End Sub

' 列を追加するサンプル
Sub sample()
  Range("テーブル1[[#Data], [数値]]").Offset(0, 1) = "=[@数値]*2"
End Sub

' 例外処理のサンプル
Sub sample()
  On Error GoTo Error1
  Sheets(2).Range("a1") = 100
  MsgBox "代入しました"
  Exit Sub
Error1:
  MsgBox "エラーが発生しました"
End Sub

' エラー分岐など
Sub sample()
  On Error GoTo Error1
  Worksheets.Add.name = "sheet1"
  Sheets("sheet1").name = Range("a1")
  Exit Sub
Error1:
  Select Case Err.Number
    Case 9
      MsgBox "sheet1が存在しません"
    Case 1004
      MsgBox "同名のシートが存在します"
    Case Else
      MsgBox "想定していないエラーです"
  End Select
End Sub

' 発生したエラーを無視する。
Sub sample()
  On Error Resume Next
  ActiveWorkbook.SaveAs "book1.xlsm"
  If ActiveWorkbook.Saved = True Then
    MsgBox "保存されました"
  Else
    MsgBox "保存されていません"
  End If
End Sub

' 半角を全角に
Sub sample()
  Dim i As Long
  For i = 1 To 8
    Cells(i, 2) = StrConv(Cells(i, 1), vbWide)
  Next i
End Sub

' 全角を半角に
Sub sample()
  Dim i As Long
  For i = 1 To 8
    Cells(i, 2) = StrConv(Cells(i, 1), vbNarrow)
  Next i
End Sub

' 置換
Sub sample()
  Dim i As Long
  For i = 1 To 8
    Cells(i, 2) = Replace(Cells(i, 1), "-", "")
  Next i
End Sub

' 日付操作
Sub sample()
  Dim i As Long
  For i = 2 To 8
    Cells(i, 4) = DateSerial(Cells(i, 1), Cells(i, 2), Cells(i, 3))
  Next i
End Sub

' 文字列を数値に変換
Sub sample()
  Dim i As Long
  For i = 2 To 8
    Cells(i, 4).NumberFormat = "yyyy/m/d"
    Cells(i, 4).Value = Cells(i, 4).Value
  Next i
End Sub

' イミディエイトウィンドウにデバッグプリントを出力
Sub sample()
  Debug.Print 100
End Sub
