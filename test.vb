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

Sub arraySample()
  Dim a() As String
  ReDim a(1)
  a(0) = "佐藤"
  a(1) = "山本"
  ReDim Preserve a(2)
  a(2) = "おとか"
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
