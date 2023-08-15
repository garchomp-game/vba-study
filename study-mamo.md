# VBAのメモ

## 1 プロシージャ

### 1-1 ほかのプロシージャを呼び出す

プロシージャは、以下のように宣言することができる。

```vb
Sub Sample1()
  Range("A1") = 100
End Sub
```

subの外に宣言された変数をモジュールレベルの変数という

### 1-2 Functionプロシージャ

subは値を返すことができないが、functionは値を返すことができる。

例

```vb
Function Sample3()
    Sample3 = Range("A1") * 2
End Function
```

### 1-3 引数を渡す

引数は以下のように渡せる

```vb
Sub Sample5(A As Long)
  ...
End Sub
```

デフォルトでは参照渡しになっている、以下の二つは同義である。

```vb
Sub Sample(A As Long)
End Sub
Sub Sample(ByRef A As Long)
End Sub
```

値渡しをするときは、ByValを使う。

```vb
Sub Sample(ByVal A As Long)
End Sub
```

### 1-4 引数を使わずに値を共有する

参照渡し以外に、モジュールレベル変数を使うことで、値の共有が可能。

```vb
Dim A As Long

Sub Sample()
    A = 100
    Call Sample2
End Sub

Sub Sample2()
    MsgBox A
End Sub
```

## 2 変数

### 配列

```vb
sub sample()
    dim a(1) as string
    a(0) = "佐藤"
    a(1) = "山本"
    msgbox a
end sub
```

配列のインデックスは0から。通常の配列と同じ。

また、配列の宣言の仕方は

- **dim 変数名(要素数 - 1) as 要素内の型**

で宣言できる

配列の要素数を確認するとき、最小のインデックスを取得するときはLBound関数、最大の要素数を取得するにはUBound関数である。

```vb
sub sample()
    dim a(2) as string
    a(0) = "佐藤"
    a(1) = "山本"
    a(2) = "おとか"
    msgbox "最小は" & lbound(a) & "、最大は" & ubound(a) & "。"
end sub
```

### 2-2 動的配列

動的配列は、あとから要素数を指定するタイプの配列。最終的には指定しないといけない。

また、preserveを使わないと、再度要素数を変更するとき、以前の値が引き継がれなくなる。

```vb
sub sample()
    dim a() as string
    redim a(1)
    a(0) = "佐藤"
    a(1) = "山本"
    redim preserve a(2)
    a(2) = "おとか"
    msgbox "最小は" & lbound(a) & "、最大は" & ubound(a) & "。"
end sub
```

### 2-3 オブジェクト変数

オブジェクト変数の宣言の仕方も通常と同じ、しかし、格納する際に必ずsetをつけなければならない。

```vb
sub sample8()
dim a as Range
set a = range("A1")
a.font.colorindex = 3
```

ワークシート変更のサンプル

```vb
sub sample()
    dim WS1 as worksheet, WS2 as worksheet
    set WS1 = activesheet
    set WS2 = worksheets.add
    WS1.activate
    ws2.name = "合計"
end sub
```

おまけ

```vb

Sub sample()
    Dim WS1 As Worksheet, WS2 As Worksheet
    Set WS1 = ActiveSheet
    Set WS2 = Worksheets.Add
    WS1.Activate
    WS2.Name = "合計"
    Set WS1 = Nothing
    Set WS2 = Nothing
End Sub

Sub deletesheet()
    Dim ws As Worksheet
    For Each sheet In Worksheets
        If sheet.Name = "合計" Then
        Application.DisplayAlerts = False
            sheet.delete
        End If
    Next
End Sub
```

### 2-4 変数の演算

変数は、普通のプログラミングと同じ感じで演算が可能。

```vb
A = A + 100
A = A + Cells(2, 4)
etc...
```

### 2-5 文字列の結合

文字列の結合は、&を使う。

```vb
A = A & Cells(1, 1)
```

## 3 ステートメント

### exitステートメント

exitステートメントは、通常のbreak文と同じである。

```vb
sub sample()
    dim i as Long
    for i = 1 to 100
        if cells(i, 1) = "" Then
            exit sub
        end if
    next i
    msgbox "終わりました。"
end sub
```

exit forもある。

```vb
sub sample()
    dim i as Long
    for i = 1 to 100
        if cells(i, 1) = "" Then
            exit for
        end if
    next i
    msgbox "終わりました。"
end sub
```

その他、do-whileの場合はdoといったような感じで、複数のパターンがある。

### 3-2 select caseステートメント

いわゆるswich文

```vb
sub sample()
    select case range("a1").value
        case "月曜"
            msgbox 1
        case "火曜"
            msgbox 2
        case "水曜"
            msgbox 3
    end select
end sub
```

### 3-3 do...loopステートメント

do-whileと同じ。後ろに置くバージョンもある

```vb
do 条件

loop

do

loop 条件

do while 条件

loop
etc...
```

例

```vb
sub sample()
    dim i as Long
    dim a as Long
    i = 1
    do while cells(i, 1) <> ""
        a = a + cells(i, 1)
        i = i + 1
    loop
    MsgBox a
end sub
```

逆の条件として、untilも存在する。

### for each...nextステートメント

普通のfor-each

```vb
for each 変数 in グループ名
    変数を使た捜査
next 変数
```

コレクションを操作する

```vb
sub sample()
    dim wb as workbook
    for each wb in workbooks
        if wb.name = "合計.xlsx" Then
            msgbox "存在します"
        end if
    next wb
end sub
```

### 3-4 ifステートメント

if文。

```vb
sub sample()
    dim i as Long
    for i = 2 to 10
        if cells(i, 1) = "広瀬" or cells(i, 1) = "西野" or cells(i, 1) = "桜井" Then
            cells(i, 3) = cells(i, 2) * 2
        end if
    next i
end sub
```

## 4 ファイル操作

### 4-1 ブックを開く

```vb
sub openBook()
    workbooks.open "C:\Work\営業部_売上.xlsx"
end sub
```

### 4-2 ブックを保存する

```vb
sub saveBook()
    activewrokbook.saveas "C:\Work\営業部_売上.xlsx"
end sub
```

フォーマットを合わせたセーブとかも

```vb
activewrokbook.saveas "C:\Work\" & Format(Now, "yyyymmdd") & ".xlsx"
```

### 4-3 ファイルをコピーする

FileCopy コピー元ファイル, コピー先のファイル

```vb
Sub sample()
    FileCopy "C:\Users\user\Downloads\合計.xlsx", "C:\Users\user\Downloads\Book2.xlsx"
End Sub
```

なお、filecopyする際は、コピー元のファイルを閉じた状態でなければならない、
そのため、現在開いているファイルをコピー元に指定することはできない。

### フォルダーを操作する

フォルダーの操作は、以下のようなものがある。

MkDir 作成するフォルダ名

sub Sample()
    MkDir "C:\Work\2023"
end sub

なお、一気に複数の改装を作ることはできない。

## ワークシート関数

### worksheetFunctionの使い方

普段エクセルで使っている関数は、WorksheetFunction.関数名(引数)の構文で呼び出すことができる。

```vb
WorksheetFunction.Sum(Range("A1:A5"))
```

### いろいろな関数

まずはsum関数

```vb
sub sample()
    range("a6") = worksheetfunction.sum(range("a1:a5"))
end sub
```

countif/sumif関数

```vb
sub sample()
    with worksheetFunction
        range("e1") = .contif(range("a1:a6"), "佐々木")
        range("e2") = .sumif(range("a1:a6"), "佐々木", range("b1:b6"))
        range("e3") = range("e2") / range("e1")
    end with
end sub
```

※ withステートメントは、パッケージ名を省略できる機能を持っているだけで、その中に別の関数があっても問題はない。

存在確認

```vb
sub sample()
    dim a as range
    set a = range("a1:a6").find(what:="佐々木")
    if a is nothing Then
        msgbox "存在しません"
    else
        msgbox "存在します"
    end if
end sub
```

large/small関数

```vb
sub sample()
    with worksheetfunction
        range("d1") = .large(range("a2:a6"), 1)
        range("d2") = .large(range("a2:a6"), 2)
        range("d3") = .large(range("a2:a6"), 3)
    end with
end sub
```

vlookup関数

```vb
sub sample()
    range("e1") = worksheetfunction.vlookup(range("d1"), range("a2:b7"), 2, false)
end sub
```

match + index関数

```vb
sub sample()
    dim n as Long
    with worksheetfunction
        N = .match(range("d1"), range("b2:b7"), 0)
        range("e1") = .index(range("a2:a7"), N)
    end with
end sub
```

eomonth関数

```vb
sub sample()
    range("b1") = worksheetfunction.eomonth(range("a1"), 0)
end sub
```

## セルの検索とオートフィルターの操作

### セルの検索

findメソッド

|オプション名   |説明                                      |
|---------------|------------------------------------------|
|What           |検索する語句を指定します。                |
|After          |次のセルから検索開始。省略可能(左上になる)|
|LookIn         |検索対象の指定                            |
|LookAt         |完全一致検索か否か                        |
|SearchOrder    |検索の方向(右か下か)                      |
|SearchDirection|次か前か                                  |
|MatchCase      |大文字小文字区別するか                    |
|MatchByte      |半角全角区別するか                        |

LookAtには、**xlWhole**が完全一致、**xlPart**が部分位置である。

```vb
sub sample()
    dim a as range
    set a = range("a1:a8").find(what:="佐々木", lookat:=xlhwole)
    a.offset(0, 1) = 100
end sub
```

見つからなかった時の処理は、基本的には以下のとおりである。

```vb
sub sample()
    if Not A Is Nothing Then
        ~~処理
    end if
end sub
```

この仕組みを利用したサンプル

```vb
sub sample()
    dim a as Range
    set a = range("a:a").find(what:="佐々木")
    if not a is nothing Then
        a.offset(0, 1) = 100
    else
        "見つかりません"
    end if
end sub
```

### 6-2 検索結果の操作

見つかったセルを含む行を削除するには、以下のとおりである。

```vb
削除する行.Delete
```

これを利用した関数

```vb
sub sample()
    dim a as range
    set a = range("a:a").find(what:="佐々木")
    if a is nothing Then
        msgbox "見つかりません"
    else
        a.entirerow.delete
    end if
end sub
```

なお、行はrow、列はcolumnであるが、それぞれ行全体、列全体を指定する場合はentirerow,entirecolumnを指定する。

見つかったセルを起点にして、別のセルを操作する場合は、offsetを使う

```vb
起点セル.Offset(行, 列)
```

これを利用した例

```vb
sub sample()
    dim a as range
    set a = range("a:a").find(what:="石橋")
    if a is nothing Then
        msgbox "見つかりません"
    else
        a.offset(0, 1) = 1000
    end if
end sub
```

見つかったセルを含むセル範囲をコピーするには、以下の構文で可能。

```vb
コピー元のセル範囲.copy コピー先のセル
```

例

```vb
sub sample()
    dim a as range
    set a = range("a:a").find(what:="佐々木")
    if a is nothing Then
        msgbox "見つかりません"
    else
        range(a, a.end(xltoright)).copy range("E2")
    end if
end sub
```

resizeの構文は以下

```vb
range("b2").resize(4, 3)
```

これを利用した例が以下の通り

```vb
sub sample()
    dim a as range
    set a = range("a:a").find(what:="佐々木")
    if a is nothing Then
        msgbox "見つかりません"
    else
        a.resize(1, 3).copy range("e2")
    end if
end sub
```

### 6-3 オートフィルターの操作

オートフィルターの構文は以下の通り

```vb
セル.AutoFilter Field, Criteria1, Operator, Criteria2
```

使用例

```vb
sub sample()
    range("a1").autofilter field:=1, criteria1:="佐々木"
end sub
```

省略記法もある

```vb
sub sample()
    range("a1").autofilter 1, "佐々木"
end sub
```

xlAndや、xlOrなどをつなげて複数の上限を指定することができる。

```vb
sub sample()
    range("a1").autofilter 1, "佐々木", xlOr, "桜井"
end sub
```

三つ以上指定するときは、配列にして渡す

```vb
sub sample()
    dim a(2) as string
    a(0) = "佐々木"
    a(1) = "桜井"
    a(2) = "松本"
    range("a1").autofilter 1, a, xlfiltervalues
end sub
```

絞り込んだ結果をコピーする

```vb
sub sample()
    range("a1").autofilter 1, "佐々木"
    range("a1").currentregion.copy sheets("sheet2").range("a1")
end sub
```

絞り込んだ結果をカウントする

文法は以下

```vb
SUBTOTAL(集計方法, セル範囲)
```

例は以下の通り

```vb
sub sample()
    dim n as Long
    range("a1").autofilter 1, "佐々木"
    n = worksheetfunction.subtotal(3, range("a:a"))
    msgbox "佐々木は、" & N - 1 & "件あります"
end sub
```

絞り込んだ結果の列を編集する

```vb
sub sample()
    range("a1").autofilter 1, "佐々木"
    range(range("d2"), cells(rows.count, 4).end(xlUp)) = 1000
    range("a1").autofilter
end sub
```

※rows.countは、最終行を取得する読み取り専用メンバといえる。ちなみに、列の最終の取得は、columns.countである。
