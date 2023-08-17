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

## データの並べ替え

### Excel2007以降の並べ替え

サンプル

```vb
sub sample()
    activeworkbook.worksheets("sheet1").sort.sortfields.clear
    activeworkbook.worksheets("sheet1").sort.sortfields.add2 key:=range("d2"), _
        sorton:=xlsortonvalues, order:=xlascending,dataoption:=xlsortnormal
    with activeworkbook.worksheets("sheet1").sort
        .setrange range("a2:d9")
        .header = xlno
        .matchcase = false
        .sortmethod = xlpinyin
        .apply
    end witk
end sub
```

引数sortOnに指定できる定数

|定数             |数値|意味                              |規定値|
|-----------------|----|----------------------------------|------|
|xlSortOnValues   |0   |セル内のデータで並べ替える        |〇    |
|xlSortOnCellColor|1   |セルの背景色で並べ替える          |      |
|xlSortOnFontColor|2   |セルの文字色で並べ替える          |      |
|xlSortOnIcon     |3   |条件付き書式のアイコンで並べ替える|      |

orderに指定できる定数

|定数        |数値|意味|規定値|
|------------|----|----|------|
|xlAxcending |1   |昇順|〇    |
|xlDescending|2   |降順|      |

DataOptionに指定できる引数

|定数               |数値|意味                            |規定値|
|-------------------|----|--------------------------------|------|
|xlSortNormal       |0   |数値と文字列を別々に並べ替える  |〇    |
|xlSortTextAsNumbers|1   |文字列を数値とみなして並べ替える|      |

並べ替えの挙動を指定して並べ替える。

headerプロパティに指定できる定数

|定数   |数値|意味                      |規定値|
|-------|----|--------------------------|------|
|xlGuess|0   |excelが自動判定する       |〇    |
|xlYes  |1   |壱行目はタイトル行        |      |
|xlNo   |1   |壱行目はタイトル行ではない|      |

orientatioプロパティに指定できる定数

|定数         |数値|意味            |規定値|
|-------------|----|----------------|------|
|xlTopToBottom|1   |上下に並べ替える|〇    |
|xlLeftToRight|2   |左右に並べ替える|      |

sortmethodプロパティに指定できる定数

|定数    |数値|意味                          |規定値|
|--------|----|------------------------------|------|
|xlPinYin|1   |日本語をふりがなで並べ替える  |〇    |
|xlStroke|2   |日本語を文字コードで並べ替える|      |

コンパクトに実装する場合は以下の通り

```vb
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
```

### 7-2 excel 2003までの並べ替え

従来の並べ替えの基礎構文

```vb
並べ替えるセル範囲.sort key1, order1, header
```

例

```vb
sub sample()
    range("a1").sort key1:=range("d1"), order1:=xlascending, header:=xlyes
end sub
```

ふりがなのは、phometic.textで確認することができる。

例：

```vb
sub sample()
    msgbox range("a2").phometic.text
end sub
```

## 8 テーブルの操作

### 8－1テーブルを特定する

テーブルのセルから特定する

```vb
range("a1").listobject
```

テーブルが存在するシートから特定する場合は

```vb
対象のシート.listobjects(インデックス)
対象のシート.listobjects(テーブル名)
```

といったことが可能。

例えば、対象のシート`.ListObjects(1)`といったような形で指定が可能。

または、`Range("テーブル1")`といったような形で指定することも可能。

#### 見出し(タイトル)行を含むテーブル全体

見出し行を含むテーブル全体は、LisObject.Rangeで表される。

```vb
Range("a1").ListObject.Range.Select
```

見出しを含まないデータ行全体はListObject.DataBodyRangeで表される。

```vb
Range("a1").ListObject.DataBodyRange.Select
```

見出し行はHeaderRowRangeで表される。

```vb
Range("a1").ListObject.HeaderRowRange.Select
```

列は、ListColumn、行はListRowで表される。

まとめると以下の通り。

```vb
Sub sample()
  Range("a1").ListObject.Range.Select
  Range("a1").ListObject.DataBodyRange.Select
  Range("a1").ListObject.ListColumns(3).Range.Select
  Range("a1").ListObject.ListColumns(3).DataBodyRange.Select
  Range("a1").ListObject.ListRows(3).Range.Select
  ' ListRows(行選択)はDataBodyRangeがない。
End Sub
```

### 8-3 構造化参照を使って特定する

データを選択するときは、今まで紹介した方法以外に、構造化参照で選択する方法もある。

```vb
テーブル名[[特殊項目演算子], [列指定子]]
```

項目は、All, Data, Headers, Totalsなどがある。

```vb
Range("テーブル1[#All]")
```

### 8-4 特定のデータを操作する

テーブル内のデータを探すとき、autofilterを組み合わせる検索方法がある。

```vb
sub sample()
    range("a1").listobject.range.autofilter 2, "田中"
end sub
```

見出し行ごと別シートにコピーすることもできる。

```vb
sub sample()
    range("a1").listobject.range.autofilter 2, "新垣"
    range("a1").listobject.range.copy sheets("sheet2").range("a1")
end sub
```

見出しを含まない行をコピーする場合は、rangeをdatabodyrangeに置き換えればよい

また、これらの応用で、構造化参照を使ったコピーも可能。

ListColumnsなどを使い、特定の列だけをコピーするといった使い方も可能。

行を削除する場合は、ListColumns.delete、追加する場合はListColumns.addなどがある。

また、構造化参照を使った挿入方法もある

```vb
sub sample()
    range("テーブル1[[#Data], [数値]]").offset(0, 1) = "=[@数値]*2"
end sub
```

## 9章 エラー対策

### 9-1 エラーの種類

#### 記述エラー

構文が間違っている状態だと、コンパイルエラーが発生する

#### 論理エラー

構文は間違ってないが、配列の枠数の問題など、論理的な矛盾が発生した場合は、実行時に論理エラーがスローされる。

論理エラーには、コンパイルエラーと実行時エラーの二種類がある。

エラーの対処で最も時間がかかるのは実行エラーである。

### 9-2 エラーへの対処

#### エラーが発生したら別の処理にジャンプする

例えば、gotoを使って別の処理にジャンプすることができる

```vb
On Error GoTo ジャンプ先のラベル名

...

ジャンプ先のラベル名:
    ...
```

例

```vb
sub sample()
    On Error GoTo Error1
    sheets(2).range("a1") = 100
    mesbox "代入しました"
    exit sub
Error1:
    MsgBox "エラーが発生しました"
end sub
```

#### どんなエラーが発生したかを調べる

エラーが発生した際、どのようなエラーが発生したのかは、Errオブジェクトの中身を調べることで判明する。

Errオブジェクトでよく使われるプロパティとメソッドは以下の通り

|プロパティ/メソッド  |説明                                      |
|---------------------|------------------------------------------|
|Numberプロパティ     |エラーごとに決まっているエラー番号を返す。|
|Descriptionプロパティ|エラーの意味を表すメッセージ(文字列)を返す|
|Clearメソッド        |エラー情報をクリアする                    |

エラーのサンプル

```vb
sub sample()
    on error goto Error1
    sheets("sheet1").name = range("a1")
    exit sub
Error1:
    select case err.Number
    case 9
        msgbox "sheet1が存在しません"
    case 1004
        msgbox "同盟のシートが存在します"
    case else
        msgbox "想定していないエラーです"
    end select
end sub
```

#### 発生したエラーを無視する

発生したエラーを無視するには、resumeとnextキーワードを組み合わせることで実現可能。

例

```vb
sub sample()
    on error resume next
    activeworkbook.saveas "book1.xlsm"
    if activeworkbook.saved = true Then
        msgbox "保存されました"
    else
        msgbox "保存されていません"
    end if
end sub
```

### 9-3 データのクレンジング

#### 不正なデータを修正する

マクロがエラーになる原因は、主に対の三つがあげられる。

- コードの間違い
- 操作の間違い
- データの間違い

このうち、データの間違いをきれいな形に修正するクレンジングをここでは学習する。

#### 半角文字列と全角文字列

文字列の半角と全角を変換するには、strConv関数を使う。

文法は以下の通り

```vb
StrConv(元の文字列, 変換する文字種)
```

全角はvbWide、半角はvbNarrowを指定する

半角→全角

```vb
sub sample()
    dim i as Long
    for i = 1 to 8
        cells(i, 2) = strconv(cells(i, 1), vbwide)
    next i
end sub
```

#### 不要な文字を除去する

文字の置換は、replace関数でできる。

```vb
    replace(元の文字列, 検索文字, 置換文字)
```

例

```vb
sub sample()
    dim i as Long
    for i = 1 to 8
        cells(i, 2) = replace(cells(i, 1), "-", "")
    next i
end sub
```

#### 日付の操作

日付の操作は、dateserial関数でできる。

```vb
DateSerial(年, 月, 日)
```

例

```vb
sub sample()
    dim i as Long
    for i = 2 to 8
        cells(i, 4) = dateserial(cells(i, 1), cells(i, 2), cells(i, 3))
    next i
end sub
```

文字列の日付を直すやつ

```vb
sub sample()
    dim i as Long
    for i = 2 to 8
        cells(i, 4).numberformat = "yyyy/m/d"
        cells(i, 4).value = cells(i, 4).value
    next i
end sub
```

## 10章　デバッグ

### 10-1 デバッグとは

初級者と中級者の差は、デバッグ能力だともいわれている、VBEには、便利なデバッグ機能が多数ある。

#### 文法エラーと論理エラー

VBAでデラーとなる要因は次の二つに大別される。

- 文法エラー
- 論理エラー

文法エラーは、VBAの書式や構文を誤っているエラーである。

論理エラーは、文法や構文敵には正しいものの、プログラムとしては論理的に間違っているというミスである。

### 10-2 イミディエイトウィンドウ

デバッグ作業は一般的に、マクロがエラーで一時停止した状態(デバッグモード)で行う。

一時停止状態であるため、まだ終了していない点に注意。

この時、変数の中身などを確認するのに便利なのがイミディエイトウィンドウである。

イミディエイトウィンドウは、次のとおりである

- **表示メニュー→イミディエイトウィンドウ**

もしくは\<C-g>でひらける。

プロパティを確認する場合、例えば以下のようにして確認ができる。

```vb
?range("a1").value
```

値の確認ではなく、値を代入してみたり、実際に動作をさせるような関数の場合は、クエスチョンマークは先頭につける必要はない。

イミディエイトウィンドウに何らかの値を出力するときは、debug.printを使う。

```vb
sub sample()
    debug.print 100
end sub
```

### 10-3 マクロを一時停止する

ブレークポイントの設定するには、対象の行でF9をするか、もしくはエディタの左の細い線をクリックする。

そうすると赤くなるので、これがブレークポイントとなる。

ブレークポイントを設定した後、実行すると途中で止まる。この状態でイミディエイトウィンドウで値を調べることができる。

stopステートメントを使うことで、一時中断させることができる。

### 10-4 ステップ実行

ステップ実行する場合、F8キーでステップインする

### 10-5 デバッグモードでよく使う関数

typeを調べるときは、TypeName関数を使う

```vb
sub sample()
    TypeName(Range("A1").Value)
end sub
```

isNumeric関数で、数字であるかを調べることができる。

```vb
sub sample()
    isNumeric(Range("A1").Value)
end sub
```

ほかにも、isDateなどもある。
