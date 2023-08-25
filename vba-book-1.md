# vba自動化のすべてメモ

ProtectStructureプロパティは、ブックの保護がかかっているかの確認ができる。

シート内のデータ使用領域を選択するには、userdRange

例：Sheets(1).UserdRange.Rows.Count

最終セルを取得するには、SpecialCellsを使う。

引数に入れる定数は以下の通り

|セルタイプ        |意味          |
|------------------|--------------|
|xlCellTypeLastCell|空白セル      |
|xlCellTypeBlanks  |空白セル      |
|xlCellTypeFormulas|数式を含むセル|

途中に空白行があっても、正確に最終行を取得したい場合に便利。
