# 試験対策メモ

- 1-8 function内での計算では、引数だけではなく、戻り値の関数名も最初から計算対象とすることができる。引数のデフォルトはintegerがたの場合0。
- 1-10 参照渡しする前の値を変更させたくない場合は値私、変更させる場合は参照渡しにする。
- 2-1 インデックスの開始を変える場合、dim numlist(3 to 5) as integer などとする
- 2-14 forでかっこの中に配列名が入っている場合、LBound(arr), UBound(arr)を疑うべし。
- 2-16 作問ミス発見
- 3-1 caseで1 to 9といった指定は可能。また、それ以外はswitchだとTRUEだが、vbaだとcase elseとなる。
- 3-3 case文では、if文と同じように等式が使える。等式を使う場合、select caseの対象はisに置き換え、case is >= 60 などという指定が可能。
- 3-4 case文ではカンマで区切って複数指定可能。補足として、1 to 3という指定方法は間の小数点も含まれるため、この問題では間違いとなる。
- 3-7 case文では、is >= 80の後、60 to 79だと、79.5でそれ以外の判定になってしまうため、60 to 80とすることでちゃんとカバーする必要がある。
- 3-11 do-while/loopのほかにdo/loop-whileもあるということは、当然のことながらdo/loop-untileもある。意味はrubyと同じ。
- 3-13 CStr()は、char stringの略。
- 3-20 activateはselectとほぼ同じ意味である。範囲選択されたところから、offcet(0, 1)移動するということは、範囲選択左上から一個右にずれたセルが通常選択されることになる。つまり、範囲選択の下段までは選択されない。範囲選択した範囲を足す場合は、Selection(for-each)を使う。
- 3-24 workbookのコレクションはworkbooksである。activeworkbookは、現在表示しているシート、thisworkbookは、現在VBEで開いている元シートを指す
- 3-28 何かループを抜けるときはexit ステートメントで抜けることができる。例えばexit doやexit for exit subなどである。