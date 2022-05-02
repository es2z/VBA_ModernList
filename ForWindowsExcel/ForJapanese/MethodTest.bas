Attribute VB_Name = "MethodTest"

'Ctrl+Gで各メソッドがどのように動作しているか確認することができます。

Private List As List, List1 As List, List2 As List, List3 As List

Public Sub AllTestExecute()
    Call AllTest
End Sub


'こういうことができるというデモです。もし実際にこういうことをやるならいくつかごとに説明変数に代入するべきだと思います。
Private Sub Demolition()

    Set List1 = New List
    Set List2 = New List
   
    Dim CSV As String
    CSV = List1.CreateEnumRange(1, 150) _
        .IntersectToList(List2.CreateSeqNumbers(40, 200)) _
        .SortByDescending _
        .Slice(20, 100) _
        .MAP("x", "floor(x*PI()*2,10)") _
        .Filter("x", "Mod(x,20)=0") _
        .DistinctToList _
        .ToBuildCSV(5, vbTab, vbCr)
       Debug.Print CSV
   
   'メソッドチェーン中のどこででも .DebugPrint を入れて内容を確認することが可能です。
   
    'やってること
    'List1と2に連続する値を追加
    '※List1は(開始数,作られる数)形式で、List2は作成される連番をn、引数をi,j,kとすると n = i to j (step k)形式で連番を作成します。
    'List1と2の積集合(両方にある値のみ残す)を作成して別のリストを作成(ListXとする)
    'ListXを降順ソート
    'ListXを範囲指定でスライスして別のリストを返す(ListYとする)
    'ListYに射影処理を行う(Evaluateを使用,射影後のリストをListZとする) ※この場合πと2を掛けて10の倍数に切り下げている
    'ListZにフィルター処理を行い別のリスト(ListA)この場合20の倍数のみにする 実装はMAPとほぼ同じ
    'ListAの重複を削除(ListBとする)
    'ListBからセパレータがtabで改行コードがCrな文字列を作成(引数なしの場合カンマとCrLf、実装はStringBuilderなのではやい!)
   
End Sub

'リストの中身を表示します。 "メソッドチェーンの式の途中にも挟むことが可能です!!"
Private Sub DebugPrintTest()
    Set List = New List
    Call List.CreateEnumRange(500000000, 3) '5億から始まる値3つを取得
    Call List.DebugPrint("Debug.Print =>  ", "#,##0", "円欲しい!!")
End Sub

'データを追加します。
Private Sub AddTest()

    Set List = New List
    For i = 0 To 3
        List.Add (i)
    Next i
   
    Call List.DebugPrint("Add => ")

End Sub

'リストを初期化します
Private Sub ClearTest()

    Set List = New List
    Call List.CreateEnumRange(1, 5)
    Call List.DebugPrint("ClearTest1=> ")
    Call List.Clear.DebugPrint("ClearTest2=> ") '中身が無いので表示されません

End Sub

'データを複数追加します。配列が対象です。 対象が1次元の場合ほぼ理論値が出ると思います。
Private Sub AddRangeTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.AddRange(Array(1, 2, 3)).DebugPrint("AddRange(1Dimension) => ")
   
    Dim buffer(1 To 2, 1 To 2)
    For i = 1 To 2
        For j = 1 To 2
            buffer(i, j) = i * j
        Next j
    Next i
   
    Call List2.AddRange(buffer).DebugPrint("AddRange(2Dimension)=>")
   
End Sub

'2つのリストを結合します。引数に入れたほうがあとに来ます
Private Sub ConcatTest()

    Set List1 = New List
    Set List2 = New List
    
    Call List1.CreateEnumRange(start:=5, Count:=5) '開始位置から1ずつ増える形で特定の個数の値を作ります。
    Call List2.CreateSeqNumbers(First:=5, Last:=100, step:=20) 'for i みたいな形式で値を作ります。
    Call List1.Concat(List2).DebugPrint("Concat=> ", "#,##0")
    
End Sub

'最初の値を取得
Private Sub FirstTest()
    Set List = New List
    Debug.Print "First=> " & List.CreateEnumRange(100, 5).First
End Sub

'最後の値を取得
Private Sub LastTest()
    Set List = New List
    Debug.Print "First=> " & List.CreateEnumRange(100, 5).Last
End Sub

'要素があるか
Private Sub AnyTest()
    Set List = New List
    Debug.Print "Any1=> " & List.Any_
    List.Add (1)
    Debug.Print "Any2=> " & List.Any_
End Sub

'要素が無いか
Private Sub NothingTest()
    Set List = New List
    Debug.Print "Nothing1=> " & List.Nothing_
    List.Add (1)
    Debug.Print "Nothing2=> " & List.Nothing_
End Sub

'すべての値が一致しているか
Private Sub SequenceEqualTest()
    Set List1 = New List
    Set List2 = New List
    
    Call List1.CreateEnumRange(5, 5)
    Call List2.CreateEnumRange(5, 5)
    Debug.Print "SequenceEqual1=> " & List1.SequenceEqual(List2)
    Debug.Print "SequenceEqual2=> " & List1.SequenceEqual(List2.Clear)
End Sub

'中身をクリアして連番を作成します for文の書き方と同じです。
Private Sub CreateSeqNumbersTest()
    Set List = New List
    Call List.CreateSeqNumbers(First:=100, Last:=500, step:=80).DebugPrint("CreateSeqNumbers=> ")
End Sub

'中身をクリアして.NetのEnumlabre.Rangeのような感じで 開始位置から1ずつ増える連番を指定個数作ります。
Private Sub CreateEnumRangeTest()
    Set List = New List
    Call List.CreateEnumRange(start:=100, Count:=5).DebugPrint("CreateEnumRange=> ")
End Sub

'インデックス値に対応する値を取得します。配列の[n]と同じです。
Private Sub GetValueOfIndexTest()
    Set List = New List
    Call List.CreateEnumRange(100, 5)
    Debug.Print "GetValueOfIndexTest=> " & List.GetValueOfIndex(3)
End Sub

'特定の範囲を抜き出して新しいリストとして返します。(minIndex<=抜き出すもの<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
Private Sub SliceTest()
    Set List = New List
    Call List.CreateEnumRange(100, 20)
    Call List.Slice(minIndex:=-50, maxIndex:=3).DebugPrint("SliceTest=> ")
End Sub

'indexに対応する要素を削除してデータを前に詰めます。効率は悪いです。
Private Sub RemoveTest()
    Set List = New List
    Call List.CreateEnumRange(1, 5)
    Call List.Remove(2).DebugPrint("Remove => ")
End Sub

'引数の数値に一致するものを全て削除してデータを前に詰めます。効率は(略
Private Sub RemoveAllTest()
    Set List = New List
    Call List.AddRange(Array(5, 3, 2, 5, 3, 2, 3, 4, 3))
    Call List.RemoveAll(3).DebugPrint("RemoveAll => ")
End Sub

'特定の範囲を削除して詰めます。(minIndex<=ここ消える<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
Private Sub RemoveRangeTest()
    Set List = New List
    Call List.AddRange(Array(5, 3, 2, 5, 3, 2, 3, 4, 3))
    Call List.RemoveRange(minIndex:=4, maxIndex:=15).DebugPrint("RemoveRange => ")
End Sub

'特定の範囲を削除して詰めます。(minIndex<=ここ消える<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
Private Sub PopTest()
    Set List = New List
    Call List.CreateEnumRange(100, 5)
    Debug.Print "Pop(出力値)=> "; List.Pop(3)
    Call List.DebugPrint("pop(残った値)=> ")
End Sub

''特定の範囲を返すと同時に削除して詰めます。(minIndex<=ここ消える<=maxIndex)
''Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
''返り値は取得した値です。
Private Sub PopRangeTest()
    Set List = New List
    Call List.CreateEnumRange(100, 6)
    Call List.PopRange(minIndex:=2, maxIndex:=4).DebugPrint("PopRange(取得値)=> ")
    Call List.DebugPrint("PopRange(残った値)=> ")
End Sub

'配列に変換して返します。 配列の要素数はデータ数に切り詰められます。
Private Sub ToArrayTest()
    Set List = New List
    For Each elm In List.CreateEnumRange(1, 5).ToArray
        Debug.Print "Toarray=> " & elm
    Next
End Sub

'(0,n)の配列に変換して返します。 配列の要素数はデータ数に切り詰められます。 縦方向の値の貼付にとても使えるかと。
Private Sub ToTransposeArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 5).ToTransposeArray
    Rem Range("A1:A5").Value = buffer  例えばこういう使い方ができます。
    For Each elm In buffer
        Debug.Print "ToTransposeArray=> " & elm; "※要ウォッチウィンドウ"
    Next
   
End Sub

'2次元配列に変換して返します。 2次元目の個数を設定できます。(列数を指定して分割できる形です)
Private Sub To2DArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 10).To2DArray(columnCount:=5)
    For Each elm In buffer
        Debug.Print "To2DArray=> " & elm; "※要ウォッチウィンドウ"
    Next
   
End Sub

'2次元配列に変換して返します。 2次元目の個数を設定できます。(列数を指定して分割できる形です)
Private Sub To2DTransposeArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 10).To2DTransposeArray(rowCount:=5)
    For Each elm In buffer
        Debug.Print "To2DTransposeArray=> " & elm; "※要ウォッチウィンドウ"
    Next
   
End Sub

'chunkSizeの個数ごとに分けた配列を配列の中に入れて返します。 例100個のデータを40で分けた場合Array(Array(40),Array(40),Array(40(後半20はempty))) 端数はemptyになる仕様です。。
Private Sub ToChunkArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToChunkArray(6)
    Debug.Print "Chunk=> "
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab) 'タブ区切りの1行にする
Next
   
End Sub

'上のに加えてChunk内の配列をTransposeして2次元配列にしてます(For(Each)を回して貼付をする際に使いやすいかと思います。)
Private Sub ToTransposeChunkArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToTransposeChunkArray(6)
    Debug.Print "ToTransposeChunkArray=> ※要ウォッチウィンドウ"
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab)
Next
   
End Sub

'chunkとは逆に、何個に分けるか決めてそこに均等に配置します。例100個のデータを3で分けた場合Array(Array(33),Array(33 ラストはempty),Array(33ラストはempty))'端数はemptyになる仕様です。。
Private Sub ToDivideArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToDivideArray(6)
    Debug.Print "ToDivideArray=> "
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab)
Next
   
End Sub

'Keyだけが設定されたDictionaryを返します。重複は削除されます。existsとか大量にする場合便利かと。Windows以外では多分動きません。
Private Sub ToHashsetTest()
   
    Set List = New List
    Dim dic: Set dic = List.CreateSeqNumbers(100, 150, 10).ToHashset

    For Each Key In dic.Keys
        Debug.Print "ToHashset=> " & Key
    Next
   
End Sub

'リスト同士で和集合(結合してユニーク化)を作成して新しいリストで返します
'もとのリストの重複値も消してしまう実装なのですがあってますか・・・?
Private Sub UnionToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 60, 10)
    Call List2.CreateSeqNumbers(0, 60, 12)
    Call List1.UnionToList(List2).DebugPrint("UnionToList=> ")
   
End Sub

'リスト同士で差集合(もともとあったものから引数にあるものを除去したもの)を作成して新しいリストで返します。
Private Sub ExceptToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 15, 2)
    Call List2.CreateSeqNumbers(0, 15, 3)
    Call List1.ExceptToList(List2).DebugPrint("ExceptToList=> ")
   
End Sub

'リスト同士で積集合(両方にあるもの)を作成して新しいリストで返します
Private Sub IntersectToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 24, 2)
    Call List2.CreateSeqNumbers(0, 24, 3)
    Call List1.IntersectToList(List2).DebugPrint("IntersectToList=> ")
   
End Sub

'オブジェクトを複製(シャローコピー)します。
Private Sub CloneTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 20)
    Call List.Clone.DebugPrint("Clone=> ")
End Sub

'指定数先頭から抜き出します。
Private Sub TakeToListTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 10)
    Call List.TakeToList(5).DebugPrint("TakeToList=> ")
End Sub

'指定数をスキップしてそれ以降を抜き出します
Private Sub SkipToListTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 10)
    Call List.SkipToList(5).DebugPrint("SkipToList=> ")
End Sub

' 値をシャッフルします。精度はいまいちかもしれません。精度が欲しければメルセンヌツイスターとか良いらしいので実装して下さい()たぶん偉い人がdllとかにしてます。
Private Sub RandamizeTest()
    Set List = New List
    Call List.CreateSeqNumbers(5, 10)
    Call List.Randamize.DebugPrint("Rndamize=> ")
End Sub

'昇順でソートします。(実装はクイックソートになっているはずです) 目玉機能の1つです。
Private Sub SortTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Randamize.Sort.DebugPrint("Sort=> ")
End Sub

'降順でソートします。
Private Sub SortByDescendingTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Randamize.SortByDescending.DebugPrint("SortByDescending=> ")
End Sub

'格納されている値を反転します。
Private Sub ReverseTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Reverse.DebugPrint("Reverse=> ")
End Sub

'値を文字列としてみなして、その対象がリストに含まれるか検査します
Private Sub StringContainsTest()
    Set List = New List
    Call List.AddRange(Array(1, 2, 3, 4, 5, 487, "ああああ"))
    Debug.Print "StringContains1 =>  " & List.StringContains("1")
    Debug.Print "StringContains2 =>  " & List.StringContains("7")
End Sub

'数値を文字列とみなして、対象文字列がリストに含まれるか正規表現で検査します。Windows以外だと動かないと思います。
Private Sub StringContains_RegExpTest()
    Set List = New List
    Call List.AddRange(Array(1, 2, 3, 4, 5, 487, "ああああ"))
    Debug.Print "StringContains_RegExp1 =>  " & List.StringContains_RegExp(".*7")
    Debug.Print "StringContains_RegExp2 =>  " & List.StringContains_RegExp("[0-9]{3}")
    Debug.Print "StringContains_RegExp3 =>  " & List.StringContains_RegExp("[亜-熙ぁ-んァ-ヶ]{3,}")
    Debug.Print "StringContains_RegExp4 =>  " & List.StringContains_RegExp("あ+")
    Debug.Print "StringContains_RegExp4 =>  " & List.StringContains_RegExp("^87")
End Sub

'StringBuilderのようにコレクションを"非常に高速に"結合してStringを返します。(10万回程度で1000倍速超えるはずです) 目玉機能の1つです
'引数を指定した場合、コレクションの区切りにその文字を付与します。(CSVみたいなものが作れますが、CSV(TSVなども含む)を作るメソッドは別にあります)
Private Sub ToBuildStringTest()
    Set List = New List
    Call List.AddRange(Array("寿限無", "寿限無", "五劫", "のすりきれ", "海砂利", "水魚の", "水行末", "雲来末", "風来末"))
    Debug.Print "ToBuildString1 =>  " & List.ToBuildString()
    Debug.Print "ToBuildString2 =>  " & List.ToBuildString(",")
End Sub

'CSV(Character Separated Value )を作成します。StringBuilderのように高速に結合します。
'AddRangeと組み合わせればセルから値を取得して、CSVに変換するのが1行で書けます。
Private Sub ToBuildCSVTest()
    Set List = New List
    Call List.AddRange(Array("氏名", "年齢", "性別", "あああ", 30, "男", "いいい", 40, "女"))
    Debug.Print "ToBuildCSV =>"
    Debug.Print List.ToBuildCSV(3, ",", vbCrLf)
End Sub

'重複のない数(ユニーク値)にフィルターして新しいリストとして返します。 目玉機能の1つです
Private Sub DistinctToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 12, 2)
    Call List2.CreateSeqNumbers(0, 12, 3)
    Call List1.Concat(List2).DistinctToList.Sort.DebugPrint("DistinctToList=> ")
   
End Sub

'射影処理を行います。内部実装はEvaluate関数であるため、Excel関数が使用できます。早くはないです。
Private Sub MapTest()

    Set List = New List
   
      'この場合それぞれの値にπを掛けて10の倍数に切り下げている。
    Call List.CreateSeqNumbers(0, 5) _
        .DebugPrint("Before MAP=> ") _
        .MAP("x", "floor(x*PI()*2,10)") _
        .DebugPrint("After Map=> ")
   
End Sub

'フィルター処理を行います。内部実装はEvaluate関数であるため、Excel関数が使用できます。早くはないです。
Private Sub FilterTest()

    Set List = New List
   
      'この場合20で割ったあまりが0になるものを残します。
    Call List.CreateSeqNumbers(0, 50, 10) _
        .DebugPrint("Before filtering =>") _
        .Filter("x", "Mod(x,20)=0") _
        .DebugPrint("After filtering=> ")
   
End Sub

'数式処理です。
'組み込み関数で行ける範囲の場合はそれを使い高速に処理します。
'2次元配列に再配置すれば、組み込み関数が適用できる範囲が増えますが、遅くなる場合はそれを行いません。
'前提知識(WorksheetFunctionでは1次元配列には65535個程度の制限があったりするが、2次元は1048576個だったりそれ以上可能だったりする)
Private Sub MathematicalFunctionsTest()

    Set List1 = New List

    Call List1.CreateSeqNumbers(1, 2000000) '1億とかあまり増やしすぎると落ちるかも
    Debug.Print "Sum =>  " & List1.Math_Sum '合計
    Debug.Print "Average =>  " & List1.Math_Average '平均
    Debug.Print "Median =>  " & List1.Math_Median '中央値
    Debug.Print "Max =>  " & List1.Math_Max '最大値
    Debug.Print "Min =>  " & List1.Math_Min '最小値
    Debug.Print "StDevP =>  " & List1.Math_StDevP '標準偏差
   
    '最頻値の取得
    Set List2 = New List
    For i = 1 To 100000
        List2.Add (Int(Rnd() * 2147483647))
    Next
    
    For i = 1 To 65535
        For j = 1 To 10
            List2.Add (j)
        Next j
    Next i

    Debug.Print "ModeSingle =>  " & List2.Math_ModeSingle '最頻値(複数ある場合1つのみ)
   
    For Each buf In List2.Math_ModeMulti '最頻値(全て)
        Debug.Print "ModeMulti =>  " & buf
    Next
   
End Sub

Private Sub AllTest()

    Call DebugPrintTest
    Call AddTest
    Call ClearTest
    Call AddRangeTest
    Call ConcatTest
    Call FirstTest
    Call LastTest
    Call AnyTest
    Call NothingTest
    Call SequenceEqualTest
    Call CreateSeqNumbersTest
    Call CreateEnumRangeTest
    Call GetValueOfIndexTest
    Call SliceTest
    Call RemoveTest
    Call RemoveAllTest
    Call RemoveRangeTest
    Call PopTest
    Call PopRangeTest
    Call ToArrayTest
    Call ToTransposeArrayTest
    Call To2DArrayTest
    Call To2DTransposeArrayTest
    Call ToChunkArrayTest
    Call ToTransposeChunkArrayTest
    Call ToDivideArrayTest
    Call ToHashsetTest
    Call UnionToListTest
    Call ExceptToListTest
    Call IntersectToListTest
    Call CloneTest
    Call TakeToListTest
    Call SkipToListTest
    Call RandamizeTest
    Call SortTest
    Call SortByDescendingTest
    Call ReverseTest
    Call StringContainsTest
    Call StringContains_RegExpTest
    Call ToBuildStringTest
    Call ToBuildCSVTest
    Call DistinctToListTest
    Call MapTest
    Call FilterTest
    Call MathematicalFunctionsTest
    
End Sub



