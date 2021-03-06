トップページはこちらから→ [こちら](https://github.com/es2z/VBA_ModernList/)   
リリースはこちらから→  [こちら](https://github.com/es2z/VBA_ModernList/releases)

# VBA_ModernList
ExcelVBA(Windows)で動作する、極めて高速かつ高機能な一方向連接リストです。   
※他の環境でも動くバージョンは今後追加予定です。おそらくは

VBA_ModernListは内部実装において配列が使われており、純粋な配列に近い速度で動作します。
 ![](/BenchMark.png?raw=true) 

# 特徴
･参照設定を追加することなく使用可能です。  
･ソートやユニーク化、範囲抜き出しなど、リストに求められるメソッドを持ちます。  
･DictionaryやArrayListと比較して1桁以上高速に動作します  
･メソッドチェーンをつなげて、殆どの操作を1行で繋げて書けます(.NetならLINQに相当)  
･Excelシートと向けに極めて多様な配列化を内包しています(2次元配列化/天地逆転2次元配列化など)  
･メソッドチェーン内のどこでも .DebugPrintを付けて内容を確認することが可能です。  
･範囲切り出しや 和/差/積 集合などの切り出しに対応してます。  
･StringBuilderを内包しており、高速な文字列結合が行え、CSV化を1行でかけるメソッドもあります。  
･極めて大きな範囲の数値計算を行うことも可能です(Sum,Average,Median,Max,Min,StDevP,Mode)  

# デモンストレーション
もしも望むならこういうことまでできるというデモです。
実際にやるならいくつかごとに説明変数に代入するべきだと思います。
```VBA

    Dim List1 as List:Set List1 = New List
    Dim List2 as List:Set List2 = New List
   
    Dim CSV As String
    CSV = List1.CreateEnumRange(1, 150) _
        .IntersectToList(List2.CreateSeqNumbers(40, 200)) _
        .SortByDescending _
        .Slice(20, 100) _
        .DebugPrint("変換前 => ","#,##0") _
        .MAP("x", "floor(x*PI()*2,10)") _
        .Filter("x", "Mod(x,20)=0") _
        .DebugPrint("変換後 => ","0.000","になりましたよー") _
        .DistinctToList _
        .ToBuildCSV(5, vbTab, vbCr)

    Debug.Print CSV
   
    'やってること
    'List1と2に連続する値を追加
    '※List1は(開始数,作られる数)形式
    '  List2は作成される連番をn、引数をi,j,kとすると n = i to j (step k)形式で連番を作成
    'List1と2の積集合(両方にある値のみ残す)を作成して別のリストを作成(ListXとする)
    'ListXを降順ソート
    'ListXを範囲指定でスライスして別のリストを返す(ListYとする)
    '現在格納されているすべての値を3桁カンマ区切りで列挙
    'ListYに射影処理を行う(Evaluateを使用,射影後のリストをListZとする) ※この場合πと2を掛けて10の倍数に切り下げている
    'ListZにフィルター処理を行い別のリスト(ListA)この場合20の倍数のみにする 実装はMAPとほぼ同じ
    '現在格納されているすべての値を小数点以下3桁まで列挙
    'ListAの重複を削除(ListBとする)
    'ListBからセパレータがtabで改行コードがCrな文字列を作成(引数なしの場合カンマとCrLf、実装はStringBuilderなのではやい!)
    '結果を表示
```

# 基本操作
 
以下のようにインスタンスを作成することで、使用可能になります。
```VBA
Dim list as List:Set list = new List
```  
※以下の内容はMethodTest.clsを追加して、内容を見たほうがわかりやすいかもしれません。

値の追加
```VBA
List.Add(1)
call List.AddRange(Array(1,2,3,4,5))
call List.AddRange(Range("A1:A100").Value)
call List.Concat(List)) '別のリストの値を追加します
```  

値の取得
```VBA
val = List.GetValueOfIndex(0) ''0番目の値を取得
arr = List.ToArray 'Arrayとしてすべての値を取得、通常For eachをする場合はこれを使います。
set newList = List.Slice(5,10) '5番目から10番目の値をListとして取得
set newList = List.PopRange(5,10) '5番目から10番目の値をListとして取得/その範囲を元のリストから削除
```  

ソートなど
```VBA
call List.Sort '昇順ソート
call List.SortByDescending '降順ソート
```  


ユニーク化(重複のない値にする)
```VBA
set newList = List.DistinctToList 'もとのリストは保持される。
```  

  
デバッグ
```VBA
call List.DebugPrint("先頭の内容","Formatの形式 #,##0など ","後方の内容")
```  

メソッドチェーン(殆どのメソッドは自身(List型)を返すので、そのまま次のメソッドを発行することが可能)
```VBA
Call List.CreateSeqNumbers(0, 5) _
        .DebugPrint("Before MAP=> ") _
        .MAP("x", "x*PI()") _
        .DebugPrint("After Map=> ")
        
        '0から5の連番を作成
        '内容をイミディエイトに表示
        'すべての内容にΠを掛ける
        '内容を表示
```  

文字列結合
```VBA
str = List.ToBuildString '格納されている値をすべて結合します
str = List.ToBuildString(",") '格納されている値をすべて結合します。要素の区切りとして","が追加されます)
csv = List.ToBuildSCSV(5,vbTab,vbCr) '5行改行,区切り文字タブ,改行文字列CrのCSV形式の文字列に結合します。
```  

配列化(MethodTest内のウォッチウィンドウ等で見てくれたほうが良いです)
```VBA
※nは格納要素数-1とします
str = List.ToArray ' Array(n) 形式になります
str = List.ToTransposeArray ' Array(0,n) 形式になります1列に貼付する際に便利
str = List.To2DArray(i) ' Array( n/i ,0 to i ) 形式になります。列数が指定できる感じ。
str = List.To2DTransposeArray(i) 'Array( 0 to i ,n/i) 形式になります。行数が指定できる感じ。
str = List.ToChunkArray(i) ' Array((n/i)(0 to i))形式になります For eachで配列を分けたい場合に
str = List.ToTransposeChunkArray' Array((n/i)(0,0 to i))形式になります For each+貼付で困った際に
str = List.ToDivideArray(i) ' Array((0 to i)(i/n))形式になります予め何個の配列に分けたいか決まっている際に
```  

集合
```VBA
str = List.UnionToList '和集合を作ります
str = List.ExceptToList '差集合を作ります
str = List.IntersectToList '積集合を作ります
```  
 
# ライセンス
[MIT license](https://en.wikipedia.org/wiki/MIT_License).
