# VBA_ModernList
主にExcelVBAで動作する、極めて高速かつ高機能な一方向連接リストを提供します。

It has the following features

･It can be used without adding reference settings.  
･It has methods required for lists, such as sorting, uniqueness, range extraction, etc., while in a list.  
･It is almost theoretically fast (more than one order of magnitude faster than Dictionary or ArrayList).  
･It is possible to chain methods together and write everything from adding values to processing output in a single line (a feature known as LINQ in .Net).  
･Extremely diverse arrayization is included for use with Excel sheets (two-dimensional arrayization, inverted two-dimensional arrayization, etc.).  
･DebugPrint can be attached anywhere in the method chain to check the contents.  
･It includes a StringBuilder for fast string merging and a method to convert to CSV in a single line.  
･It is possible to calculate beyond the range of WokrSheetFunction (Sum,Average,Median,Max,Min,StDevP,Mode).  

Translated with www.DeepL.com/Translator (free version)

 
# ベンチマーク
VBA_ModernListは内部実装において配列が使われており、純粋な配列に近い速度で動作します。
 ![](https://github.com/es2z/VBA_ModernList/blob/main/Img/BenchMark.png?raw=true) 
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
val = List.GetValueOfIndexTest(0) ''0番目の値を取得
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
str = List.IntersectToListTest '積集合を作ります
```  
 
# ライセンス
[MIT license](https://en.wikipedia.org/wiki/MIT_License).
