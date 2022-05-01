# VBA_ModernList
主にExcelVBAで動作する、極めて高速かつ高機能な一方向連接リストを提供します。

以下のような特徴があります。

･参照設定を追加することなく使用可能
･リストに入れた状態でソートやユニーク化、範囲抜き出しなど、リストに求められるメソッドを持ちます。  
･ほぼ理論値と思われる速さ(DictionaryやArrayListと比較して1桁以上高速に動作します)  
･メソッドチェーンをつなげて、値の追加から加工出力まで1行で書けます(.NetならLINQと言われている機能)  
･Excelシートと連携することを考え、極めて多様な配列化を内包しています(2次元配列化/天地逆転2次元配列化など)  
･メソッドチェーン内のどこでも .DebugPrintを付けて内容を確認することが可能です。  
･StringBuilderを内包しており、高速な文字列結合が行えます。CSV化を1行でかけるメソッドもあります。  
･WokrSheetFunctionの範囲を超える計算をすることが可能です(Sum,Average,Median,Max,Min,StDevP,Mode)  

 
# ベンチマーク
VBA_ModernListは内部実装において配列が使われており、純粋な配列に近い速度で動作します。
 ![](https://raw.githubusercontent.com/es2z/VBA_ModernList/main/Img/BenchMark.png?token=GHSAT0AAAAAABUDPON5NOGBAPODAPZANNGMYTOJ2AA) 
# 基本操作
 
以下のようにインスタンスを作成することで、使用可能になります。
```VBA
Dim list as List:Set list = new List
```

値の追加
```VBA
List.Add(1)
call List.AddRange(Array(1,2,3,4,5))
call List.AddRange(Range("A1:A100"))
call List.Concat(List)) '別のリストの値を追加します
```

値の取得
```VBA
val = List.GetValueOfIndexTest(0) ''0番目の値を取得
arr = List.ToArray 'Arrayとしてすべての値を取得
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
str = List.To2DTransposeArrayTest(i) 'Array( 0 to i ,n/i) 形式になります。行数が指定できる感じ。
str = List.ToChunkArray(i) ' Array((n/i)(0 to i))形式になります For eachで配列を分けたい場合に
str = List.ToTransposeChunkArrayTest' Array((n/i)(0,0 to i))形式になります For each+貼付で困った際に
```

 
# Requirement
 
* Python 3.6.5
* pyxel 1.0.2
 
Environments under [Anaconda for Windows](https://www.anaconda.com/distribution/) is tested.
 
```bash
conda create -n pyxel pip python=3.6 Anaconda
activate pyxel
```
 
# Installation
 
Install Pyxel with pip command.
 
```bash
pip install pyxel
```
 
# Usage
 
Please create python code named "demo.py".
And copy &amp; paste [Day4 tutorial code](https://cpp-learning.com/pyxel_physical_sim4/).
 
Run "demo.py"
 
```bash
python demo.py
```
 
# Note
 
I don't test environments under Linux and Mac.
 
# Author
 
* Hayabusa
* R&D Center
* Twitter : https://twitter.com/Cpp_Learning
 
# License
 
"Physics_Sim_Py" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).
 
Enjoy making cute physics simulations!
 
Thank you!
