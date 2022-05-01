# VBA_ModernList
主にExcelVBAで動作する、極めて高速かつ高機能な一方向連接リストを提供します。

以下のような特徴があります。

･参照設定を追加することなく使用可能(Windows以外では使えないメソッドがあるかもしれません)

･ほぼ理論値と思われる追加と削除の速さ(Dictionaryや.Net3.5のArrayListと比較して1桁以上高速に動作します)

･リストに入れた状態でソートやユニーク化、範囲抜き出しなど、通常単方向連結リストに求められるメソッドを持ちます。

･メソッドチェーンをつなげて、値の追加から加工、出力までを1行で書くことが可能です(.NetならLINQと言われている機能です)

･Excelシートと連携することを考え、極めて多種多様な配列化を内包しています(2次元配列化/天地逆転2次元配列化など)

･メソッドチェーン内のどこでも .DebugPrintを入れることで内容を確認することが可能です

･StringBuilderを内包しており、極めて高速な文字列結合が行えます。例:シート内容のCSV化などは以下のように1行で行えます。

  List.AddRange(Range("A1:C100")).ToBuildCSV(3,VbTab,VbCr)  '3列改行、セパレータはtab、改行文字は Cr
  
･WokrSheetFunctionの範囲を超える計算をすることが可能です(Sum,Average,Median,Max,Min,StDevP,Mode_Single/Multi)

 
# DEMO
 
You can learn how to making cute physics simulations (looks retro game).
 
![](https://cpp-learning.com/wp-content/uploads/2019/05/pyxel-190505-161951.gif)
 
This animation is a "Cat playing on trampoline"!
You can get basic skills for making physics simulations.
 
# Features
 
Physics_Sim_Py used [pyxel](https://github.com/kitao/pyxel) only.
 
```python
import pyxel
```
[Pyxel](https://github.com/kitao/pyxel) is a retro game engine for Python.
You can feel free to enjoy making pixel art style physics simulations.
 
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
