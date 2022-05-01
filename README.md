# VBA_ModernList
主にExcelVBAで動作する、極めて高速かつ高機能な一方向連接リストを提供します。

以下のような特徴があります。

･参照設定を追加することなく使用可能(Windows以外では使えないメソッドがあるかもしれません)  
･リストに入れた状態でソートやユニーク化、範囲抜き出しなど、リストに求められるメソッドを持ちます。  
･ほぼ理論値と思われる速さ(DictionaryやArrayListと比較して1桁以上高速に動作します)  
･メソッドチェーンをつなげて、値の追加から加工出力まで1行で書けます(.NetならLINQと言われている機能)  
･Excelシートと連携することを考え、極めて多様な配列化を内包しています(2次元配列化/天地逆転2次元配列化など)  
･メソッドチェーン内のどこでも .DebugPrintを付けて内容を確認することが可能です。  
･StringBuilderを内包しており、高速な文字列結合が行えます。CSV化を1行でかけるメソッドもあります。  
･WokrSheetFunctionの範囲を超える計算をすることが可能です(Sum,Average,Median,Max,Min,StDevP,Mode)  

 
# ベンチマーク
VBA_ModernListは内部実装において配列が使われており、純粋な配列に近い速度を維持することが可能です。
 ![](https://raw.githubusercontent.com/es2z/VBA_ModernList/main/Img/BenchMark.png?token=GHSAT0AAAAAABUDPON5NOGBAPODAPZANNGMYTOJ2AA) 
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
