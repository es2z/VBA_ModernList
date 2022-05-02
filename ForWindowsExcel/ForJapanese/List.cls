VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'極めて高速、多機能に動く組み込み型で使用できる単方向連接リスト(のように振る舞うもの)です。
'Sortやユニーク化、Linq様のメソッドチェーンや様々な配列に再配置したり、抜き出し、集合、StringBuilderによる連結などにも対応しています。
'メソッドチェーン中の殆どの箇所にDebugPrintを置いて状態を確認することも可能です。
'自作クラス向けがObjectListになります。

'こちらのリポジトリは以下になります。スターとかおなしゃす!!
'https://github.com/es2z/VBA_ModernList

Option Base 0
Private myArr() As Variant 'ここに値が入ります
Private myCount As Long '配列の要素数です。
Private myCapacity As Long '配列の最大サイズです。

Property Get arr(): arr = myArr: End Property 'Arrayは予約文字
Property Get Count() As Long: Count = myCount: End Property
Property Get Capacity() As Long: Capacity = myCapacity: End Property

'コンストラクタです。クラスの初期化を行います。 配列の初期サイズ決めたかったのですが、仕様上引数渡せないみたいです()
Private Sub Class_Initialize()
    myCapacity = 0
    ReDim myArr(myCapacity)
End Sub

'デストラクタです、メモリを開放します。
Private Sub Class_Terminate()
    Call Clear
    VBA.DoEvents
End Sub

'クラスを初期化します。
Public Function Clear() As list
    myCount = 0
    Call Class_Initialize
    Set Clear = Me 'メソッドチェーンの為
End Function

'デバッグ出力です。なんとメソッドチェーンの中で使えます!!
Public Function DebugPrint(Optional Prefix As String, Optional FormatStr As String, Optional Suffix As String) As list

    Dim i&
    If FormatStr = "" Then
        For i = LBound(myArr) To myCount - 1
            Debug.Print Prefix & myArr(i) & Suffix
        Next
    Else
        For i = LBound(myArr) To myCount - 1
            Debug.Print Prefix & Format(myArr(i), FormatStr) & Suffix
        Next
    End If
   
    Set DebugPrint = Me

End Function

'値をリストに追加します。
'動作速度とメモリの観点から、格納アイテム数が配列のキャパシティを超えた場合のみReDim Preserveを行いキャパシティを倍にします。
'なお、ReDim Preserveの速度は新しい動的配列を確保して値を代入→入れ替えるのとほぼ同速になります
Public Sub Add(val)

    If myCount > myCapacity Then
        myCapacity = (myCapacity + 1) * 2 - 1
        ReDim Preserve myArr(myCapacity)
    End If

    myArr(myCount) = val
    myCount = myCount + 1

End Sub

'配列を対象とした高速な範囲追加を行います。 予め配列範囲を確保してそこに値を入れていきます。
'ExcelならList.AddRange(Range("A1:B100").Value)のような使い方も便利です。
Public Function AddRange(args As Variant) As list
   
   '次元数の判定
   Dim dimensionsCount&: dimensionsCount = 1
   Dim buffer&
On Error GoTo err
    While True
        buffer = UBound(args, dimensionsCount)
        dimensionsCount = dimensionsCount + 1
    Wend
err:
    dimensionsCount = dimensionsCount - 1

    '範囲を確保
    Dim insertPoint As Long: insertPoint = myCount
    
    If dimensionsCount = 1 Then
        myCount = myCount + UBound(args) - LBound(args) + 1
    Else
        Dim additionalElmCnt&: additionalElmCnt = 1
        For i = 1 To dimensionsCount
            additionalElmCnt = additionalElmCnt * (UBound(args, i) - LBound(args, i) + 1)
        Next i
        myCount = myCount + additionalElmCnt
    End If
   
    Dim isCapacityUp As Boolean: isCapacityUp = False
    Do While myCount > myCapacity
            myCapacity = (myCapacity + 1) * 2 - 1
            isCapacityUp = True
    Loop
   
    If isCapacityUp Then ReDim Preserve myArr(myCapacity)
   
   '追加
    For Each arg In args
        myArr(insertPoint) = arg
        insertPoint = insertPoint + 1
    Next
    
    Set AddRange = Me

End Function

'2つのリストを結合します。引数に入れたほうがあとに来ます。
'ToArrayではなく Property Getで相手方の配列取ってくるのも試しましたが、異常だと思えるぐらいオーバーヘッドがあるのでどうやっても1.5倍程度遅くなるようです。
Public Function Concat(argList As list) As list

    If argList.Count = 0 Then
        Set Concat = Me
        Exit Function
    End If
   
    Call AddRange(argList.ToArray)
    Set Concat = Me

End Function

'最初の要素を取得します。
Public Function First(): First = IIf(myCount = 0, Empty, myArr(0)): End Function

'最後の要素を取得します。
Public Function Last(): Last = IIf(myCount = 0, Empty, myArr(myCount - 1)): End Function

'何か要素が格納されていればTrueを返します。
Public Function Any_(): Any_ = myCount >= 1: End Function

'何も要素が格納されていなければTrueを返します。
Public Function Nothing_(): Nothing_ = myCount = 0: End Function

'2つのリストの内容が全て一致していればTrueを返します。
Public Function SequenceEqual(compairTarget As list) As Boolean
   
    If compairTarget.Count <> Me.Count Then Exit Function
   
    Dim targetArr(): targetArr = compairTarget.ToArray
    Dim i&
    For i = 0 To Me.Count - 1
        If Me.arr(i) <> targetArr(i) Then Exit Function
    Next
   
    SequenceEqual = True
   
End Function

'中身をクリアして連番を作成します for文の書き方と同じです。
Public Function CreateSeqNumbers(First As Long, Last As Long, Optional step As Long = 1) As list

    Me.Clear
    Dim i&
    For i = First To Last Step step
        Me.Add (i)
    Next i
   
    Set CreateSeqNumbers = Me
   
End Function

'中身をクリアして.NetのEnumlabre.Rangeのような感じで連番を作成します。
Public Function CreateEnumRange(start As Long, Count As Long) As list

    Me.Clear
    Dim i&
    For i = start To start + Count - 1
        Me.Add (i)
    Next i
   
    Set CreateEnumRange = Me
   
End Function

 'インデックス値に対応する値を取得します。配列の[n]と同じです。
 Public Function GetValueOfIndex(index As Long)

    If index > myCount Then
        GetValueOfIndex = Empty
    Else
        GetValueOfIndex = myArr(index)
    End If

End Function

'特定の範囲を抜き出して新しいリストとして返します。(minIndex<=抜き出すもの<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
Public Function Slice(minIndex As Long, maxIndex As Long) As list

    minIndex = IIf(minIndex < LBound(myArr), LBound(myArr), minIndex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minIndex > maxIndex Then
        Set Slice = New list
        Exit Function
    End If

    Dim newlist As list: Set newlist = New list
    Dim i&
    For i = minIndex To maxIndex
        newlist.Add (myArr(i))
    Next i

    Set Slice = newlist

End Function

'indexに対応する要素を削除してデータを前に詰めます。効率は悪いです。
Public Function Remove(index As Long) As list

    If myCount = 0 Then Exit Function
    If index < LBound(myArr) Or UBound(myArr) < index Then Exit Function
    Dim i&
    For i = index To myCount - 1
        myArr(i) = myArr(i + 1)
    Next i

    myCount = myCount - 1
    myArr(myCount) = Empty
   
    Set Remove = Me

End Function

'引数の数値に一致するものを全て削除してデータを前に詰めます(実装は新しいリストを作って入れ替えています)
Public Function RemoveAll(val) As list

    Dim newlist As list: Set newlist = New list

    Dim buf
    Dim i&
    For i = LBound(myArr) To myCount - 1
        buf = myArr(i)
        If buf <> val Then
            newlist.Add (buf)
        End If
    Next i

    myArr = newlist.arr
    myCount = newlist.Count
    myCapacity = newlist.Capacity
   
    Set RemoveAll = Me

End Function

'特定の範囲を削除して詰めます。(minIndex<=ここ消える<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
Public Function RemoveRange(minIndex As Long, maxIndex As Long) As list

    minIndex = IIf(minIndex < LBound(myArr), LBound(myArr), minIndex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minIndex > maxIndex Then Exit Function

    Dim difference: difference = maxIndex - minIndex
    Dim i&
    For i = minIndex To myCount - 1
        Dim swapIndex: swapIndex = i + difference + 1
        If swapIndex >= myCount Then Exit For
        myArr(i) = myArr(swapIndex)
    Next i

    For j = myCount - difference - 1 To myCapacity
        myArr(j) = Empty
    Next j

    myCount = myCount - difference - 1
    Set RemoveRange = Me

End Function

'indexに対応する要素を返した上で、リストからその要素を削除します。
'返り値は取得した値です。
Public Function Pop(index As Long) As Variant

    If index < LBound(myArr) Or UBound(myArr) < index Then
        Pop = Empty: Exit Function
    End If

    Pop = myArr(index)
    Call Remove(index)

End Function

'特定の範囲を返すと同時に削除して詰めます。(minIndex<=ここ消える<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
'返り値は取得した値です。
Public Function PopRange(minIndex As Long, maxIndex As Long) As list

    minIndex = IIf(minIndex < LBound(myArr), LBound(myArr), minIndex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minIndex > maxIndex Then
        Set PopRange = New list
        Exit Function
    End If

    Dim difference: difference = maxIndex - minIndex
    Dim result As list: Set result = New list

    For index = minIndex To maxIndex
        Call result.Add(myArr(index))
    Next index

    Dim i&
    For i = minIndex To myCount - 1
        Dim swapIndex: swapIndex = i + difference + 1
        If swapIndex >= myCount Then Exit For
        myArr(i) = myArr(swapIndex)
    Next i

    For j = myCount - difference - 1 To myCapacity
        myArr(j) = Empty
    Next j

    myCount = myCount - difference - 1

    Set PopRange = result

End Function

'配列に変換して返します。 配列の要素数はデータ数に切り詰められます。
Public Function ToArray() As Variant()

    If myCount = 0 Then
        ToArray = Array(Empty)
        Exit Function
    End If
   
    Dim bufArr(): ReDim bufArr(myCount - 1)
    Dim i&
    For i = 0 To (myCount - 1)
        bufArr(i) = myArr(i)
    Next
   
    ToArray = bufArr

End Function

'2次元配列に変換して1次元要素に値を入れて返します。 配列の要素数はデータ数に切り詰められます。セルの貼付に使えます。
Public Function ToTransposeArray() As Variant()

    If myCount = 0 Then
        ToTransposeArray = Array(Empty)
        Exit Function
    End If
   
    Dim bufArr()
    ReDim bufArr(myCount - 1, 0)
    Dim i&
    For i = 0 To myCount - 1
        bufArr(i, 0) = myArr(i)
    Next

    ToTransposeArray = bufArr

End Function

'2次元配列に変換して返します。 2次元目の個数(Excelなら列列に相当)を設定できます。(複数の行列を取得してそのまま貼り付ける際などに便利です)
Public Function To2DArray(columnCount As Long) As Variant()

    If myCount = 0 Then
        To2DArray = Array(Empty)
        Exit Function
    End If

    Dim oneD&: oneD = WorksheetFunction.RoundUp(myCount / columnCount, 0) - 1
    Dim towD&: twoD = columnCount - 1
   
    Dim bufArr(): ReDim bufArr(oneD, twoD)
    Dim i&, j&, k&
   
    For i = 0 To oneD
        For j = 0 To twoD
            bufArr(i, j) = myArr(k)
            k = k + 1
        Next
    Next
   
    To2DArray = bufArr

End Function

'2次元配列に変換して返します。 1次元目の個数(Excelなら行数に相当)を設定できます。
Public Function To2DTransposeArray(rowCount As Long) As Variant()

    If myCount = 0 Then
        To2DTransposeArray = Array(Empty)
        Exit Function
    End If

    Dim oneD&: oneD = rowCount - 1
    Dim twoD&: twoD = WorksheetFunction.RoundUp(myCount / rowCount, 0) - 1
   
    Dim bufArr(): ReDim bufArr(oneD, twoD)
    Dim i&, j&, k&
   
    For i = 0 To oneD
        For j = 0 To twoD
            bufArr(i, j) = myArr(k)
            k = k + 1
        Next
    Next
   
    To2DTransposeArray = bufArr

End Function

'chunkSizeの個数ごとに分けた配列を配列の中に入れて返します。 例100個のデータを40で分けた場合Array(Array(40),Array(40),Array(40(後半20はempty)))
Public Function ToChunkArray(ChunkSize As Long) As Variant()

    If myCount = 0 Then
        ToChunkArray = Array(Empty)
        Exit Function
    End If

    Dim oneD&: oneD = WorksheetFunction.RoundUp(myCount / ChunkSize, 0) - 1
    Dim towD&: twoD = ChunkSize - 1
   
    Dim result(): ReDim result(oneD)
    Dim bufArr()
    Dim i&, j&, k&
   
    For i = 0 To oneD
        ReDim bufArr(twoD)
        For j = 0 To twoD
            If k >= myCount Then Exit For
            bufArr(j) = myArr(k)
            k = k + 1
        Next
        result(i) = bufArr
    Next
   
    ToChunkArray = result

End Function

'上のに加えてChunk内の配列をTransposeして2次元配列にしてます(For(Each)を回して貼付をする際に使いやすいかと思います。)
Public Function ToTransposeChunkArray(ChunkSize As Long) As Variant()

    If myCount = 0 Then
        ToTransposeChunkArray = Array(Empty)
        Exit Function
    End If

    Dim oneD&: oneD = WorksheetFunction.RoundUp(myCount / ChunkSize, 0) - 1
    Dim towD&: twoD = ChunkSize - 1
   
    Dim result(): ReDim result(oneD)
    Dim bufArr()
    Dim i&, j&, k&
   
    For i = 0 To oneD
        ReDim bufArr(twoD, 0)
        For j = 0 To twoD
            If k >= myCount Then Exit For
            bufArr(j, 0) = myArr(k)
            k = k + 1
        Next
        result(i) = bufArr
    Next
   
    ToTransposeChunkArray = result

End Function

'chunkとは逆に、何個に分けるか決めてそこに均等に配置します。例100個のデータを3で分けた場合Array(Array(33),Array(33 ラストはempty),Array(33ラストはempty))
Public Function ToDivideArray(partsCount As Long) As Variant()

    If myCount = 0 Then
        ToDivideArray = Array(partsCount - 1)
        Exit Function
    End If
   
    If partsCount <= 0 Then
        ToDivideArray = Array(Empty)
        Exit Function
    End If

    Dim oneD&: oneD = partsCount - 1
    Dim towD&: twoD = WorksheetFunction.RoundUp(myCount / partsCount, 0) - 1
   
    Dim result(): ReDim result(oneD)
    Dim bufArr()
    Dim i&, j&, k&
   
    For i = 0 To oneD
        ReDim bufArr(twoD)
        For j = 0 To twoD
            If k >= myCount Then Exit For
            bufArr(j) = myArr(k)
            k = k + 1
        Next
        result(i) = bufArr
    Next
   
    ToDivideArray = result

End Function

'Keyだけが設定されたDictionaryを返します。重複は削除されます。existsとか大量にする場合便利かと
Public Function ToHashset() As Object

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
   
    If myCount = 0 Then
        Set ToHashset = dic
        Exit Function
    End If
   
    For Each elm In Me.ToArray
        dic(elm) = 0
    Next
    Set ToHashset = dic
   
End Function

'リスト同士で和集合(結合してユニーク化)を作成して新しいリストで返します
Public Function UnionToList(argsList As list) As list

    Dim dic As Object: Set dic = Me.ToHashset
    Dim newlist As list: Set newlist = Me.Clone
   
    For Each elm In argsList.ToArray
        If Not dic.Exists(elm) Then newlist.Add (elm)
    Next
   
    Set UnionToList = newlist
   
End Function

'リスト同士で差集合(もともとあったものから引数にあるものを除去したもの)を作成して新しいリストで返します。
Public Function ExceptToList(argsList As list) As list

    Dim dic As Object: Set dic = argsList.ToHashset
    Dim list As list: Set list = New list
    Dim newLis As list: Set newlist = New list
   
    For Each elm In Me.ToArray
        If Not dic.Exists(elm) Then newlist.Add (elm)
    Next
   
    Set ExceptToList = newlist
   
End Function

'リスト同士で積集合(両方にあるもの)を作成して新しいリストで返します。
Public Function IntersectToList(argsList As list) As list

    Dim dic As Object: Set dic = argsList.ToHashset
    Dim list As list: Set list = New list
   
    For Each elm In Me.ToArray
        If dic.Exists(elm) Then list.Add (elm)
    Next
   
    Set IntersectToList = list
   
End Function

'オブジェクトを複製(シャローコピー)します。
Public Function Clone() As list

    Dim newlist As list: Set newlist = New list
    Set Clone = newlist.Concat(Me)

End Function

'指定数先頭から抜き出す
Public Function TakeToList(itemCount As Long) As list

    '指定数が要素数を超えている場合クローンを返す
    Dim minIndex As Long: minIndex = LBound(myArr)
    If itemCount + 1 - minIndex > myCount Then
        Set TakeToList = Me.Clone()
    Else
        Dim newlist As list: Set newlist = New list
        For i = minIndex To itemCount - 1 + minIndex
            newlist.Add (myArr(i))
        Next i
        Set TakeToList = newlist
    End If

End Function

'指定数をスキップしてそれ以降を抜き出す
Public Function SkipToList(skipCount As Long) As list

    '指定数が要素数を超えている場合nullのアイテムを返す
    If skipCount > myCount Then
        Set SkipToList = New list
    Else
        Dim newlist As list: Set newlist = New list
        For i = skipCount To myCount - 1
            newlist.Add (myArr(i))
        Next i
        Set SkipToList = newlist
    End If

End Function

'昇順でソートします。(実装はクイックソートになっているはずです)
Public Function Sort() As list

    If myCount < 1 Then Exit Function
    Call QuickSortAscending(LBound(myArr), myCount - 1)
    Set Sort = Me
   
End Function

Private Sub QuickSortAscending(ByVal lowest As Long, ByVal highest As Long)

    Dim l As Long: l = lowest
    Dim h As Long: h = highest
    Dim pivot: pivot = myArr(Int((lowest + highest) \ 2))

    Do
        Do While myArr(l) < pivot
            l = l + 1
        Loop
        Do While myArr(h) > pivot
            h = h - 1
        Loop

        If l >= h Then Exit Do

        Dim buf: buf = myArr(l)
        myArr(l) = myArr(h)
        myArr(h) = buf
        l = l + 1
        h = h - 1
    Loop

    If (lowest < l - 1) Then
        Call QuickSortAscending(lowest, l - 1)
    End If
    If (highest > h + 1) Then
        Call QuickSortAscending(h + 1, highest)
    End If

End Sub

'降順でソートします。
Public Function SortByDescending() As list

    If myCount < 1 Then Exit Function
    Call QuickSortDescending(LBound(myArr), myCount - 1)
    Set SortByDescending = Me

End Function

Private Sub QuickSortDescending(ByVal lowest As Long, ByVal highest As Long)

    Dim l As Long: l = lowest
    Dim h As Long: h = highest
    Dim pivot: pivot = myArr(Int((lowest + highest) \ 2))

    Do
        Do While myArr(l) > pivot
            l = l + 1
        Loop
        Do While myArr(h) < pivot
            h = h - 1
        Loop

        If l >= h Then Exit Do

        Dim buf: buf = myArr(l)
        myArr(l) = myArr(h)
        myArr(h) = buf
        l = l + 1
        h = h - 1
    Loop

    If (lowest < l - 1) Then
        Call QuickSortDescending(lowest, l - 1)
    End If
    If (highest > h + 1) Then
        Call QuickSortDescending(h + 1, highest)
    End If

End Sub

'格納されている値を反転します。
Public Function Reverse() As list

    If myCount < 1 Then Exit Function
    Dim maxIndxNo As Long: maxIndxNo = myCount - 1
    Dim buf(): ReDim buf(0 To maxIndxNo)

    Dim counter As Long: counter = maxIndxNo

    For i = 0 To maxIndxNo
        buf(counter) = myArr(i)
        counter = counter - 1
    Next i

    myArr = buf
   
    Set Reverse = Me

End Function

' 値をシャッフルします。精度はいまいちかもしれません。精度が欲しければメルセンヌツイスターとか良いらしいので実装して下さい(別言語でdll作るほうが現実的?)
Public Function Randamize() As list

    Dim maxIndxNo As Long: maxIndxNo = myCount - 1
    Dim i As Long, rndIndex As Variant, tmp As Variant
    For i = 0 To maxIndxNo
        rndIndex = Int(maxIndxNo * Rnd)
        tmp = myArr(i)
        myArr(i) = myArr(rndIndex)
        myArr(rndIndex) = tmp
    Next
   
    Set Randamize = Me

End Function

'数値を文字列に変換し、その対象がリストに含まれるか検査します。
Public Function StringContains(str As String) As Boolean

    For Each elm In myArr
        If str = CStr(elm) Then
            StringContains = True
            Exit Function
        End If
    Next

    StringContains = False

End Function

'StringBuilderのようにコレクションを高速に結合してStringを返します。
'引数を指定した場合、コレクションの区切りにその文字を付与します。
'長いですがStringはImmutableな参照型であり代入時は値型と同じようにコピーが作られるため、メソッドを分けるのはオーバーヘッドがクソでかいのです
Public Function ToBuildString(Optional delimiter As String = "") As String

    Dim strCount&: strCount = 1
    Dim result$
    Dim chrNumber&
    Dim i&
    Dim elmLength&

    If delimiter = "" Then '引数がない場合
   
        For Each elm In myArr
            strCount = strCount + Len(elm)
        Next
   
        If strCount = 1 Then
            ToBuildString = Empty
            Exit Function
        End If
   
        result = Space(strCount - 1)
   
        chrNumber = 1
        i = 1
        For Each elm In myArr
            elmLength = Len(elm)
            If elmLength > 0 Then
                Mid(result, chrNumber) = elm
                chrNumber = chrNumber + elmLength
            End If
            If i = myCount Then Exit For
            i = i + 1
        Next
   
        ToBuildString = result
       
    Else '引数がある場合
   
        For Each elm In myArr
            strCount = strCount + Len(elm)
        Next
   
        If strCount = 1 Then
            ToBuildString = Empty
            Exit Function
        End If
   
        Dim delimiterSize&: delimiterSize = Len(delimiter)
        result = Space(strCount + delimiterSize * (myCount - 1) - 1)
   
        chrNumber = 1
        i = 1
        For Each elm In myArr
            elmLength = Len(elm)
            If elmLength > 0 Then Mid(result, chrNumber) = elm
            If i = myCount Then Exit For
            Mid(result, chrNumber + elmLength) = delimiter
            chrNumber = chrNumber + delimiterSize + elmLength
            i = i + 1
        Next
   
        ToBuildString = result
       
    End If


End Function

'CSV(Character Separated Value )を作成します。StringBuilderのように高速に結合します。
'セルの値を一括取得して外部にデータ出力する場合極めて強力に働くと思います。

'セル範囲かんたんに書き込みメソッド例(以下のメソッドだとValue化必要です)
'Sub CSVOutputFromRangeValues(arr, destPath as string)
'    Dim list: Set list = New list
'    Open destPath For Output As #1
'        Print #1, list.AddRange(arr).ToBuildCSV(UBound(arr, 2))
'    Close #1
'End Sub

Public Function ToBuildCSV(columnCount As Long, Optional delimiter As String = ",", Optional newlineCode As String = vbCrLf) As String

    Dim strCount&: strCount = 1
    For Each elm In myArr
        strCount = strCount + Len(elm)
    Next

    If strCount = 1 Then
        ToBuildCSV = Empty
        Exit Function
    End If

    Dim delimiterSize&: delimiterSize = Len(delimiter)
    Dim newlineCodeSize&: newlineCodeSize& = Len(newlineCode)
    Dim newLineCount&: newLineCount& = WorksheetFunction.RoundDown((myCount - 1) / columnCount, 0)
    Dim result$: result = Space(strCount + delimiterSize * (myCount - 1) + newlineCodeSize * newLineCount - newLineCount * delimiterSize - 1)

    Dim chrNumber&: chrNumber = 1
    Dim i&: i = 1
    Dim additionalSize&
    Dim elmLength&
   
    For Each elm In myArr
       
        elmLength = Len(elm)
        If elmLength > 0 Then Mid(result, chrNumber) = elm
        If i = myCount Then Exit For
       
        If i Mod columnCount = 0 Then
            Mid(result, chrNumber + elmLength) = newlineCode
            chrNumber = chrNumber + elmLength + newlineCodeSize
        Else
            Mid(result, chrNumber + elmLength) = delimiter
            chrNumber = chrNumber + elmLength + delimiterSize
        End If
       
        i = i + 1
    Next

    ToBuildCSV = result

End Function

'一意の値(ユニーク値)にフィルターして新しいリストとして返します。
Public Function DistinctToList() As list

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim newlist As list: Set newlist = New list

    For Each elm In Me.ToArray
        dic(elm) = 0
    Next
   
    Call newlist.AddRange(dic.Keys)
    Set DistinctToList = newlist

End Function

'数値を文字列に変換し、対象文字列がリストに含まれるか正規表現で検査します。
Public Function StringContains_RegExp(Pattern As String) As Boolean

    Dim regExp As Object
    Set regExp = CreateObject("VBScript.RegExp")
          regExp.Pattern = Pattern

    For Each element In myArr
        If regExp.Test(CStr(element)) Then
            StringContains_RegExp = True
            Exit Function
        End If
    Next

    ContainsRegExp = False

End Function

'射影処理を行います。Excel関数が使用できます。早くはないです。
'使い方はメソッドテストを見てください。
Public Function MAP(variableName As String, evaluateFunction As String) As list

    Dim newlist As list: Set newlist = New list
    Dim func$
   
    For Each elm In Me.ToArray
        func = Replace(evaluateFunction, variableName, elm)
        newlist.Add (Evaluate(func))
    Next
   
    Set MAP = newlist
   
End Function

'フィルター処理を行います。Excel関数が使用できます。早くはないです。
'使い方はメソッドテストを見てください
Public Function Filter(variableName As String, evaluateFunction As String) As list

    Dim newlist As list: Set newlist = New list
    Dim func$
   
    For Each elm In Me.ToArray
        func = Replace(evaluateFunction, variableName, elm)
        If Evaluate(func) Then newlist.Add (elm)
    Next
   
    Set Filter = newlist
   
End Function

'ここから下は数式処理です。
'組み込み関数で行ける範囲の場合はそれを使い高速に処理します。
'2次元配列に再配置すれば、組み込み関数が適用できる範囲が増えますが、遅くなる場合はそれを行いません。

'合計値を取得します。
Public Function Math_Sum()

    If myCount < 65535 Then
        Math_Sum = WorksheetFunction.Sum(myArr)
    Else
        Dim buf
        For Each elm In myArr 'emptyの加算はエラーにならないしはやい
            buf = buf + elm
        Next
        Math_Sum = buf
    End If
   
End Function

'平均値を取得します。
Public Function Math_Average()

    If myCount < 65535 Then
        Math_Average = WorksheetFunction.Average(myArr)
    Else
        Dim buf
        Dim counter As Long: counter = 1
        Dim i&
        For i = LBound(myArr) To myCount - 1
            buf = buf + myArr(i)
        Next
        Math_Average = buf / myCount
    End If
   
End Function


'中央値を取得します。
Public Function Math_Median()

    If myCount < 65535 Then
        Math_Median = WorksheetFunction.Median(myArr)
    ElseIf myCount < 1048575 Then
        Math_Median = WorksheetFunction.Median(Me.ToTransposeArray)
    Else

        Dim newlist As list
        Set newlist = Me.Clone.Sort
        If (myCount Mod 2 = 0) Then
            Math_Median = newlist.arr((newlist.Count) / 2 - 1)
        Else
            Math_Median = (newlist.arr(WorksheetFunction.RoundUp((newlist.Count) / 2, 0) - 1) _
                                + newlist.arr(WorksheetFunction.RoundDown((newlist.Count) / 2, 0) - 1)) _
                                / 2
        End If
    End If
   
End Function

'最大値を返します。
Public Function Math_Max()

    If myCount < 65535 Then
        Math_Max = WorksheetFunction.max(myArr)
    ElseIf myCount < 1048575 Then '意外と拮抗しているので外しても良い
        Math_Max = WorksheetFunction.max(Me.ToTransposeArray)
    Else

    Dim max
    For Each elm In Me.ToArray
           max = IIf(elm > max, elm, max)
    Next
    Math_Max = max
    End If

End Function

'最小値を返します。
Public Function Math_Min()

    If myCount < 65535 Then
        Math_Min = WorksheetFunction.Min(myArr)
    ElseIf myCount < 1048575 Then
        Math_Min = WorksheetFunction.Min(Me.ToTransposeArray)
    Else

    Dim Min: Min = myArr(0)
    For Each elm In Me.ToArray
        Min = IIf(elm < Min, elm, Min)
    Next elm
    Math_Min = Min
    End If

End Function

'母集団の標準偏差を求めます
Public Function Math_StDevP()

    If myCount < 65535 Then
        Math_StDevP = WorksheetFunction.StDev(myArr)
    ElseIf myCount < 1048575 Then
        Math_StDevP = WorksheetFunction.StDev(Me.ToTransposeArray)
    Else
        Dim buf, arg#, v#
        avg = Me.Math_Average
       
        For Each buf In Me.ToArray
            v = v + (avg - buf) ^ 2
        Next
        Math_StDevP = Sqr(v / myCount)
    End If

End Function


'最頻値を1つとってきます。
Public Function Math_ModeSingle()

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim cnt&: cnt = 1
    '数値,個数の辞書を作る
    For Each elm In Me.arr
        If Not (IsEmpty(elm)) Then
            dic(elm) = dic(elm) + 1
        End If
        cnt = cnt + 1
        If cnt >= 2500 Then '辞書に入れる数が多くなるとものすごく遅くなる 多分O(N^2)
            Math_ModeSingle = ModeSinge_LowOverlap
            Exit Function
        End If
    Next

    If dic.Count = 0 Then Exit Function

    '最大の個数を調べる
    Dim maxCounter As Long
    Dim val As Long
    For Each elm In dic.Items
        maxCounter = IIf(maxCounter < elm, elm, maxCounter)
    Next elm

    '最大値と一致する値を返す
    For Each elm In dic.Keys
        If dic(elm) = maxCounter Then
            Math_ModeSingle = elm
            Exit Function
        End If
    Next

End Function


'あまりにも重複が少ない場合
Private Function ModeSinge_LowOverlap()
   
    Dim result
    Dim maxCount&, currentCount&
    Dim beforeValue
   
    For Each elm In Me.Clone.Sort.ToArray
   
        If elm = beforeValue Then
            currentCount = currentCount + 1
        Else
            If currentCount > maxCount Then
                maxCount = currentCount
                result = beforeValue
            End If
            currentCount = 1
            beforeValue = elm
        End If
    Next
   
    If currentCount > maxCount Then result = beforeValue
   
    ModeSinge_LowOverlap = result
   
End Function


'最頻値を配列で返します。
Public Function Math_ModeMulti()

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim cnt&: cnt = 1
    '数値,個数の辞書を作る
    For Each elm In Me.arr
        If Not (IsEmpty(elm)) Then
            dic(elm) = dic(elm) + 1
        End If
        cnt = cnt + 1
        If cnt >= 2500 Then '辞書に入れる数が多くなるとものすごく遅くなる 多分O(N^2)
            Math_ModeMulti = ModeMulti_LowOverlap
            Exit Function
        End If
    Next

    If dic.Count = 0 Then Exit Function

    '最大の個数を調べる
    Dim maxCounter As Long
    Dim val As Long
    For Each elm In dic.Items
        maxCounter = IIf(maxCounter < elm, elm, maxCounter)
    Next elm

    '最大値と一致する値を入れる
    Dim newlist As list: Set newlist = New list
    For Each elm In dic.Keys
        If dic(elm) = maxCounter Then
            Call newlist.Add(elm)
        End If
    Next

    Math_ModeMulti = newlist.ToArray

End Function

'あまりにも重複が少ない場合
Private Function ModeMulti_LowOverlap()
   
    Dim resultBuffer As list: Set resultBuffer = New list
    Dim maxCount&, currentCount&
    Dim beforeValue
 
    For Each elm In Me.Clone.Sort.ToArray
   
        If elm = beforeValue Then
            currentCount = currentCount + 1
        Else
            If currentCount > maxCount Then
                resultBuffer.Clear.Add (beforeValue)
                maxCount = currentCount
            ElseIf currentCount = maxCount Then
                resultBuffer.Add (beforeValue)
            End If
            currentCount = 1
            beforeValue = elm
        End If
    Next

    If currentCount > maxCount Then
        resultBuffer.Clear.Add (beforeValue)
    ElseIf currentCount = maxCount Then
        resultBuffer.Add (beforeValue)
    End If

    ModeMulti_LowOverlap = resultBuffer.ToArray
   
End Function



