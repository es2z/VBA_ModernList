VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'極めて高速、多機能に動く組み込み型で使用できる単方向連接リスト(のように振る舞うもの)です。
'Sortやユニーク化、Linq様のメソッドチェーンや様々な配列に再配置したり、抜き出し、集合に対応しています。
'こちらのクラスは自作クラス向けで、利用にはList.clsのインポートが必要です(List型に変換するメソッドが有るため)
'こちらは、List.clsに比べてデバッグがおざなりになっている状態です。問題があれば報告等お願いします。

'こちらのリポジトリは以下になります。スターとかおなしゃす!!
'https://github.com/es2z/VBA_ModernList

Option Base 0
Private myArr() As Variant 'ここに値が入ります
Private myCount As Long '配列の要素数です。
Private myCapacity As Long '配列の最大サイズです。

Property Get arr() As Variant: arr = myArr: End Property
Property Get Count() As Long: Count = myCount: End Property
Property Get Capacity() As Long: Capacity = myCapacity: End Property

'コンストラクタです。クラスの初期化を行います。 配列の初期サイズ決めたかったのですが、仕様上引数渡せないみたいです()
Private Sub Class_Initialize()
    myCapacity = 0
    ReDim myArr(myCapacity)
End Sub

'デストラクタです、メモリを開放します。Variant()にnothingを代入することはできません
Private Sub Class_Terminate()
    Call Clear
    VBA.DoEvents
End Sub

'クラスを初期化します。
Public Function Clear() As ObjectList
    myCount = 0
    Call Class_Initialize
    Set Clear = Me 'メソッドチェーンの為
End Function

'値をリストに追加します。
'動作速度とメモリの観点から、格納アイテム数が配列のキャパシティを超えた場合のみReDim Preserveを行いキャパシティを倍にします。
Public Sub Add(obj As Variant)
   
    If myCount > myCapacity Then
        myCapacity = (myCapacity + 1) * 2 - 1
        ReDim Preserve myArr(myCapacity)
    End If
   
    Set myArr(myCount) = obj
    myCount = myCount + 1
   
End Sub

'配列を対象とした高速な範囲追加を行います。 予め配列範囲を確保してそこに値を入れていきます。
Public Function AddRange(args As Variant) As list
   
   '次元数判定
   Dim dimensionsCount&: dimensionsCount = 1
   Dim buffer&
On Error GoTo err
    While True
        buffer = UBound(args, dimensionsCount)
        dimensionsCount = dimensionsCount + 1
    Wend
err:
    dimensionsCount = dimensionsCount - 1

    '範囲の確保
    Dim insertPoint As Long: insertPoint = myCount
    
    If dimensionsCount = 1 Then
        myCount = myCount + UBound(args) - LBound(args) + 1
    Else
        Dim additionalElmCnt&: additionalElmCnt = 1
        For i = 1 To dimensionsCount
            additionalElmCnt = additionalElmCnt * (UBound(args, dimensionsCount) - LBound(args, dimensionsCount) + 1)
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
        Set myArr(insertPoint) = arg
        insertPoint = insertPoint + 1
    Next
    
    Set AddRange = Me

End Function

'2つのリストを結合します。引数に入れたほうがあとに来ます
Public Function Concat(argList As ObjectList) As ObjectList

    If argList.Count = 0 Then Exit Function
    Call AddRangeForOneDimension(argList.ToArray)
   
    Set Concat = Me

End Function

'最初の要素を取得します。
Public Function First(): Set First = IIf(myCount = 0, Empty, myArr(0)): End Function
'最後の要素を取得します。
Public Function Last(): Set Last = IIf(myCount = 0, Empty, myArr(myCount - 1)): End Function

'何か要素が格納されていればTrueを返します。
Public Function Any_(): Any_ = myCount >= 1: End Function

'何も要素が格納されていなければTrueを返します。
Public Function Nothing_(): Nothing_ = myCount = 0: End Function

'配列に変換して返します。 配列の要素数はデータ数に切り詰められます。
Public Function ToArray() As Variant()

    Dim bufArr(): ReDim bufArr(myCount - 1)
    Dim i&
    For i = 0 To (myCount - 1)
        Set bufArr(i) = myArr(i)
    Next
   
    ToArray = bufArr
   
End Function

'chunkSizeの個数ごとに分けた配列を配列の中に入れて返します。 例100個のデータを40で分けた場合Array(Array(40),Array(40),Array(20))
Public Function ToChunkArray(ChunkSize As Long) As Variant()

    Dim oneD&: oneD = WorksheetFunction.RoundUp(myCount / ChunkSize, 0) - 1
    Dim towD&: twoD = ChunkSize - 1
   
    Dim result(): ReDim result(oneD)
    Dim bufArr()
    Dim i&, j&, k&
   
    For i = 0 To oneD
        ReDim bufArr(twoD)
        For j = 0 To twoD
            Set bufArr(j) = myArr(k)
            k = k + 1
        Next
        result(i) = bufArr
    Next
   
    ToChunkArray = result

End Function


'オブジェクトを複製します。実装はShallowCopyです。
'DeepCopyを実装するには、オブジェクト側にDeepCopyを実装して、こいつのSelectメソッドを使うなどでできるかなと思います。
Public Function Clone() As ObjectList

    Dim newlist As ObjectList: Set newlist = New ObjectList
    Set Clone = newlist.Concat(Me)

End Function

'要素名を基準に昇順でソートします。(実装は多分クイックソートになっているはずです)
'キャッシュを作って高速化しました。
Public Function SortByElementNameIn_Ascending(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
   
    '引数のメンバー名が見つからなかった場合
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception

    Dim minIndex As Long: minIndex = LBound(myArr)
    Dim maxIndex As Long: maxIndex = UBound(myArr)
    Dim cache: ReDim cache(minIndex To maxIndex)
    Dim i&: i = minIndex
   
    Dim arg
    'ソート対象となる属性をキャッシュする。この内容を比較してクイックソートする形、キャッシュしないと~1桁重くなるかと
    For Each arg In myArr
        cache(i) = CallByName(arg, memberName, VbGet)
        i = i + 1
    Next arg
   
    Call SortBy(cache, minIndex, maxIndex)
    Set SortByElementNameIn_Ascending = Me

Exit Function

Exception:
Call MsgBox("リストに対してソートを実行しようとしましたが、引数のメンバー名が見つかりませんでした", vbCritical)

End Function
Private Sub SortBy(ByRef cache, ByVal minIndex As Long, ByVal maxIndex As Long)

    Dim low As Long, high As Long
    Dim pivot As Variant
   
    pivot = cache(Int((minIndex + maxIndex) / 2))
   
    low = minIndex
    high = maxIndex
    Do
        Do While cache(low) < pivot
            low = low + 1
        Loop

        Do While cache(high) > pivot
            high = high - 1
        Loop
       
        If low >= high Then Exit Do
       
    Dim vSwap As Variant, vSwapCache As Variant
        Set vSwap = myArr(low)
        vSwapCache = cache(low)
        Set myArr(low) = myArr(high)
        cache(low) = cache(high)
        Set myArr(high) = vSwap
        cache(high) = vSwapCache
       
        low = low + 1
        high = high - 1
    Loop

    If (minIndex < low - 1) Then
        Call SortBy(cache, minIndex, low - 1)
    End If
    If (maxIndex > high + 1) Then
        Call SortBy(cache, high + 1, maxIndex)
    End If
   
End Sub


'要素名を基準に昇順でソートします。(実装は多分クイックソートになっているはずです
'キャッシュを作って高速化しています。
Public Function SortByElementNameIn_Descending(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
   
    '引数のメンバー名が見つからなかった場合
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception

    Dim minIndex As Long: minIndex = LBound(myArr)
    Dim maxIndex As Long: maxIndex = UBound(myArr)
    Dim cache: ReDim cache(minIndex To maxIndex)
    Dim i: i = minIndex
   
    Dim arg
    'ソート対象となる属性をキャッシュする。この内容を比較してクイックソートする形
    For Each arg In myArr
        cache(i) = CallByName(arg, memberName, VbGet)
        i = i + 1
    Next arg
   
    Call SortByDescending(cache, minIndex, maxIndex)
    Set SortByElementNameIn_Descending = Me

Exit Function

Exception:
Call MsgBox("リストに対してソートを実行しようとしましたが、引数のメンバー名が見つかりませんでした", vbCritical)

End Function
Private Sub SortByDescending(ByRef cache, ByVal minIndex As Long, ByVal maxIndex As Long)

    Dim low As Long, high As Long
    Dim pivot As Variant
   
    pivot = cache(Int((minIndex + maxIndex) / 2))
   
    low = minIndex
    high = maxIndex
   
    Do
        Do While cache(low) > pivot
            low = low + 1
        Loop

        Do While cache(high) < pivot
            high = high - 1
        Loop
       
        If low >= high Then Exit Do
       
    Dim vSwap As Variant, vSwapCache As Variant
        Set vSwap = myArr(low)
        vSwapCache = cache(low)
        Set myArr(low) = myArr(high)
        cache(low) = cache(high)
        Set myArr(high) = vSwap
        cache(high) = vSwapCache
       
        low = low + 1
        high = high - 1
    Loop

    If (minIndex < low - 1) Then
        Call SortByDescending(cache, minIndex, low - 1)
    End If
    If (maxIndex > high + 1) Then
        Call SortByDescending(cache, high + 1, maxIndex)
    End If
   
End Sub

 'インデックス値に対応する値を取得します。配列の[n]と同じです。
 Public Function GetItmeOfIndex(index As Long)
 
    If index > myCount Then
         Set GetItmeOfIndex = Null
    Else
        Set GetItmeOfIndex = myArr(index)
    End If
   
End Function

'特定の範囲を抜き出して新しいリストとして返します。(minIndex<=抜き出すもの<=maxIndex)
'Minが不当に小さい、Maxが不当に大きい場合、最小インデックス、最大インデックスまでの要素を対象とします。
Public Function Slice(minIndex As Long, maxIndex As Long) As ObjectList

    If minIndex > maxIndex Then
        Set Slice = New ObjectList
        Exit Function
    End If
   
    minIndex = IIf(minIndex < LBound(myArr), LBound(myArr), minIndex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)
   
    Dim newlist As ObjectList: Set newlist = New ObjectList
   
    For i = minIndex To maxIndex
        Call newlist.Add(myArr(i))
    Next i
   
    Set Slice = newlist

End Function

'指定数先頭から抜き出す
Public Function TakeToList(itemCount As Long) As ObjectList

    '指定数が要素数を超えている場合クローンを返します
    Dim minIndex As Long: minIndex = LBound(myArr)
    If itemCount + 1 - minIndex > myCount Then
        Set TakeToList = Me.DeepCopy()
    Else
        Dim newlist As ObjectList: Set newlist = New ObjectList
        For i = minIndex To itemCount - 1 + minIndex
            Call newlist.Add(myArr(i))
        Next i
        Set TakeToList = newlist
    End If

End Function

'指定数をスキップしてそれ以降を抜き出す
Public Function SkipToList(skipCount As Long) As ObjectList

    '指定数が要素数を超えている場合空のObjectListを返します。
    If skipCount > myCount Then
        Set SkipToList = New ObjectList
    Else
        Dim newlist As ObjectList: Set newlist = New ObjectList
        For i = skipCount To myCount - 1
            Call newlist.Add(myArr(i))
        Next i
        Set SkipToList = newlist
    End If
   
End Function

'格納されている値を反転します。
Public Function Reverse() As ObjectList

    If myCount < 1 Then Exit Function
    Dim maxIndxNo As Long: maxIndxNo = myCount - 1
    Dim buf(): ReDim buf(0 To maxIndxNo)
   
    Dim counter As Long: counter = maxIndxNo
   
    For i = 0 To maxIndxNo
        Set buf(counter) = myArr(i)
        counter = counter - 1
    Next i
   
    myArr = buf
   
    Set Reverse = Me
   
End Function

' 値をシャッフルします。精度はいまいちかもしれません。精度が欲しければメルセンヌツイスターとか良いらしいので実装して下さい()
Public Function Randamize() As ObjectList

    Dim maxIndxNo As Long: maxIndxNo = myCount - 1
    Dim i As Long, rndIndex As Variant, tmp As Variant
    For i = 0 To maxIndxNo
        rndIndex = Int(maxIndxNo * Rnd)
        Set tmp = myArr(i)
        Set myArr(i) = myArr(rndIndex)
        Set myArr(rndIndex) = tmp
    Next
   
    Set Randamize = Me
   
End Function

'一意の値(ユニーク値)にフィルターして新しいリストとして返します。
Public Function DistinctToListByElementName(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
       
    '引数のメンバー名が見つからなかった場合
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception
   
    Dim minIndex As Long: minIndex = LBound(myArr)
    Dim maxIndex As Long: maxIndex = UBound(myArr)
    Dim cache: ReDim cache(minIndex To maxIndex)
    Dim i As Long: i = minIndex
    Dim arg
   
    'ソート対象となる属性をキャッシュする。この内容を比較してディクショナリに入れる
    For Each arg In myArr
        cache(i) = CallByName(arg, memberName, VbGet)
        i = i + 1
    Next arg
   
    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim newlist As ObjectList: Set newlist = New ObjectList
   
    For j = minIndex To maxIndex
        Set dic(cache(j)) = myArr(j)
    Next j
   
    For Each Item In dic.Items
        Call newlist.Add(Item)
    Next
   
    Set DistinctToListByElementName = newlist

Exit Function

Exception:
    Call MsgBox("リストに対してユニーク化を実行しようとしましたが、引数のメンバー名が見つかりませんでした", vbCritical)
End Function

'リスト内オブジェクトの任意要素に対して正規表現でフィルターを行い、ヒットしたものを新しいリストに格納して返す。
Public Function FilterRegExpInElementName(ElementName As String, RegexpPattern As String, Optional IgnoreCase As Boolean = False) As ObjectList

    '引数のメンバー名が見つからなかった場合
    On Error Resume Next
        buf = CallByName(myArr(0), ElementName, VbGet)
    On Error GoTo Exception

    Dim newlist As ObjectList: Set newlist = New ObjectList
    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = RegexpPattern
    regExp.IgnoreCase = IgnoreCase
   
    For i = 0 To myCount - 1
        If regExp.Test(CallByName(myArr(i), ElementName, VbGet)) Then
            Call newlist.Add(myArr(i))
        End If
    Next i

    Set FilterRegExpInElementName = newlist

Exit Function
Exception:
    Call MsgBox("リストに対してフィルター処理(Filter)を実行しようとしましたが、引数のメンバー名が見つかりませんでした", vbCritical)
   
End Function

'リスト内オブジェクトの任意要素を新しいlist(値型)として返す。
'ユーザー定義型以外だといまいち上手く動かないみたいです
Public Function SliceElementByName(ElementName As String) As list

    '引数のメンバー名が見つからなかった場合
    On Error Resume Next
        buf = CallByName(myArr(0), ElementName, VbGet)
    On Error GoTo Exception

    Dim newlist As list: Set newlist = New list
   
    For i = 0 To myCount - 1
        Call newlist.Add(CallByName(myArr(i), ElementName, VbGet))
    Next i

    Set SelectByElementName = newlist
   
Exit Function
Exception:
    Call MsgBox("リストに対して切り出し処理(Slice)を実行しようとしましたが、引数のメンバー名が見つかりませんでした", vbCritical)
   
End Function
