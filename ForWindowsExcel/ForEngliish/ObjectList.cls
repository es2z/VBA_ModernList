VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'The repository here is as follows.
'https://github.com/es2z/VBA_ModernList

'Comments are machine translated by DeepL translation.
'It may be more accurate to read the code for the actual behavior.

'It is almost identical to a list object, except that it can contain objects.

Option Base 0
Private myArr() As Variant
Private myCount As Long
Private myCapacity As Long

Property Get arr() As Variant: arr = myArr: End Property
Property Get Count() As Long: Count = myCount: End Property
Property Get Capacity() As Long: Capacity = myCapacity: End Property

Private Sub Class_Initialize()
    myCapacity = 0
    ReDim myArr(myCapacity)
End Sub

'Free up memory.
Private Sub Class_Terminate()
    Call Clear
    VBA.DoEvents
End Sub

'Delete all stored contents.
Public Function Clear() As ObjectList
    myCount = 0
    Call Class_Initialize
    Set Clear = Me 'For method chain
End Function

' Append the value to the list.
' From the point of view of speed and memory, ReDim Preserve is performed only when the number of stored items exceeds the capacity of the array, doubling the capacity.
Public Sub Add(obj As Variant)
   
    If myCount > myCapacity Then
        myCapacity = (myCapacity + 1) * 2 - 1
        ReDim Preserve myArr(myCapacity)
    End If
   
    Set myArr(myCount) = obj
    myCount = myCount + 1
   
End Sub

'Performs fast append targeting arrays.
'Allocates memory space first and then performs the addition.
Public Function AddRange(args As Variant) As List
   
   'Determination of number of dimensions
   Dim dimensionsCount&: dimensionsCount = 1
   Dim buffer&
On Error GoTo err
    While True
        buffer = UBound(args, dimensionsCount)
        dimensionsCount = dimensionsCount + 1
    Wend
err:
    dimensionsCount = dimensionsCount - 1

    'Allocation of Array space
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
   
   'Append
    For Each arg In args
        Set myArr(insertPoint) = arg
        insertPoint = insertPoint + 1
    Next
    
    Set AddRange = Me

End Function

' Combines two listings. The one you put in the argument comes after.
Public Function Concat(argList As ObjectList) As ObjectList

    If argList.Count = 0 Then Exit Function
    Call AddRange(argList.ToArray)
   
    Set Concat = Me

End Function

' Get the first element.
Public Function First(): Set First = IIf(myCount = 0, Empty, myArr(0)): End Function

'Get the last element
Public Function Last(): Set Last = IIf(myCount = 0, Empty, myArr(myCount - 1)): End Function

' Returns True if any element is stored.
Public Function Any_(): Any_ = myCount >= 1: End Function

' Returns True if no elements are stored.
Public Function Nothing_(): Nothing_ = myCount = 0: End Function

' Returns True if the contents of the two lists all match.
Public Function ToArray() As Variant()

    Dim bufArr(): ReDim bufArr(myCount - 1)
    Dim i&
    For i = 0 To (myCount - 1)
        Set bufArr(i) = myArr(i)
    Next
   
    ToArray = bufArr
   
End Function

'Returns an array divided by the number of 'chunkSize' in the array.
'Example 100 data divided by 40 Array(Array(39),Array(39),Array(39(the last 20 are empty)))
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

'Duplicate (shallow copy) an List.
Public Function Clone() As ObjectList

    Dim newlist As ObjectList: Set newlist = New ObjectList
    Set Clone = newlist.Concat(Me)

End Function

'Sort in ascending order by element name. (Implementation is quick sort.)
'A cache is created to speed up the process.
Public Function SortByElementNameIn_Ascending(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
   
    'If the argument member name is not found
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception

    Dim minIndex As Long: minIndex = LBound(myArr)
    Dim maxIndex As Long: maxIndex = UBound(myArr)
    Dim cache: ReDim cache(minIndex To maxIndex)
    Dim i&: i = minIndex
   
    Dim arg
    'Cache the attributes to be sorted. Compare these contents for quick sorting, which may be ~1 order of magnitude heavier if not cached.
    For Each arg In myArr
        cache(i) = CallByName(arg, memberName, VbGet)
        i = i + 1
    Next arg
   
    Call SortBy(cache, minIndex, maxIndex)
    Set SortByElementNameIn_Ascending = Me

Exit Function

Exception:
Call MsgBox("Attempted to perform a sort on a list, but could not find the member name for the argument", vbCritical)

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

'Sort in descending order based on element name.
'The content is almost the same as the ascending order
Public Function SortByElementNameIn_Descending(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
   
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception

    Dim minIndex As Long: minIndex = LBound(myArr)
    Dim maxIndex As Long: maxIndex = UBound(myArr)
    Dim cache: ReDim cache(minIndex To maxIndex)
    Dim i: i = minIndex
   
    Dim arg
    For Each arg In myArr
        cache(i) = CallByName(arg, memberName, VbGet)
        i = i + 1
    Next arg
   
    Call SortByDescending(cache, minIndex, maxIndex)
    Set SortByElementNameIn_Descending = Me

Exit Function

Exception:
Call MsgBox("Attempted to perform a sort on a list, but could not find the member name for the argument", vbCritical)

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

 ' Get the value corresponding to the index value. Same as [index] in the array.
 Public Function GetItmeOfIndex(index As Long)
 
    If index > myCount Then
         Set GetItmeOfIndex = Null
    Else
        Set GetItmeOfIndex = myArr(index)
    End If
   
End Function

'Extracts a specific range and returns it as a new list. (minIndex<=what is extracted<=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored is targeted for retrieval.
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

'Extracts the specified number of values from the beginning.
Public Function TakeToList(itemCount As Long) As ObjectList

    'If the specified number exceeds the number of elements, an empty list is returned.
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

'skip the specified number of values and extract the rest.
Public Function SkipToList(skipCount As Long) As ObjectList

    'If the specified number exceeds the number of elements, an empty list is returned.
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

' Reverses the order of stored values.
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

'Shuffles stored values.Accuracy may not be good.
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

'Unique by 'element name
Public Function DistinctToListByElementName(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
       
    'If the argument member name is not found
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception
   
    Dim minIndex As Long: minIndex = LBound(myArr)
    Dim maxIndex As Long: maxIndex = UBound(myArr)
    Dim cache: ReDim cache(minIndex To maxIndex)
    Dim i As Long: i = minIndex
    Dim arg
   
    ' Cache the attributes to be sorted. Compare this content and put it in the dictionary.
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
    Call MsgBox("Tried to perform unification on a list, but could not find the member name for the argument", vbCritical)
End Function

''Filter by regular expression for any element of an object in the list, and store the hits in a new list and return them
Public Function FilterRegExpInElementName(ElementName As String, RegexpPattern As String, Optional IgnoreCase As Boolean = False) As ObjectList

    'If the argument member name is not found
    On Error Resume Next
        buf = CallByName(myArr(0), ElementName, VbGet)
    On Error GoTo Exception

    Dim newlist As ObjectList: Set newlist = New ObjectList
    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")  'New regExp
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
    Call MsgBox("Tried to filter a list, but could not find the member name for the argument.", vbCritical)
   
End Function

'Returns any element of an object in a list as a new list (value type).
' It doesn't seem to work well for non-user-defined types.
Public Function SliceElementByName(ElementName As String) As List

    'If the argument member name is not found
    On Error Resume Next
        buf = CallByName(myArr(0), ElementName, VbGet)
    On Error GoTo Exception

    Dim newlist As List: Set newlist = New List
   
    For i = 0 To myCount - 1
        Call newlist.Add(CallByName(myArr(i), ElementName, VbGet))
    Next i

    Set SelectByElementName = newlist
   
Exit Function
Exception:
    Call MsgBox("Attempted to perform MAP processing on a list, but could not find the member name for the argument", vbCritical)
   
End Function
