VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The repository here is as follows.
'https://github.com/es2z/VBA_ModernList

'Comments are machine translated by DeepL translation.
'It may be more accurate to read the code for the actual behavior.

Option Base 0
Private myArr() As Variant
Private myCount As Long
Private myCapacity As Long

Property Get arr(): arr = myArr: End Property
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
Public Function Clear() As List
    myCount = 0
    Call Class_Initialize
    Set Clear = Me 'For method chain
End Function

'Debugging output. It works no matter where you add it in the method chain.
Public Function DebugPrint(Optional Prefix As String, Optional FormatStr As String, Optional Suffix As String) As List

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

' Append the value to the list.
' From the point of view of speed and memory, ReDim Preserve is performed only when the number of stored items exceeds the capacity of the array, doubling the capacity.
Public Sub Add(val)

    If myCount > myCapacity Then
        myCapacity = (myCapacity + 1) * 2 - 1
        ReDim Preserve myArr(myCapacity)
    End If

    myArr(myCount) = val
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
        myArr(insertPoint) = arg
        insertPoint = insertPoint + 1
    Next
    
    Set AddRange = Me

End Function

' Combines two listings. The one you put in the argument comes after.
Public Function Concat(argList As List) As List

    If argList.Count = 0 Then
        Set Concat = Me
        Exit Function
    End If
   
    Call AddRange(argList.ToArray)
    Set Concat = Me

End Function

' Get the first element.
Public Function First(): First = IIf(myCount = 0, Empty, myArr(0)): End Function

'Get the last element
Public Function Last(): Last = IIf(myCount = 0, Empty, myArr(myCount - 1)): End Function

' Returns True if any element is stored.
Public Function Any_(): Any_ = myCount >= 1: End Function

' Returns True if no elements are stored.
Public Function Nothing_(): Nothing_ = myCount = 0: End Function

' Returns True if the contents of the two lists all match.
Public Function SequenceEqual(compairTarget As List) As Boolean
   
    If compairTarget.Count <> Me.Count Then Exit Function
   
    Dim targetArr(): targetArr = compairTarget.ToArray
    Dim i&
    For i = 0 To Me.Count - 1
        If Me.arr(i) <> targetArr(i) Then Exit Function
    Next
   
    SequenceEqual = True
   
End Function

' Clears the contents and creates a sequential number. It is similar in form to a for statement.
Public Function CreateSeqNumbers(First As Long, Last As Long, Optional step As Long = 1) As List

    Me.Clear
    Dim i&
    For i = First To Last Step step
        Me.Add (i)
    Next i
   
    Set CreateSeqNumbers = Me
   
End Function

'Clears the contents and creates a sequential number for the specified number of pieces from the starting position
Public Function CreateEnumRange(start As Long, Count As Long) As List

    Me.Clear
    Dim i&
    For i = start To start + Count - 1
        Me.Add (i)
    Next i
   
    Set CreateEnumRange = Me
   
End Function

 ' Get the value corresponding to the index value. Same as [index] in the array.
 Public Function GetValueOfIndex(index As Long)

    If index > myCount Then
        GetValueOfIndex = Empty
    Else
        GetValueOfIndex = myArr(index)
    End If

End Function

'Extracts a specific range and returns it as a new list. (minIndex<=what is extracted<=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored is targeted for retrieval.
Public Function Slice(minIndex As Long, maxIndex As Long) As List

    minIndex = IIf(minIndex < LBound(myArr), LBound(myArr), minIndex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minIndex > maxIndex Then
        Set Slice = New List
        Exit Function
    End If

    Dim newlist As List: Set newlist = New List
    Dim i&
    For i = minIndex To maxIndex
        newlist.Add (myArr(i))
    Next i

    Set Slice = newlist

End Function

'Remove the element corresponding to the index and pack the data in front of it. It is not efficient.
Public Function Remove(index As Long) As List

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

' Removes all matches to the numerical value of the argument and prepends the data.
' (The implementation creates a new list from scratch.)
Public Function RemoveAll(val) As List

    Dim newlist As List: Set newlist = New List

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

'Delete a specific range.(minIndex<=what is deleted <=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored  is subject to deletion.
Public Function RemoveRange(minIndex As Long, maxIndex As Long) As List

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

' Returns the element corresponding to the index and then removes that element from the list.
' The return value is the value obtained.
Public Function Pop(index As Long) As Variant

    If index < LBound(myArr) Or UBound(myArr) < index Then
        Pop = Empty: Exit Function
    End If

    Pop = myArr(index)
    Call Remove(index)

End Function

'Retrieve a specific range of values and simultaneously delete the targeted values. (minIndex<=deleted<=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored is targeted for retrieval.
' The return value is the value obtained.
Public Function PopRange(minIndex As Long, maxIndex As Long) As List

    minIndex = IIf(minIndex < LBound(myArr), LBound(myArr), minIndex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minIndex > maxIndex Then
        Set PopRange = New List
        Exit Function
    End If

    Dim difference: difference = maxIndex - minIndex
    Dim result As List: Set result = New List

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

' Converted to an array and returned. The number of elements in the array is truncated to the stored value.
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

'Converts to a 2-dimensional array and returns it with values in the 1-dimensional elements. The number of array elements is truncated to the number of data.
'(It can be conveniently used to attach data to an Excel sheet.)
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


' Converts to a 2-dimensional array and returns it. You can set the number of the second dimension.
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

' Converts to a 2-dimensional array and returns it. You can set the number of the first dimension.
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

'Returns an array divided by the number of 'chunkSize' in the array.
'Example 100 data divided by 40 Array(Array(39),Array(39),Array(39(the last 20 are empty)))
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

'Transpose the array in Chunk and make it a 2-dimensional array.
'(This was useful when turning and pasting a For(Each).
'Example 100 data divided by 40 Array(Array(0,39),Array(0,39),Array(0,39(the last 20 are empty)))
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

'Contrary to 'chunk', you can decide how many pieces of data to divide into and place them evenly there.
'Example: 100 data divided by 3 Array(Array(33),Array(33 last empty),Array(33 last empty))
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

'Returns a Dictionary with only the 'Keys' set. Duplicates will be removed, which may be useful if you have a large number of exists and so on.
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

'Creates a union set (combined and unique) of lists and returns it in a new list
Public Function UnionToList(argsList As List) As List

    Dim dic As Object: Set dic = Me.ToHashset
    Dim newlist As List: Set newlist = Me.Clone
   
    For Each elm In argsList.ToArray
        If Not dic.Exists(elm) Then newlist.Add (elm)
    Next
   
    Set UnionToList = newlist
   
End Function

'Creates a difference set (what was originally there minus what is in the argument) between lists and returns it in a new list.
Public Function ExceptToList(argsList As List) As List

    Dim dic As Object: Set dic = argsList.ToHashset
    Dim List As List: Set List = New List
    Dim newLis As List: Set newlist = New List
   
    For Each elm In Me.ToArray
        If Not dic.Exists(elm) Then newlist.Add (elm)
    Next
   
    Set ExceptToList = newlist
   
End Function

' Creates a product set (what is in both) between lists and returns it in a new list.
Public Function IntersectToList(argsList As List) As List

    Dim dic As Object: Set dic = argsList.ToHashset
    Dim List As List: Set List = New List
   
    For Each elm In Me.ToArray
        If dic.Exists(elm) Then List.Add (elm)
    Next
   
    Set IntersectToList = List
   
End Function

'Duplicate (shallow copy) an List.
Public Function Clone() As List

    Dim newlist As List: Set newlist = New List
    Set Clone = newlist.Concat(Me)

End Function

'Extracts the specified number of values from the beginning.
Public Function TakeToList(itemCount As Long) As List

    'If the specified number exceeds the number of elements, an empty list is returned.
    Dim minIndex As Long: minIndex = LBound(myArr)
    If itemCount + 1 - minIndex > myCount Then
        Set TakeToList = Me.Clone()
    Else
        Dim newlist As List: Set newlist = New List
        For i = minIndex To itemCount - 1 + minIndex
            newlist.Add (myArr(i))
        Next i
        Set TakeToList = newlist
    End If

End Function

'skip the specified number of values and extract the rest.
Public Function SkipToList(skipCount As Long) As List

    'If the specified number exceeds the number of elements, an empty list is returned.
    If skipCount > myCount Then
        Set SkipToList = New List
    Else
        Dim newlist As List: Set newlist = New List
        For i = skipCount To myCount - 1
            newlist.Add (myArr(i))
        Next i
        Set SkipToList = newlist
    End If

End Function

'Sort in ascending order. (Implementation is quick sort.)
Public Function Sort() As List

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

' Sort in descending order.
Public Function SortByDescending() As List

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

' Reverses the order of stored values.
Public Function Reverse() As List

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

'Shuffles stored values.Accuracy may not be good.
Public Function Randamize() As List

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

' Assumes the stored value is a string and inspects whether the target is included in the list.
Public Function StringContains(str As String) As Boolean

    For Each elm In myArr
        If str = CStr(elm) Then
            StringContains = True
            Exit Function
        End If
    Next

    StringContains = False

End Function

'Fast concatenates a collection like StringBuilder and returns a String.
' If an argument is specified, the character is given to delimit the collection.
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

'Create CSV (Character Separated Value ), fast join like StringBuilder.
'If combined with Excel's Range(area).Value, etc., it should work very conveniently for external data output.
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

'Filter to unique values and return as a new list.
Public Function DistinctToList() As List

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim newlist As List: Set newlist = New List

    For Each elm In Me.ToArray
        dic(elm) = 0
    Next
   
    Call newlist.AddRange(dic.Keys)
    Set DistinctToList = newlist

End Function

' Assuming the stored value to be a string, checks whether the target string is included in the list using a regular expression.
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

' MAP processing; Excel functions can be used. It is not fast.
'See method test for usage.
Public Function MAP(variableName As String, evaluateFunction As String) As List

    Dim newlist As List: Set newlist = New List
    Dim func$
   
    For Each elm In Me.ToArray
        func = Replace(evaluateFunction, variableName, elm)
        newlist.Add (Evaluate(func))
    Next
   
    Set MAP = newlist
   
End Function

' Filter processing; Excel functions can be used. It is not fast.
'See method test for usage.
Public Function Filter(variableName As String, evaluateFunction As String) As List

    Dim newlist As List: Set newlist = New List
    Dim func$
   
    For Each elm In Me.ToArray
        func = Replace(evaluateFunction, variableName, elm)
        If Evaluate(func) Then newlist.Add (elm)
    Next
   
    Set Filter = newlist
   
End Function

'From here on down is formula processing.
'If the range can be covered by the built-in functions, they are used to speed up the process.
'Re-arranging into a 2-dimensional array increases the range where the built-in functions can be applied, but does not do so if it slows down the process.

'!!! Since it is assumed that numerical values are stored, it will not work correctly if the elements stored in the list contain strings. !!!
'!!! I use the desktop version of Excel in a Windows environment, so I have not verified that it works in other environments. !!!

' Return the total value.
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

' Return the average value.
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

' Return  the median value.
Public Function Math_Median()

    If myCount < 65535 Then
        Math_Median = WorksheetFunction.Median(myArr)
    ElseIf myCount < 1048575 Then
        Math_Median = WorksheetFunction.Median(Me.ToTransposeArray)
    Else

        Dim newlist As List
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

' Return the maximum value.
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

' Return the minimum value.
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

'Return the standard deviation of the population
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

' Return one mode.
Public Function Math_ModeSingle()

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim cnt&: cnt = 1
    For Each elm In Me.arr
        If Not (IsEmpty(elm)) Then
            dic(elm) = dic(elm) + 1
        End If
        cnt = cnt + 1
        If cnt >= 2500 Then 'The more you put into the dictionary, the slower things get, maybe O(N^2).
            Math_ModeSingle = ModeSinge_LowOverlap
            Exit Function
        End If
    Next

    If dic.Count = 0 Then Exit Function

    'Find the largest number of pieces
    Dim maxCounter As Long
    Dim val As Long
    For Each elm In dic.Items
        maxCounter = IIf(maxCounter < elm, elm, maxCounter)
    Next elm

    ' Returns a value that matches the maximum value
    For Each elm In dic.Keys
        If dic(elm) = maxCounter Then
            Math_ModeSingle = elm
            Exit Function
        End If
    Next

End Function


'If there is too little overlap
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

' Return an array of mode values.
' About 95% the same as the method above.
Public Function Math_ModeMulti()

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim cnt&: cnt = 1
    For Each elm In Me.arr
        If Not (IsEmpty(elm)) Then
            dic(elm) = dic(elm) + 1
        End If
        cnt = cnt + 1
        If cnt >= 2500 Then
            Math_ModeMulti = ModeMulti_LowOverlap
            Exit Function
        End If
    Next

    If dic.Count = 0 Then Exit Function

    Dim maxCounter As Long
    Dim val As Long
    For Each elm In dic.Items
        maxCounter = IIf(maxCounter < elm, elm, maxCounter)
    Next elm

    Dim newlist As List: Set newlist = New List
    For Each elm In dic.Keys
        If dic(elm) = maxCounter Then
            Call newlist.Add(elm)
        End If
    Next

    Math_ModeMulti = newlist.ToArray

End Function

'If there is too little overlap
Private Function ModeMulti_LowOverlap()
   
    Dim resultBuffer As List: Set resultBuffer = New List
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

