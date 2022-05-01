Attribute VB_Name = "MethodTest"

'Comments are machine translated by DeepL translation.
'It may be more accurate to read the code for the actual behavior.

'You can check how each method is working by using Ctrl+G.

Private List As List, List1 As List, List2 As List, List3 As List

Public Sub AllTestExecute()
    Call AllTest
End Sub


' This is a demonstration of how this can be done. If we are actually going to do this, shouldn't we assign every few to an explanatory variable?
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
       
    ' DebugPrint anywhere in the 'method chain' to see the contents!!
   
    'What we are doing.
    'Adding consecutive values to List1 and 2
    '*List1 is in (start number, number to be created) format, while List2 creates a sequential number in n = i to j (step k) format, where n is the sequential number to be created and the arguments are i,j,k.
    'Create another list (say ListX) by creating the product set of List1 and 2 (leaving only the values that are in both)
    'Sort ListX in descending order
    'Slice ListX by range and return another list (say ListY)
    'Projection processing is performed on ListY (Evaluate is used, and the projected list is called ListZ) *In this case, pi is multiplied by 2 and rounded down to a multiple of 10.
    'Filter ListZ and make another list (ListA), in this case only multiples of 20.
    'Delete duplicates from ListA (to be ListB)
    'Create a string from ListB with tab as separator and Cr as newline code (comma and CrLf in the case of no argument, implementation is StringBuilder, so it's fast!)
    
End Sub

' Displays the contents of the list. "It is also possible to insert it in the middle of a method chain expression!!!"
Private Sub DebugPrintTest()
    Set List = New List
    Call List.CreateEnumRange(1000000000, 3) 'Get 3 values starting with 500 million
    Call List.DebugPrint("Debug.Print =>  I want $", "#,##0", " !!!")
End Sub

' Add data.
Private Sub AddTest()

    Set List = New List
    For i = 0 To 3
        List.Add (i)
    Next i
   
    Call List.DebugPrint("Add => ")

End Sub

'Initialize 'list
Private Sub ClearTest()

    Set List = New List
    Call List.CreateEnumRange(1, 5)
    Call List.DebugPrint("ClearTest1=> ")
    Call List.Clear.DebugPrint("ClearTest2=> ") 'Not shown because there is no content

End Sub

'Add multiple data. Arrays are the target.
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

' Combines two listings. The one you put in the argument comes after.
Private Sub ConcatTest()

    Set List1 = New List
    Set List2 = New List
    
    Call List1.CreateEnumRange(start:=5, Count:=5) 'Create 5 sequential numbers starting with 5
    Call List2.CreateSeqNumbers(First:=5, Last:=100, step:=20) 'Create a sequential number in the form for i = j step k
    Call List1.Concat(List2).DebugPrint("Concat=> ", "#,##0")
    
End Sub

' Get the first element.
Private Sub FirstTest()
    Set List = New List
    Debug.Print "First=> " & List.CreateEnumRange(100, 5).First
End Sub

'Get the last element
Private Sub LastTest()
    Set List = New List
    Debug.Print "First=> " & List.CreateEnumRange(100, 5).Last
End Sub

' Return True if any element is stored.
Private Sub AnyTest()
    Set List = New List
    Debug.Print "Any1=> " & List.Any_
    List.Add (1)
    Debug.Print "Any2=> " & List.Any_
End Sub

' Return True if no elements are stored.
Private Sub NothingTest()
    Set List = New List
    Debug.Print "Nothing1=> " & List.Nothing_
    List.Add (1)
    Debug.Print "Nothing2=> " & List.Nothing_
End Sub

' Return True if the contents of the two lists all match.
Private Sub SequenceEqualTest()
    Set List1 = New List
    Set List2 = New List
    
    Call List1.CreateEnumRange(5, 5)
    Call List2.CreateEnumRange(5, 5)
    Debug.Print "SequenceEqual1=> " & List1.SequenceEqual(List2)
    Debug.Print "SequenceEqual2=> " & List1.SequenceEqual(List2.Clear)
End Sub

' Clears the contents and creates a sequential number. It is similar in form to a for statement.
Private Sub CreateSeqNumbersTest()
    Set List = New List
    Call List.CreateSeqNumbers(First:=100, Last:=500, step:=80).DebugPrint("CreateSeqNumbers=> ")
End Sub

'Clears the contents and creates a sequential number for the specified number of pieces from the starting position
Private Sub CreateEnumRangeTest()
    Set List = New List
    Call List.CreateEnumRange(start:=100, Count:=5).DebugPrint("CreateEnumRange=> ")
End Sub

 ' Get the value corresponding to the index value. Same as [index] in the array.
Private Sub GetValueOfIndexTest()
    Set List = New List
    Call List.CreateEnumRange(100, 5)
    Debug.Print "GetValueOfIndexTest=> " & List.GetValueOfIndex(3)
End Sub

'Extracts a specific range and returns it as a new list. (minIndex<=what is extracted<=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored is targeted for retrieval.
Private Sub SliceTest()
    Set List = New List
    Call List.CreateEnumRange(100, 20)
    Call List.Slice(minIndex:=-50, maxIndex:=3).DebugPrint("SliceTest=> ")
End Sub

'Remove the element corresponding to the index and pack the data in front of it. It is not efficient.
Private Sub RemoveTest()
    Set List = New List
    Call List.CreateEnumRange(1, 5)
    Call List.Remove(2).DebugPrint("Remove => ")
End Sub

' Removes all matches to the numerical value of the argument and prepends the data.
' (The implementation creates a new list from scratch.)
Private Sub RemoveAllTest()
    Set List = New List
    Call List.AddRange(Array(5, 3, 2, 5, 3, 2, 3, 4, 3))
    Call List.RemoveAll(3).DebugPrint("RemoveAll => ")
End Sub

'Delete a specific range.(minIndex<=what is deleted <=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored  is subject to deletion.
Private Sub RemoveRangeTest()
    Set List = New List
    Call List.AddRange(Array(5, 3, 2, 5, 3, 2, 3, 4, 3))
    Call List.RemoveRange(minIndex:=4, maxIndex:=15).DebugPrint("RemoveRange => ")
End Sub

' Returns the element corresponding to the index and then removes that element from the list.
' The return value is the value obtained.
Private Sub PopTest()
    Set List = New List
    Call List.CreateEnumRange(100, 5)
    Debug.Print "Pop(Get Value)=> "; List.Pop(3)
    Call List.DebugPrint("pop(Remaining values)=> ")
End Sub

'Retrieve a specific range of values and simultaneously delete the targeted values. (minIndex<=deleted<=maxIndex)
'If Min is unreasonably small or Max is unreasonably large, only the range where the value is stored is targeted for retrieval.
' The return value is the value obtained.
Private Sub PopRangeTest()
    Set List = New List
    Call List.CreateEnumRange(100, 6)
    Call List.PopRange(minIndex:=2, maxIndex:=4).DebugPrint("PopRange(Get Values)=> ")
    Call List.DebugPrint("PopRange(Remaining values)=> ")
End Sub

' Converted to an array and returned. The number of elements in the array is truncated to the stored value.
Private Sub ToArrayTest()
    Set List = New List
    For Each elm In List.CreateEnumRange(1, 5).ToArray
        Debug.Print "Toarray=> " & elm
    Next
End Sub

'Converts to a 2-dimensional array and returns it with values in the 1-dimensional elements. The number of array elements is truncated to the number of data.
'(It can be conveniently used to attach data to an Excel sheet.)
Private Sub ToTransposeArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 5).ToTransposeArray
    'Example. Range("A1:A5").Value = buffer
    For Each elm In buffer
        Debug.Print "ToTransposeArray=> " & elm; "*Watch window required"
    Next
   
End Sub

' Converts to a 2-dimensional array and returns it. You can set the number of the second dimension.
Private Sub To2DArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 10).To2DArray(columnCount:=5)
    For Each elm In buffer
        Debug.Print "To2DArray=> " & elm; "*Watch window required"
    Next
   
End Sub

' Converts to a 2-dimensional array and returns it. You can set the number of the first dimension.
Private Sub To2DTransposeArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 10).To2DTransposeArray(rowCount:=5)
    For Each elm In buffer
        Debug.Print "To2DTransposeArray=> " & elm; "*Watch window required"
    Next
   
End Sub

'Returns an array divided by the number of 'chunkSize' in the array.
'Example 100 data divided by 40 Array(Array(39),Array(39),Array(39(the last 20 are empty)))
Private Sub ToChunkArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToChunkArray(6)
    Debug.Print "Chunk=> "
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab) 'ƒ^ƒu‹æØ‚è‚Ì1s‚É‚·‚é
Next
   
End Sub

'Transpose the array in Chunk and make it a 2-dimensional array.
'(This was useful when turning and pasting a For(Each).
'Example 100 data divided by 40 Array(Array(0,39),Array(0,39),Array(0,39(the last 20 are empty)))
Private Sub ToTransposeChunkArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240).ToTransposeChunkArray(6)
    Debug.Print "ToTransposeChunkArray=> *Watch window required"
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab)
Next
   
End Sub

'Contrary to 'chunk', you can decide how many pieces of data to divide into and place them evenly there.
'Example: 100 data divided by 3 Array(Array(33),Array(33 last empty),Array(33 last empty))
Private Sub ToDivideArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToDivideArray(6)
    Debug.Print "ToDivideArray=> "
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab)
Next
   
End Sub

'Returns a Dictionary with only the 'Keys' set. Duplicates will be removed, which may be useful if you have a large number of exists and so on.
Private Sub ToHashsetTest()
   
    Set List = New List
    Dim dic: Set dic = List.CreateSeqNumbers(100, 150, 10).ToHashset

    For Each Key In dic.Keys
        Debug.Print "ToHashset=> " & Key
    Next
   
End Sub

'Creates a union set (combined and unique) of lists and returns it in a new list
Private Sub UnionToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 60, 10)
    Call List2.CreateSeqNumbers(0, 60, 12)
    Call List1.UnionToList(List2).DebugPrint("UnionToList=> ")
   
End Sub

'Creates a difference set (what was originally there minus what is in the argument) between lists and returns it in a new list.
Private Sub ExceptToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 15, 2)
    Call List2.CreateSeqNumbers(0, 15, 3)
    Call List1.ExceptToList(List2).DebugPrint("ExceptToList=> ")
   
End Sub

' Creates a product set (what is in both) between lists and returns it in a new list.
Private Sub IntersectToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 24, 2)
    Call List2.CreateSeqNumbers(0, 24, 3)
    Call List1.IntersectToList(List2).DebugPrint("IntersectToList=> ")
   
End Sub

'Duplicate (shallow copy) an List.
Private Sub CloneTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 20)
    Call List.Clone.DebugPrint("Clone=> ")
End Sub

'Extracts the specified number of values from the beginning.
Private Sub TakeToListTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 10)
    Call List.TakeToList(5).DebugPrint("TakeToList=> ")
End Sub

'skip the specified number of values and extract the rest.
Private Sub SkipToListTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 10)
    Call List.SkipToList(5).DebugPrint("SkipToList=> ")
End Sub

'Shuffles stored values.Accuracy may not be good.
Private Sub RandamizeTest()
    Set List = New List
    Call List.CreateSeqNumbers(5, 10)
    Call List.Randamize.DebugPrint("Rndamize=> ")
End Sub

'Sort in ascending order. (Implementation is quick sort.)
Private Sub SortTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Randamize.Sort.DebugPrint("Sort=> ")
End Sub

' Sort in descending order.
Private Sub SortByDescendingTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Randamize.SortByDescending.DebugPrint("SortByDescending=> ")
End Sub

' Reverses the order of stored values.
Private Sub ReverseTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Reverse.DebugPrint("Reverse=> ")
End Sub

' Assumes the stored value is a string and inspects whether the target is included in the list.
Private Sub StringContainsTest()
    Set List = New List
    Call List.AddRange(Array(1, 2, 3, 4, 5, 487, "AAAA"))
    Debug.Print "StringContains1 =>  " & List.StringContains("1")
    Debug.Print "StringContains2 =>  " & List.StringContains("7")
End Sub

' Assuming the stored value to be a string, checks whether the target string is included in the list using a regular expression.
Private Sub StringContains_RegExpTest()
    Set List = New List
    Call List.AddRange(Array(1, 2, 3, 4, 5, 487, "ABCD"))
    Debug.Print "StringContains_RegExp1 =>  " & List.StringContains_RegExp(".*7")
    Debug.Print "StringContains_RegExp2 =>  " & List.StringContains_RegExp("[0-9]{3}")
    Debug.Print "StringContains_RegExp3 =>  " & List.StringContains_RegExp("[a-z]{3,}")
    Debug.Print "StringContains_RegExp4 =>  " & List.StringContains_RegExp("A+")
    Debug.Print "StringContains_RegExp4 =>  " & List.StringContains_RegExp("^87")
End Sub

'Fast concatenates a collection like StringBuilder and returns a String.
' If an argument is specified, the character is given to delimit the collection.
Private Sub ToBuildStringTest()
    Set List = New List
    Call List.AddRange(Array("foo", "bar", "baz", "qux", "quux", "corge", "grault", "garply", "waldo"))
    Debug.Print "ToBuildString1 =>  " & List.ToBuildString()
    Debug.Print "ToBuildString2 =>  " & List.ToBuildString(",")
End Sub

'Create CSV (Character Separated Value ), fast join like StringBuilder.
'If combined with Excel's Range(area).Value and AddRange method, etc., it should work very conveniently for external data output.
Private Sub ToBuildCSVTest()
    Set List = New List
    Call List.AddRange(Array("Name", "Age", "Gender", "Alice", 30, "FeMale", "Bob", 40, "Bob"))
    Debug.Print "ToBuildCSV =>"
    Debug.Print List.ToBuildCSV(3, ",", vbCrLf)
End Sub

'Filter to unique values and return as a new list.
Private Sub DistinctToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 12, 2)
    Call List2.CreateSeqNumbers(0, 12, 3)
    Call List1.Concat(List2).DistinctToList.Sort.DebugPrint("DistinctToList=> ")
   
End Sub

' MAP processing; Excel functions can be used. It is not fast.
'See method test for usage.
Private Sub MapTest()

    Set List = New List
   
      'Multiply pi by 2 and round down to multiples of 10
    Call List.CreateSeqNumbers(0, 5) _
        .DebugPrint("Before MAP=> ") _
        .MAP("x", "floor(x*PI()*2,10)") _
        .DebugPrint("After Map=> ")
   
End Sub

'Filter to unique values and return as a new list.
Private Sub FilterTest()

    Set List = New List
   
    'Only those that divide by 20 and the remainder is zero
    Call List.CreateSeqNumbers(0, 50, 10) _
        .DebugPrint("Before filtering =>") _
        .Filter("x", "Mod(x,20)=0") _
        .DebugPrint("After filtering=> ")
   
End Sub

'From here on down is formula processing.
'If the range can be covered by the built-in functions, they are used to speed up the process.
'Re-arranging into a 2-dimensional array increases the range where the built-in functions can be applied, but does not do so if it slows down the process.

'!!! Since it is assumed that numerical values are stored, it will not work correctly if the elements stored in the list contain strings. !!!
'!!! I use the desktop version of Excel in a Windows environment, so I have not verified that it works in other environments. !!!
Private Sub MathematicalFunctionsTest()

    Set List1 = New List

    Call List1.CreateSeqNumbers(1, 2000000) 'If too large, the application may crash.
    Debug.Print "Sum =>  " & List1.Math_Sum
    Debug.Print "Average =>  " & List1.Math_Average
    Debug.Print "Median =>  " & List1.Math_Median
    Debug.Print "Max =>  " & List1.Math_Max
    Debug.Print "Min =>  " & List1.Math_Min
    Debug.Print "StDevP =>  " & List1.Math_StDevP
   
    'The following are Mode calculations
    Set List2 = New List
    For i = 1 To 100000
        List2.Add (Int(Rnd() * 2147483647))
    Next
    
    For i = 1 To 65535
        For j = 1 To 10
            List2.Add (j)
        Next j
    Next i

    Debug.Print "ModeSingle =>  " & List2.Math_ModeSingle
   
    For Each buf In List2.Math_ModeMulti
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

