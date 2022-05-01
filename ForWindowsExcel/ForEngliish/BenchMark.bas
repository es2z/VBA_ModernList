Attribute VB_Name = "BenchMark"
Private startTime As Date

Sub BenchMark()
  
   'The results are displayed in the Immediate window, so press Ctrl+G.
  
    Dim testCount&: testCount = 1000000
    Call Benchmark_Collection(testCount / 200) 'Impossibly slow! This method reduces the speed by O(n^2).   e.g. A(10), B(1000) then  B is 10,000 times slower than A
    Call Benchmark_Dictionary(testCount / 20) 'Moderately slow
    Call Benchmark_List(testCount)
    Call Benchmark_Array(testCount)
    ''Call Benchmark_ArrayList(testCount) 'Not available in environments where Dot Net Framework 3.5 is not enabled.
   
'------------------My laptop results--------------------
    'Specs
    'Windows 10 64bit pro 21H1
    'Intel i5 8265U @ 1.6~3.9Ghz 4C8T (All cores about 3.0Ghz)
    'Memory 8GB One Way (Frequency and timing unknown)

        'Add 1M(1,000K)
            'Collection:0.1211Sec
            'Dictionary:17.047Sec
            'List            :0.1172Sec(Add)
            'List            :0.0313Sec (AddRange)
            'Array         :0.0156Sec
            'ArrayList  :3.9805Sec
            'ArrayList  :0.0078Sec(AddRange) *Only joins of ArrayLists are allowed.

        'Get 1M(1,000K)
            'Collection    :2530.500Sec (Derived from 100K)
            'Dictionary    :0.0230Sec
            'List               :0.0273Sec
            'Array            :0.0156Sec
            'ArrayList        :5.3594Sec
           
        'Add 10M(10,000K)
            'List           :1.2578Sec(Add)
            'List           :0.2969Sec(AddRange)
            'Array        :0.1406Sec
            'ArrayList :40.7891Sec
            'ArrayList :0.0508Sec(AddRange) *Only joins of ArrayLists are allowed.
           
        'Get 10M(10,000K)
            'List            :0.2969Sec
            'Array         :0.1328Sec
            'ArrayList :5.3594Sec
'-----------------------------------------------------------------------
End Sub

Private Sub Benchmark_Collection(testCount As Long)

    Dim collection As collection: Set collection = New collection
   
    startTime = Timer
    For i = 1 To testCount
        collection.Add (i)
    Next
    Debug.Print "Collection Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
   
    startTime = Timer
    For i = 1 To collection.Count
        collection.Item (i)
    Next
    Debug.Print "Collection Get" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
    Debug.Print
   
End Sub

Private Sub Benchmark_Dictionary(testCount As Long)

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
   
    startTime = Timer
    For i = 1 To testCount
        dic.Add i, i
    Next
    Debug.Print "Dictionary Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
   
    startTime = Timer
    For Each Item In dic.Keys
    Next
    Debug.Print "Dictionary Get" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
    Debug.Print
   
End Sub

Private Sub Benchmark_List(testCount As Long)

    Dim List As List: Set List = New List
    startTime = Timer
    Dim i&
    For i = 1 To testCount
        List.Add (i)
    Next
    Debug.Print "List Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec") & "(Add)"
   

    Dim List2 As List: Set List2 = New List
    Dim buffer:  buffer = List.ToArray
   
    Dim List3 As List: Set List3 = New List
    startTime = Timer
    Call List3.AddRange(buffer)
    Debug.Print "List Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec") & "(AddRange)"
   
   'cube root
    Dim cnt&: cnt = WorksheetFunction.RoundUp(testCount ^ (1 / 3), 0)
    Dim buffer2(): ReDim buffer2(1 To cnt, 1 To cnt, 1 To cnt)
    Dim j&, k&, l&: l = 1
    For i = 1 To cnt
        For j = 1 To cnt
            For k = 1 To cnt
                buffer2(i, j, k) = l
                l = l + 1
            Next
        Next
    Next

    Dim list4 As List: Set list4 = New List
    startTime = Timer
    Call list4.AddRange(buffer2)
    Debug.Print "List Add" & Format((cnt - 1) ^ 3 / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec") & "(AddRange(3Dimensions))"
   
   
   
   
    startTime = Timer
    For Each Item In List.ToArray
    Next
    Debug.Print "List Get" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
    Debug.Print
   
End Sub

Private Sub Benchmark_Array(testCount As Long)

    Dim arr(): ReDim arr(testCount - 1)
   
    startTime = Timer
    For i = 0 To testCount - 1
        arr(i) = i
    Next
    Debug.Print "Array Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
   
    startTime = Timer
    For Each Item In arr
    Next
    Debug.Print "Array Get" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")
    Debug.Print
   
End Sub

'Not available in environments where Dot Net Framework 3.5 is not enabled.
Private Sub Benchmark_ArrayList(testCount As Long)

    Dim arrayList1 As Object: Set arrayList1 = CreateObject("System.Collections.ArrayList")

    startTime = Timer
    For i = 0 To testCount - 1
        arrayList1.Add (i)
    Next
    Debug.Print "ArrayList Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")

     Dim arrayList2 As Object: Set arrayList2 = CreateObject("System.Collections.ArrayList")
    startTime = Timer
    Call arrayList2.AddRange(arrayList1)
    Debug.Print "ArrayList Add" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec") & "(AddRange) *Only joins of ArrayLists are allowed."

    startTime = Timer
    For Each Item In arrayList1
    Next
    Debug.Print "ArrayList Get" & Format(testCount / 1000, "#,##0") & "K(" & Format(Int(testCount / 1000000), "#,##0") & "M):" & Format(Timer - startTime, "0.0000Sec")

End Sub
