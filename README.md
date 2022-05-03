日本語はこちらから→ [こちら](/README_Japanese.md)  
The following is a machine translation. Therefore, the translation may not be accurate.  
# VBA_ModernList
This is an extremely fast and highly functional one-way concatenated list that runs in ExcelVBA (Windows).   
*Versions that work in other environments will be added in the future. Probably.  
  
VBA_ModernList uses arrays in its internal implementation and operates at speeds similar to pure arrays.
 ![](/BenchMark.png?raw=true)   

# Features  
･It can be used without adding reference settings.  
･It has methods required for lists, such as sorting, uniqueness, range extraction, etc., while in a list.  
･It is almost theoretically fast (more than one order of magnitude faster than Dictionary or ArrayList).  
･It is possible to chain methods together and write everything from adding values to processing output in a single line (a feature known as LINQ in .Net).  
･DebugPrint can be attached anywhere in the method chain to check the contents.  
･Extremely diverse arrayization is included for use with Excel sheets (two-dimensional arrayization, inverted two-dimensional arrayization, etc.).  
･It includes a StringBuilder for fast string merging and a method to convert to CSV in a single line.  
･It is possible to calculate beyond the range of WokrSheetFunction (Sum,Average,Median,Max,Min,StDevP,Mode).  
 
 # Demonstration  
 ' This is a demonstration of how this can be done. If we are actually going to do this, shouldn't we assign every few to an explanatory variable?

```VBA
    Set List1 = New List
    Set List2 = New List
   
    Dim CSV As String
    CSV = List1.CreateEnumRange(1, 150) _
        .IntersectToList(List2.CreateSeqNumbers(40, 200)) _
        .SortByDescending _
        .Slice(20, 100) _
        .DebugPrint("Before conversions => ", "#,##0") _
        .MAP("x", "floor(x*PI()*2,10)") _
        .Filter("x", "Mod(x,20)=0") _
        .DebugPrint("After conversions => ", "0.000", "!!!!!") _
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
    'Enumerate all currently stored values separated by 3-digit commas
    'Projection processing is performed on ListY (Evaluate is used, and the projected list is called ListZ) *In this case, pi is multiplied by 2 and rounded down to a multiple of 10.
    'Filter ListZ and make another list (ListA), in this case only multiples of 20.
    'Enumerate all currently stored values to 3 decimal places
    'Delete duplicates from ListA (to be ListB)
    'Create a string from ListB with tab as separator and Cr as newline code (comma and CrLf in the case of no argument, implementation is StringBuilder, so it's fast!)
    ```

 
 
# Details
 
It can be used by creating an instance as follows.
```VBA
Dim list as List:Set list = new List
```  
*It may be easier to add MethodTest.cls to see the contents of the following.

Adding Values
```VBA
List.Add(1)
call List.AddRange(Array(1,2,3,4,5))
call List.AddRange(Range("A1:A100").Value)
call List.Concat(List)) 'Add another list values
```  

Obtaining Values
```VBA
val = List.GetValueOfIndex(0) 'Get the 0th element
arr = List.ToArray 'Get all values as an Array, usually used when doing a For each.
set newList = List.Slice(5,10) 'Get the 5th through 10th elements in a separate List
set newList = List.PopRange(5,10) 'Get the 5th through 10th elements in a separate List and delete values.
```  

Sorting
```VBA
call List.Sort 'Ascending 
call List.SortByDescending 'Descending
```  

Uniqueness
```VBA
set newList = List.DistinctToList 'The original list is retained.
```  

  
Debug
```VBA
call List.DebugPrint("Prefix","#,##0 etc.","suffix")
```  

Method chain
(most methods return themselves (List type), so the next method can be issued without modification)
```VBA
Call List.CreateSeqNumbers(0, 5) _
        .DebugPrint("Before MAP=> ") _
        .MAP("x", "x*PI()") _
        .DebugPrint("After Map=> ")
        
        'Create a sequential number from 0 to 5
        'Display stored values
        'multiply all stored values by Π
        'Display stored values
```  

Character string association
```VBA
str = List.ToBuildString  'Combines all stored values
str = List.ToBuildString(",") 'Combines all stored values (with "," used as element delimiter)
csv = List.ToBuildSCSV(5,vbTab,vbCr) 'Combine into a CSV format string of 5 line feeds, delimiter tab, and line feed string Cr.
```  

Arrayed 
(you should be able to see it in a watch window in MethodTest, etc.)
```VBA
* n is the number of stored elements -1
str = List.ToArray ' Array(n)
str = List.ToTransposeArray ' Array(0,n)  Useful for vertical pasting in Excel
str = List.To2DArray(i) ' Array( n/i ,0 to i ) The number of columns can be specified.
str = List.To2DTransposeArray(i) 'Array( 0 to i ,n/i)  The number of Rows can be specified.
str = List.ToChunkArray(i) ' Array((n/i)(0 to i))  Useful for use in For each
str = List.ToTransposeChunkArray' Array((n/i)(0,0 to i) Useful for pasting into Excel with For each
str = List.ToDivideArray(i) ' Array((0 to i)(i/n)) Useful when you know how many arrays you want to divide into.
```  
Set theory
```VBA
str = List.UnionToList 
str = List.ExceptToList
str = List.IntersectToList 
```  
 
# License
[MIT license](https://en.wikipedia.org/wiki/MIT_License).
