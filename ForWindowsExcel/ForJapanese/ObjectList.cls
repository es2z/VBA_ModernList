VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'�ɂ߂č����A���@�\�ɓ����g�ݍ��݌^�Ŏg�p�ł���P�����A�ڃ��X�g(�̂悤�ɐU�镑������)�ł��B
'Sort�⃆�j�[�N���ALinq�l�̃��\�b�h�`�F�[����l�X�Ȕz��ɍĔz�u������A�����o���A�W���ɑΉ����Ă��܂��B
'������̃N���X�͎���N���X�����ŁA���p�ɂ�List.cls�̃C���|�[�g���K�v�ł�(List�^�ɕϊ����郁�\�b�h���L�邽��)
'������́AList.cls�ɔ�ׂăf�o�b�O�������Ȃ�ɂȂ��Ă����Ԃł��B��肪����Ε񍐓����肢���܂��B

'Copyright (c) 2022 Kazuki Horike
'Released under the MIT license
'https://github.com/es2z/VBA_ModernList/blob/main/LICENSE

'������̃��|�W�g���͈ȉ��ɂȂ�܂��B�X�^�[�Ƃ�������Ί������ł�
'https://github.com/es2z/VBA_ModernList

Option Base 0
Private myArr() As Variant '�����ɒl������܂�
Private myCount As Long '�z��̗v�f���ł��B
Private myCapacity As Long '�z��̍ő�T�C�Y�ł��B

Property Get arr() As Variant: arr = myArr: End Property
Property Get Count() As Long: Count = myCount: End Property
Property Get Capacity() As Long: Capacity = myCapacity: End Property

'�R���X�g���N�^�ł��B�N���X�̏��������s���܂��B �z��̏����T�C�Y���߂��������̂ł����A�d�l������n���Ȃ��݂����ł�()
Private Sub Class_Initialize()
    myCapacity = 0
    ReDim myArr(myCapacity)
End Sub

'�f�X�g���N�^�ł��A���������J�����܂��BVariant()��nothing�������邱�Ƃ͂ł��܂���
Private Sub Class_Terminate()
    Call Clear
    VBA.DoEvents
End Sub

'�N���X�����������܂��B
Public Function Clear() As ObjectList
    myCount = 0
    Call Class_Initialize
    Set Clear = Me '���\�b�h�`�F�[���̈�
End Function

'�l�����X�g�ɒǉ����܂��B
'���쑬�x�ƃ������̊ϓ_����A�i�[�A�C�e�������z��̃L���p�V�e�B�𒴂����ꍇ�̂�ReDim Preserve���s���L���p�V�e�B��{�ɂ��܂��B
Public Sub Add(obj As Variant)
   
    If myCount > myCapacity Then
        myCapacity = (myCapacity + 1) * 2 - 1
        ReDim Preserve myArr(myCapacity)
    End If
   
    Set myArr(myCount) = obj
    myCount = myCount + 1
   
End Sub

'�z���ΏۂƂ��������Ȕ͈͒ǉ����s���܂��B �\�ߔz��͈͂��m�ۂ��Ă����ɒl�����Ă����܂��B
Public Function AddRange(args As Variant) As ObjectList
   
   '����������
   Dim dimensionsCount&: dimensionsCount = 1
   Dim buffer&
On Error GoTo err
    While True
        buffer = UBound(args, dimensionsCount)
        dimensionsCount = dimensionsCount + 1
    Wend
err:
    dimensionsCount = dimensionsCount - 1

    '�͈͂̊m��
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
   
   '�ǉ�
    For Each arg In args
        Set myArr(insertPoint) = arg
        insertPoint = insertPoint + 1
    Next
    
    Set AddRange = Me

End Function

'2�̃��X�g���������܂��B�����ɓ��ꂽ�ق������Ƃɗ��܂�
Public Function Concat(argList As ObjectList) As ObjectList

    If argList.Count = 0 Then Exit Function
    Call AddRangeForOneDimension(argList.ToArray)
   
    Set Concat = Me

End Function

'ForEach�\�ȃA�C�e���ɂ��Ēǉ�����B���ł��Ή��\�����A���x�I��AddRange��Concat�ɗ��
Public Function AddIterableItems(args) As ObjectList

    For Each arg In args
        Call Me.Add(arg)
    Next
    
    Set AddIterableItems = Me
    
End Function


'�ŏ��̗v�f���擾���܂��B
Public Function First(): Set First = IIf(myCount = 0, Empty, myArr(0)): End Function
'�Ō�̗v�f���擾���܂��B
Public Function Last(): Set Last = IIf(myCount = 0, Empty, myArr(myCount - 1)): End Function

'�����v�f���i�[����Ă����True��Ԃ��܂��B
Public Function Any_(): Any_ = myCount >= 1: End Function

'�����v�f���i�[����Ă��Ȃ����True��Ԃ��܂��B
Public Function Nothing_(): Nothing_ = myCount = 0: End Function

'�z��ɕϊ����ĕԂ��܂��B �z��̗v�f���̓f�[�^���ɐ؂�l�߂��܂��B
Public Function ToArray() As Variant()

    Dim bufArr(): ReDim bufArr(myCount - 1)
    Dim i&
    For i = 0 To (myCount - 1)
        Set bufArr(i) = myArr(i)
    Next
   
    ToArray = bufArr
   
End Function

'chunkSize�̌����Ƃɕ������z���z��̒��ɓ���ĕԂ��܂��B ��100�̃f�[�^��40�ŕ������ꍇArray(Array(40),Array(40),Array(20))
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


'�I�u�W�F�N�g�𕡐����܂��B������ShallowCopy�ł��B
'DeepCopy����������ɂ́A�I�u�W�F�N�g����DeepCopy���������āA������Select���\�b�h���g���Ȃǂłł��邩�ȂƎv���܂��B
Public Function Clone() As ObjectList

    Dim newlist As ObjectList: Set newlist = New ObjectList
    Set Clone = newlist.Concat(Me)

End Function

'�v�f������ɏ����Ń\�[�g���܂��B(�����͑����N�C�b�N�\�[�g�ɂȂ��Ă���͂��ł�)
'�L���b�V��������č��������܂����B
Public Function SortByElementNameIn_Ascending(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
   
    '�����̃����o�[����������Ȃ������ꍇ
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception

    Const minindex& = 0
    Dim maxIndex&: maxIndex = myCount - 1
    Dim cache: ReDim cache(minindex To maxIndex)
    Dim i&
    
    '�\�[�g�ΏۂƂȂ鑮�����L���b�V������B���̓��e���r���ăN�C�b�N�\�[�g����`�A�L���b�V�����Ȃ���~1���d���Ȃ邩��
    For i = minindex To maxIndex
        cache(i) = CallByName(arr(i), memberName, VbGet)
    Next
   
    Call SortBy(cache, minindex, maxIndex)
    Set SortByElementNameIn_Ascending = Me

Exit Function

Exception:
Call MsgBox("���X�g�ɑ΂��ă\�[�g�����s���悤�Ƃ��܂������A�����̃����o�[����������܂���ł���", vbCritical)

End Function
Private Sub SortBy(ByRef cache, ByVal minindex As Long, ByVal maxIndex As Long)

    Dim low As Long, high As Long
    Dim pivot As Variant
   
    pivot = cache(Int((minindex + maxIndex) / 2))
   
    low = minindex
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

    If (minindex < low - 1) Then
        Call SortBy(cache, minindex, low - 1)
    End If
    If (maxIndex > high + 1) Then
        Call SortBy(cache, high + 1, maxIndex)
    End If
   
End Sub


'�v�f������ɏ����Ń\�[�g���܂��B(�����͑����N�C�b�N�\�[�g�ɂȂ��Ă���͂��ł�
'�L���b�V��������č��������Ă��܂��B
Public Function SortByElementNameIn_Descending(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
   
    '�����̃����o�[����������Ȃ������ꍇ
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception

    Const minindex& = 0
    Dim maxIndex&: maxIndex = myCount - 1
    Dim cache: ReDim cache(minindex To maxIndex)
    Dim i&
    
    '�\�[�g�ΏۂƂȂ鑮�����L���b�V������B���̓��e���r���ăN�C�b�N�\�[�g����`�A�L���b�V�����Ȃ���~1���d���Ȃ邩��
    For i = minindex To maxIndex
        cache(i) = CallByName(arr(i), memberName, VbGet)
    Next
   
    Call SortByDescending(cache, minindex, maxIndex)
    Set SortByElementNameIn_Descending = Me

Exit Function

Exception:
Call MsgBox("���X�g�ɑ΂��ă\�[�g�����s���悤�Ƃ��܂������A�����̃����o�[����������܂���ł���", vbCritical)

End Function
Private Sub SortByDescending(ByRef cache, ByVal minindex As Long, ByVal maxIndex As Long)

    Dim low As Long, high As Long
    Dim pivot As Variant
   
    pivot = cache(Int((minindex + maxIndex) / 2))
   
    low = minindex
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

    If (minindex < low - 1) Then
        Call SortByDescending(cache, minindex, low - 1)
    End If
    If (maxIndex > high + 1) Then
        Call SortByDescending(cache, high + 1, maxIndex)
    End If
   
End Sub

 '�C���f�b�N�X�l�ɑΉ�����l���擾���܂��B�z���[n]�Ɠ����ł��B
 Public Function GetItmeOfIndex(index As Long)
 
    If index > myCount Then
         Set GetItmeOfIndex = Null
    Else
        Set GetItmeOfIndex = myArr(index)
    End If
   
End Function

'����͈̔͂𔲂��o���ĐV�������X�g�Ƃ��ĕԂ��܂��B(minIndex<=�����o������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
Public Function Slice(minindex As Long, maxIndex As Long) As ObjectList

    If minindex > maxIndex Then
        Set Slice = New ObjectList
        Exit Function
    End If
   
    minindex = IIf(minindex < LBound(myArr), LBound(myArr), minindex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)
   
    Dim newlist As ObjectList: Set newlist = New ObjectList
   
    For i = minindex To maxIndex
        Call newlist.Add(myArr(i))
    Next i
   
    Set Slice = newlist

End Function

'�w�萔�擪���甲���o��
Public Function TakeToList(itemCount As Long) As ObjectList

    '�w�萔���v�f���𒴂��Ă���ꍇ�N���[����Ԃ��܂�
    Dim minindex As Long: minindex = LBound(myArr)
    If itemCount + 1 - minindex > myCount Then
        Set TakeToList = Me.DeepCopy()
    Else
        Dim newlist As ObjectList: Set newlist = New ObjectList
        For i = minindex To itemCount - 1 + minindex
            Call newlist.Add(myArr(i))
        Next i
        Set TakeToList = newlist
    End If

End Function

'�w�萔���X�L�b�v���Ă���ȍ~�𔲂��o��
Public Function SkipToList(skipCount As Long) As ObjectList

    '�w�萔���v�f���𒴂��Ă���ꍇ���ObjectList��Ԃ��܂��B
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

'�i�[����Ă���l�𔽓]���܂��B
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

' �l���V���b�t�����܂��B���x�͂��܂�����������܂���B���x���~������΃����Z���k�c�C�X�^�[�Ƃ��ǂ��炵���̂Ŏ������ĉ�����()
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

'��ӂ̒l(���j�[�N�l)�Ƀt�B���^�[���ĐV�������X�g�Ƃ��ĕԂ��܂��B
Public Function DistinctToListByElementName(memberName As String) As ObjectList

    If myCount < 1 Then Exit Function
       
    '�����̃����o�[����������Ȃ������ꍇ
    On Error Resume Next
        buf = CallByName(myArr(0), memberName, VbGet)
    On Error GoTo Exception
        
    Const minindex& = 0
    Dim maxIndex&: maxIndex = myCount - 1
    Dim cache: ReDim cache(minindex To maxIndex)
    Dim i&
    
    '�\�[�g�ΏۂƂȂ鑮�����L���b�V������B���̓��e���r���ăf�B�N�V���i���ɓ����
    For i = minindex To maxIndex
        cache(i) = CallByName(arr(i), memberName, VbGet)
    Next
   
    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim newlist As ObjectList: Set newlist = New ObjectList
   
    For j = minindex To maxIndex
        Set dic(cache(j)) = myArr(j)
    Next j
   
    For Each Item In dic.Items
        Call newlist.Add(Item)
    Next
   
    Set DistinctToListByElementName = newlist

Exit Function

Exception:
    Call MsgBox("���X�g�ɑ΂��ă��j�[�N�������s���悤�Ƃ��܂������A�����̃����o�[����������܂���ł���", vbCritical)
End Function

'���X�g���I�u�W�F�N�g�̔C�ӗv�f�ɑ΂��Đ��K�\���Ńt�B���^�[���s���A�q�b�g�������̂�V�������X�g�Ɋi�[���ĕԂ��B
Public Function FilterRegExpInElementName(ElementName As String, RegexpPattern As String, Optional IgnoreCase As Boolean = False) As ObjectList

    '�����̃����o�[����������Ȃ������ꍇ
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
    Call MsgBox("���X�g�ɑ΂��ăt�B���^�[����(Filter)�����s���悤�Ƃ��܂������A�����̃����o�[����������܂���ł���", vbCritical)
   
End Function

'���X�g���I�u�W�F�N�g�̔C�ӗv�f��V����list(�l�^)�Ƃ��ĕԂ��B
'���[�U�[��`�^�ȊO���Ƃ��܂�����肭�����Ȃ��݂����ł�
Public Function SliceElementByName(ElementName As String) As list

    '�����̃����o�[����������Ȃ������ꍇ
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
    Call MsgBox("���X�g�ɑ΂��Đ؂�o������(Slice)�����s���悤�Ƃ��܂������A�����̃����o�[����������܂���ł���", vbCritical)
   
End Function