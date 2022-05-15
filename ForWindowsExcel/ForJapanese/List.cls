VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'�ɂ߂č����A���@�\�ɓ����g�ݍ��݌^�Ŏg�p�ł���P�����A�ڃ��X�g(�̂悤�ɐU�镑������)�ł��B
'Sort�⃆�j�[�N���ALinq�l�̃��\�b�h�`�F�[����l�X�Ȕz��ɍĔz�u������A�����o���A�W���AStringBuilder�ɂ��A���Ȃǂɂ��Ή����Ă��܂��B
'���\�b�h�`�F�[�����̖w�ǂ̉ӏ���DebugPrint��u���ď�Ԃ��m�F���邱�Ƃ��\�ł��B
'����N���X������ObjectList�ɂȂ�܂��B

'Copyright (c) 2022 Kazuki Horike
'Released under the MIT license
'https://github.com/es2z/VBA_ModernList/blob/main/LICENSE

'������̃��|�W�g���͈ȉ��ɂȂ�܂��B�X�^�[�Ƃ�������Ί������ł�
'https://github.com/es2z/VBA_ModernList

Option Base 0
Private myArr() As Variant '�����ɒl������܂�
Private myCount As Long '�z��̗v�f���ł��B
Private myCapacity As Long '�z��̍ő�T�C�Y�ł��B

Property Get arr(): arr = myArr: End Property 'Array�͗\�񕶎�
Property Get Count() As Long: Count = myCount: End Property
Property Get Capacity() As Long: Capacity = myCapacity: End Property

'�R���X�g���N�^�ł��B�N���X�̏��������s���܂��B �z��̏����T�C�Y���߂��������̂ł����A�d�l������n���Ȃ��݂����ł�()
Private Sub Class_Initialize()
    myCapacity = 0
    ReDim myArr(myCapacity)
End Sub

'�f�X�g���N�^�ł��A���������J�����܂��B
Private Sub Class_Terminate()
    Call Clear
    VBA.DoEvents
End Sub

'�N���X�����������܂��B
Public Function Clear() As list
    myCount = 0
    Call Class_Initialize
    Set Clear = Me '���\�b�h�`�F�[���̈�
End Function

'�f�o�b�O�o�͂ł��B�Ȃ�ƃ��\�b�h�`�F�[���̒��Ŏg���܂�!!
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

'�l�����X�g�ɒǉ����܂��B
'���쑬�x�ƃ������̊ϓ_����A�i�[�A�C�e�������z��̃L���p�V�e�B�𒴂����ꍇ�̂�ReDim Preserve���s���L���p�V�e�B��{�ɂ��܂��B
'�Ȃ��AReDim Preserve�̑��x�͐V�������I�z����m�ۂ��Ēl����������ւ���̂Ƃقړ����ɂȂ�܂�
Public Sub Add(val)

    If myCount > myCapacity Then
        myCapacity = (myCapacity + 1) * 2 - 1
        ReDim Preserve myArr(myCapacity)
    End If

    myArr(myCount) = val
    myCount = myCount + 1

End Sub

'�z���ΏۂƂ��������Ȕ͈͒ǉ����s���܂��B �\�ߔz��͈͂��m�ۂ��Ă����ɒl�����Ă����܂��B
'Excel�Ȃ�List.AddRange(Range("A1:B100").Value)�̂悤�Ȏg�������֗��ł��B
Public Function AddRange(args As Variant) As list
   
   '�������̔���
   Dim dimensionsCount&: dimensionsCount = 1
   Dim buffer&
On Error GoTo err
    While True
        buffer = UBound(args, dimensionsCount)
        dimensionsCount = dimensionsCount + 1
    Wend
err:
    dimensionsCount = dimensionsCount - 1

    '�͈͂��m��
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
   
   '�ǉ�
    For Each arg In args
        myArr(insertPoint) = arg
        insertPoint = insertPoint + 1
    Next
    
    Set AddRange = Me

End Function

'2�̃��X�g���������܂��B�����ɓ��ꂽ�ق������Ƃɗ��܂��B
'ToArray�ł͂Ȃ� Property Get�ő�����̔z�����Ă���̂������܂������A�ُ킾�Ǝv���邮�炢�I�[�o�[�w�b�h������̂łǂ�����Ă�1.5�{���x�x���Ȃ�悤�ł��B
Public Function Concat(argList As list) As list

    If argList.Count = 0 Then
        Set Concat = Me
        Exit Function
    End If
   
    Call AddRange(argList.ToArray)
    Set Concat = Me

End Function

'ForEach�\�ȃA�C�e���ɂ��Ēǉ�����B���ł��Ή��\�����A���x�I��AddRange��Concat�ɗ��
Public Function AddIterableItems(args) As list

    For Each arg In args
        Call Me.Add(arg)
    Next
    
    Set AddIterableItems = Me
    
End Function

'�ŏ��̗v�f���擾���܂��B
Public Function First(): First = IIf(myCount = 0, Empty, myArr(0)): End Function

'�Ō�̗v�f���擾���܂��B
Public Function Last(): Last = IIf(myCount = 0, Empty, myArr(myCount - 1)): End Function

'�����v�f���i�[����Ă����True��Ԃ��܂��B
Public Function Any_(): Any_ = myCount >= 1: End Function

'�����v�f���i�[����Ă��Ȃ����True��Ԃ��܂��B
Public Function Nothing_(): Nothing_ = myCount = 0: End Function

'2�̃��X�g�̓��e���S�Ĉ�v���Ă����True��Ԃ��܂��B
Public Function SequenceEqual(compairTarget As list) As Boolean
   
    If compairTarget.Count <> Me.Count Then Exit Function
   
    Dim targetArr(): targetArr = compairTarget.ToArray
    Dim i&
    For i = 0 To Me.Count - 1
        If Me.arr(i) <> targetArr(i) Then Exit Function
    Next
   
    SequenceEqual = True
   
End Function

'���g���N���A���ĘA�Ԃ��쐬���܂� for���̏������Ɠ����ł��B
Public Function CreateSeqNumbers(First As Long, Last As Long, Optional step As Long = 1) As list

    Me.Clear
    Dim i&
    For i = First To Last Step step
        Me.Add (i)
    Next i
   
    Set CreateSeqNumbers = Me
   
End Function

'���g���N���A����.Net��Enumlabre.Range�̂悤�Ȋ����ŘA�Ԃ��쐬���܂��B
Public Function CreateEnumRange(start As Long, Count As Long) As list

    Me.Clear
    Dim i&
    For i = start To start + Count - 1
        Me.Add (i)
    Next i
   
    Set CreateEnumRange = Me
   
End Function

 '�C���f�b�N�X�l�ɑΉ�����l���擾���܂��B�z���[n]�Ɠ����ł��B
 Public Function GetValueOfIndex(index As Long)

    If index > myCount Then
        GetValueOfIndex = Empty
    Else
        GetValueOfIndex = myArr(index)
    End If

End Function

'����͈̔͂𔲂��o���ĐV�������X�g�Ƃ��ĕԂ��܂��B(minIndex<=�����o������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
Public Function Slice(minindex As Long, maxIndex As Long) As list

    minindex = IIf(minindex < LBound(myArr), LBound(myArr), minindex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minindex > maxIndex Then
        Set Slice = New list
        Exit Function
    End If

    Dim newlist As list: Set newlist = New list
    Dim i&
    For i = minindex To maxIndex
        newlist.Add (myArr(i))
    Next i

    Set Slice = newlist

End Function

'index�ɑΉ�����v�f���폜���ăf�[�^��O�ɋl�߂܂��B�����͈����ł��B
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

'�����̐��l�Ɉ�v������̂�S�č폜���ăf�[�^��O�ɋl�߂܂�(�����͐V�������X�g������ē���ւ��Ă��܂�)
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

'����͈̔͂��폜���ċl�߂܂��B(minIndex<=����������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
Public Function RemoveRange(minindex As Long, maxIndex As Long) As list

    minindex = IIf(minindex < LBound(myArr), LBound(myArr), minindex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minindex > maxIndex Then Exit Function

    Dim difference: difference = maxIndex - minindex
    Dim i&
    For i = minindex To myCount - 1
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

'index�ɑΉ�����v�f��Ԃ�����ŁA���X�g���炻�̗v�f���폜���܂��B
'�Ԃ�l�͎擾�����l�ł��B
Public Function Pop(index As Long) As Variant

    If index < LBound(myArr) Or UBound(myArr) < index Then
        Pop = Empty: Exit Function
    End If

    Pop = myArr(index)
    Call Remove(index)

End Function

'����͈̔͂�Ԃ��Ɠ����ɍ폜���ċl�߂܂��B(minIndex<=����������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
'�Ԃ�l�͎擾�����l�ł��B
Public Function PopRange(minindex As Long, maxIndex As Long) As list

    minindex = IIf(minindex < LBound(myArr), LBound(myArr), minindex)
    maxIndex = IIf(maxIndex > myCount - 1, myCount - 1, maxIndex)

    If minindex > maxIndex Then
        Set PopRange = New list
        Exit Function
    End If

    Dim difference: difference = maxIndex - minindex
    Dim result As list: Set result = New list

    For index = minindex To maxIndex
        Call result.Add(myArr(index))
    Next index

    Dim i&
    For i = minindex To myCount - 1
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

'�z��ɕϊ����ĕԂ��܂��B �z��̗v�f���̓f�[�^���ɐ؂�l�߂��܂��B
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

'2�����z��ɕϊ�����1�����v�f�ɒl�����ĕԂ��܂��B �z��̗v�f���̓f�[�^���ɐ؂�l�߂��܂��B�Z���̓\�t�Ɏg���܂��B
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

'2�����z��ɕϊ����ĕԂ��܂��B 2�����ڂ̌�(Excel�Ȃ���ɑ���)��ݒ�ł��܂��B(�����̍s����擾���Ă��̂܂ܓ\��t����ۂȂǂɕ֗��ł�)
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

'2�����z��ɕϊ����ĕԂ��܂��B 1�����ڂ̌�(Excel�Ȃ�s���ɑ���)��ݒ�ł��܂��B
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

'chunkSize�̌����Ƃɕ������z���z��̒��ɓ���ĕԂ��܂��B ��100�̃f�[�^��40�ŕ������ꍇArray(Array(40),Array(40),Array(40(�㔼20��empty)))
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

'��̂ɉ�����Chunk���̔z���Transpose����2�����z��ɂ��Ă܂�(For(Each)���񂵂ē\�t������ۂɎg���₷�����Ǝv���܂��B)
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

'chunk�Ƃ͋t�ɁA���ɕ����邩���߂Ă����ɋϓ��ɔz�u���܂��B��100�̃f�[�^��3�ŕ������ꍇArray(Array(33),Array(33 ���X�g��empty),Array(33���X�g��empty))
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

'Key�������ݒ肳�ꂽDictionary��Ԃ��܂��B�d���͍폜����܂��Bexists�Ƃ���ʂɂ���ꍇ�֗�����
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

'���X�g���m�Řa�W��(�������ă��j�[�N��)���쐬���ĐV�������X�g�ŕԂ��܂�
Public Function UnionToList(argsList As list) As list

    Dim dic As Object: Set dic = Me.ToHashset
    Dim newlist As list: Set newlist = Me.Clone
   
    For Each elm In argsList.ToArray
        If Not dic.Exists(elm) Then newlist.Add (elm)
    Next
   
    Set UnionToList = newlist
   
End Function

'���X�g���m�ō��W��(���Ƃ��Ƃ��������̂�������ɂ�����̂�������������)���쐬���ĐV�������X�g�ŕԂ��܂��B
Public Function ExceptToList(argsList As list) As list

    Dim dic As Object: Set dic = argsList.ToHashset
    Dim list As list: Set list = New list
    Dim newLis As list: Set newlist = New list
   
    For Each elm In Me.ToArray
        If Not dic.Exists(elm) Then newlist.Add (elm)
    Next
   
    Set ExceptToList = newlist
   
End Function

'���X�g���m�ŐϏW��(�����ɂ������)���쐬���ĐV�������X�g�ŕԂ��܂��B
Public Function IntersectToList(argsList As list) As list

    Dim dic As Object: Set dic = argsList.ToHashset
    Dim list As list: Set list = New list
   
    For Each elm In Me.ToArray
        If dic.Exists(elm) Then list.Add (elm)
    Next
   
    Set IntersectToList = list
   
End Function

'�I�u�W�F�N�g�𕡐�(�V�����[�R�s�[)���܂��B
Public Function Clone() As list

    Dim newlist As list: Set newlist = New list
    Set Clone = newlist.Concat(Me)

End Function

'�w�萔�擪���甲���o��
Public Function TakeToList(itemCount As Long) As list

    '�w�萔���v�f���𒴂��Ă���ꍇ�N���[����Ԃ�
    Dim minindex As Long: minindex = LBound(myArr)
    If itemCount + 1 - minindex > myCount Then
        Set TakeToList = Me.Clone()
    Else
        Dim newlist As list: Set newlist = New list
        For i = minindex To itemCount - 1 + minindex
            newlist.Add (myArr(i))
        Next i
        Set TakeToList = newlist
    End If

End Function

'�w�萔���X�L�b�v���Ă���ȍ~�𔲂��o��
Public Function SkipToList(skipCount As Long) As list

    '�w�萔���v�f���𒴂��Ă���ꍇnull�̃A�C�e����Ԃ�
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

'�����Ń\�[�g���܂��B(�����̓N�C�b�N�\�[�g�ɂȂ��Ă���͂��ł�)
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

'�~���Ń\�[�g���܂��B
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

'�i�[����Ă���l�𔽓]���܂��B
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

' �l���V���b�t�����܂��B���x�͂��܂�����������܂���B���x���~������΃����Z���k�c�C�X�^�[�Ƃ��ǂ��炵���̂Ŏ������ĉ�����(�ʌ����dll���ق��������I?)
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

'���l�𕶎���ɕϊ����A���̑Ώۂ����X�g�Ɋ܂܂�邩�������܂��B
Public Function StringContains(str As String) As Boolean

    For Each elm In myArr
        If str = CStr(elm) Then
            StringContains = True
            Exit Function
        End If
    Next

    StringContains = False

End Function

'StringBuilder�̂悤�ɃR���N�V�����������Ɍ�������String��Ԃ��܂��B
'�������w�肵���ꍇ�A�R���N�V�����̋�؂�ɂ��̕�����t�^���܂��B
'�����ł���String��Immutable�ȎQ�ƌ^�ł��������͒l�^�Ɠ����悤�ɃR�s�[������邽�߁A���\�b�h�𕪂���̂̓I�[�o�[�w�b�h���N�\�ł����̂ł�
Public Function ToBuildString(Optional delimiter As String = "") As String

    Dim strCount&: strCount = 1
    Dim result$
    Dim chrNumber&
    Dim i&
    Dim elmLength&

    If delimiter = "" Then '�������Ȃ��ꍇ
   
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
       
    Else '����������ꍇ
   
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

'CSV(Character Separated Value )���쐬���܂��BStringBuilder�̂悤�ɍ����Ɍ������܂��B
'�Z���̒l���ꊇ�擾���ĊO���Ƀf�[�^�o�͂���ꍇ�ɂ߂ċ��͂ɓ����Ǝv���܂��B

'�Z���͈͂��񂽂�ɏ������݃��\�b�h��(�ȉ��̃��\�b�h����Value���K�v�ł�)
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

'��ӂ̒l(���j�[�N�l)�Ƀt�B���^�[���ĐV�������X�g�Ƃ��ĕԂ��܂��B
Public Function DistinctToList() As list

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim newlist As list: Set newlist = New list

    For Each elm In Me.ToArray
        dic(elm) = 0
    Next
   
    Call newlist.AddRange(dic.Keys)
    Set DistinctToList = newlist

End Function

'���l�𕶎���ɕϊ����A�Ώە����񂪃��X�g�Ɋ܂܂�邩���K�\���Ō������܂��B
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

'�ˉe�������s���܂��BExcel�֐����g�p�ł��܂��B�����͂Ȃ��ł��B
'�g�����̓��\�b�h�e�X�g�����Ă��������B
Public Function MAP(variableName As String, evaluateFunction As String) As list

    Dim newlist As list: Set newlist = New list
    Dim func$
   
    For Each elm In Me.ToArray
        func = Replace(evaluateFunction, variableName, elm)
        newlist.Add (Evaluate(func))
    Next
   
    Set MAP = newlist
   
End Function

'�t�B���^�[�������s���܂��BExcel�֐����g�p�ł��܂��B�����͂Ȃ��ł��B
'�g�����̓��\�b�h�e�X�g�����Ă�������
Public Function Filter(variableName As String, evaluateFunction As String) As list

    Dim newlist As list: Set newlist = New list
    Dim func$
   
    For Each elm In Me.ToArray
        func = Replace(evaluateFunction, variableName, elm)
        If Evaluate(func) Then newlist.Add (elm)
    Next
   
    Set Filter = newlist
   
End Function

'�������牺�͐��������ł��B
'�g�ݍ��݊֐��ōs����͈͂̏ꍇ�͂�����g�������ɏ������܂��B
'2�����z��ɍĔz�u����΁A�g�ݍ��݊֐����K�p�ł���͈͂������܂����A�x���Ȃ�ꍇ�͂�����s���܂���B

'���v�l���擾���܂��B
Public Function Math_Sum()

    If myCount < 65535 Then
        Math_Sum = WorksheetFunction.Sum(myArr)
    Else
        Dim buf
        For Each elm In myArr 'empty�̉��Z�̓G���[�ɂȂ�Ȃ����͂₢
            buf = buf + elm
        Next
        Math_Sum = buf
    End If
   
End Function

'���ϒl���擾���܂��B
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


'�����l���擾���܂��B
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

'�ő�l��Ԃ��܂��B
Public Function Math_Max()

    If myCount < 65535 Then
        Math_Max = WorksheetFunction.max(myArr)
    ElseIf myCount < 1048575 Then '�ӊO�Ɲh�R���Ă���̂ŊO���Ă��ǂ�
        Math_Max = WorksheetFunction.max(Me.ToTransposeArray)
    Else

    Dim max
    For Each elm In Me.ToArray
           max = IIf(elm > max, elm, max)
    Next
    Math_Max = max
    End If

End Function

'�ŏ��l��Ԃ��܂��B
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

'��W�c�̕W���΍������߂܂�
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


'�ŕp�l��1�Ƃ��Ă��܂��B
Public Function Math_ModeSingle()

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim cnt&: cnt = 1
    '���l,���̎��������
    For Each elm In Me.arr
        If Not (IsEmpty(elm)) Then
            dic(elm) = dic(elm) + 1
        End If
        cnt = cnt + 1
        If cnt >= 2500 Then '�����ɓ���鐔�������Ȃ�Ƃ��̂������x���Ȃ� ����O(N^2)
            Math_ModeSingle = ModeSinge_LowOverlap
            Exit Function
        End If
    Next

    If dic.Count = 0 Then Exit Function

    '�ő�̌��𒲂ׂ�
    Dim maxCounter As Long
    Dim val As Long
    For Each elm In dic.Items
        maxCounter = IIf(maxCounter < elm, elm, maxCounter)
    Next elm

    '�ő�l�ƈ�v����l��Ԃ�
    For Each elm In dic.Keys
        If dic(elm) = maxCounter Then
            Math_ModeSingle = elm
            Exit Function
        End If
    Next

End Function


'���܂�ɂ��d�������Ȃ��ꍇ
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


'�ŕp�l��z��ŕԂ��܂��B
Public Function Math_ModeMulti()

    Dim dic As Object: Set dic = CreateObject("scripting.dictionary")
    Dim cnt&: cnt = 1
    '���l,���̎��������
    For Each elm In Me.arr
        If Not (IsEmpty(elm)) Then
            dic(elm) = dic(elm) + 1
        End If
        cnt = cnt + 1
        If cnt >= 2500 Then '�����ɓ���鐔�������Ȃ�Ƃ��̂������x���Ȃ� ����O(N^2)
            Math_ModeMulti = ModeMulti_LowOverlap
            Exit Function
        End If
    Next

    If dic.Count = 0 Then Exit Function

    '�ő�̌��𒲂ׂ�
    Dim maxCounter As Long
    Dim val As Long
    For Each elm In dic.Items
        maxCounter = IIf(maxCounter < elm, elm, maxCounter)
    Next elm

    '�ő�l�ƈ�v����l������
    Dim newlist As list: Set newlist = New list
    For Each elm In dic.Keys
        If dic(elm) = maxCounter Then
            Call newlist.Add(elm)
        End If
    Next

    Math_ModeMulti = newlist.ToArray

End Function

'���܂�ɂ��d�������Ȃ��ꍇ
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


