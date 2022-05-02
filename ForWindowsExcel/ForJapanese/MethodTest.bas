Attribute VB_Name = "MethodTest"

'Ctrl+G�Ŋe���\�b�h���ǂ̂悤�ɓ��삵�Ă��邩�m�F���邱�Ƃ��ł��܂��B

Private List As List, List1 As List, List2 As List, List3 As List

Public Sub AllTestExecute()
    Call AllTest
End Sub


'�����������Ƃ��ł���Ƃ����f���ł��B�������ۂɂ����������Ƃ����Ȃ炢�������Ƃɐ����ϐ��ɑ������ׂ����Ǝv���܂��B
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
   
   '���\�b�h�`�F�[�����̂ǂ��łł� .DebugPrint �����ē��e���m�F���邱�Ƃ��\�ł��B
   
    '����Ă邱��
    'List1��2�ɘA������l��ǉ�
    '��List1��(�J�n��,����鐔)�`���ŁAList2�͍쐬�����A�Ԃ�n�A������i,j,k�Ƃ���� n = i to j (step k)�`���ŘA�Ԃ��쐬���܂��B
    'List1��2�̐ϏW��(�����ɂ���l�̂ݎc��)���쐬���ĕʂ̃��X�g���쐬(ListX�Ƃ���)
    'ListX���~���\�[�g
    'ListX��͈͎w��ŃX���C�X���ĕʂ̃��X�g��Ԃ�(ListY�Ƃ���)
    'ListY�Ɏˉe�������s��(Evaluate���g�p,�ˉe��̃��X�g��ListZ�Ƃ���) �����̏ꍇ�΂�2���|����10�̔{���ɐ؂艺���Ă���
    'ListZ�Ƀt�B���^�[�������s���ʂ̃��X�g(ListA)���̏ꍇ20�̔{���݂̂ɂ��� ������MAP�Ƃقړ���
    'ListA�̏d�����폜(ListB�Ƃ���)
    'ListB����Z�p���[�^��tab�ŉ��s�R�[�h��Cr�ȕ�������쐬(�����Ȃ��̏ꍇ�J���}��CrLf�A������StringBuilder�Ȃ̂ł͂₢!)
   
End Sub

'���X�g�̒��g��\�����܂��B "���\�b�h�`�F�[���̎��̓r���ɂ����ނ��Ƃ��\�ł�!!"
Private Sub DebugPrintTest()
    Set List = New List
    Call List.CreateEnumRange(500000000, 3) '5������n�܂�l3���擾
    Call List.DebugPrint("Debug.Print =>  ", "#,##0", "�~�~����!!")
End Sub

'�f�[�^��ǉ����܂��B
Private Sub AddTest()

    Set List = New List
    For i = 0 To 3
        List.Add (i)
    Next i
   
    Call List.DebugPrint("Add => ")

End Sub

'���X�g�����������܂�
Private Sub ClearTest()

    Set List = New List
    Call List.CreateEnumRange(1, 5)
    Call List.DebugPrint("ClearTest1=> ")
    Call List.Clear.DebugPrint("ClearTest2=> ") '���g�������̂ŕ\������܂���

End Sub

'�f�[�^�𕡐��ǉ����܂��B�z�񂪑Ώۂł��B �Ώۂ�1�����̏ꍇ�قڗ��_�l���o��Ǝv���܂��B
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

'2�̃��X�g���������܂��B�����ɓ��ꂽ�ق������Ƃɗ��܂�
Private Sub ConcatTest()

    Set List1 = New List
    Set List2 = New List
    
    Call List1.CreateEnumRange(start:=5, Count:=5) '�J�n�ʒu����1��������`�œ���̌��̒l�����܂��B
    Call List2.CreateSeqNumbers(First:=5, Last:=100, step:=20) 'for i �݂����Ȍ`���Œl�����܂��B
    Call List1.Concat(List2).DebugPrint("Concat=> ", "#,##0")
    
End Sub

'�ŏ��̒l���擾
Private Sub FirstTest()
    Set List = New List
    Debug.Print "First=> " & List.CreateEnumRange(100, 5).First
End Sub

'�Ō�̒l���擾
Private Sub LastTest()
    Set List = New List
    Debug.Print "First=> " & List.CreateEnumRange(100, 5).Last
End Sub

'�v�f�����邩
Private Sub AnyTest()
    Set List = New List
    Debug.Print "Any1=> " & List.Any_
    List.Add (1)
    Debug.Print "Any2=> " & List.Any_
End Sub

'�v�f��������
Private Sub NothingTest()
    Set List = New List
    Debug.Print "Nothing1=> " & List.Nothing_
    List.Add (1)
    Debug.Print "Nothing2=> " & List.Nothing_
End Sub

'���ׂĂ̒l����v���Ă��邩
Private Sub SequenceEqualTest()
    Set List1 = New List
    Set List2 = New List
    
    Call List1.CreateEnumRange(5, 5)
    Call List2.CreateEnumRange(5, 5)
    Debug.Print "SequenceEqual1=> " & List1.SequenceEqual(List2)
    Debug.Print "SequenceEqual2=> " & List1.SequenceEqual(List2.Clear)
End Sub

'���g���N���A���ĘA�Ԃ��쐬���܂� for���̏������Ɠ����ł��B
Private Sub CreateSeqNumbersTest()
    Set List = New List
    Call List.CreateSeqNumbers(First:=100, Last:=500, step:=80).DebugPrint("CreateSeqNumbers=> ")
End Sub

'���g���N���A����.Net��Enumlabre.Range�̂悤�Ȋ����� �J�n�ʒu����1��������A�Ԃ��w������܂��B
Private Sub CreateEnumRangeTest()
    Set List = New List
    Call List.CreateEnumRange(start:=100, Count:=5).DebugPrint("CreateEnumRange=> ")
End Sub

'�C���f�b�N�X�l�ɑΉ�����l���擾���܂��B�z���[n]�Ɠ����ł��B
Private Sub GetValueOfIndexTest()
    Set List = New List
    Call List.CreateEnumRange(100, 5)
    Debug.Print "GetValueOfIndexTest=> " & List.GetValueOfIndex(3)
End Sub

'����͈̔͂𔲂��o���ĐV�������X�g�Ƃ��ĕԂ��܂��B(minIndex<=�����o������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
Private Sub SliceTest()
    Set List = New List
    Call List.CreateEnumRange(100, 20)
    Call List.Slice(minIndex:=-50, maxIndex:=3).DebugPrint("SliceTest=> ")
End Sub

'index�ɑΉ�����v�f���폜���ăf�[�^��O�ɋl�߂܂��B�����͈����ł��B
Private Sub RemoveTest()
    Set List = New List
    Call List.CreateEnumRange(1, 5)
    Call List.Remove(2).DebugPrint("Remove => ")
End Sub

'�����̐��l�Ɉ�v������̂�S�č폜���ăf�[�^��O�ɋl�߂܂��B������(��
Private Sub RemoveAllTest()
    Set List = New List
    Call List.AddRange(Array(5, 3, 2, 5, 3, 2, 3, 4, 3))
    Call List.RemoveAll(3).DebugPrint("RemoveAll => ")
End Sub

'����͈̔͂��폜���ċl�߂܂��B(minIndex<=����������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
Private Sub RemoveRangeTest()
    Set List = New List
    Call List.AddRange(Array(5, 3, 2, 5, 3, 2, 3, 4, 3))
    Call List.RemoveRange(minIndex:=4, maxIndex:=15).DebugPrint("RemoveRange => ")
End Sub

'����͈̔͂��폜���ċl�߂܂��B(minIndex<=����������<=maxIndex)
'Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
Private Sub PopTest()
    Set List = New List
    Call List.CreateEnumRange(100, 5)
    Debug.Print "Pop(�o�͒l)=> "; List.Pop(3)
    Call List.DebugPrint("pop(�c�����l)=> ")
End Sub

''����͈̔͂�Ԃ��Ɠ����ɍ폜���ċl�߂܂��B(minIndex<=����������<=maxIndex)
''Min���s���ɏ������AMax���s���ɑ傫���ꍇ�A�ŏ��C���f�b�N�X�A�ő�C���f�b�N�X�܂ł̗v�f��ΏۂƂ��܂��B
''�Ԃ�l�͎擾�����l�ł��B
Private Sub PopRangeTest()
    Set List = New List
    Call List.CreateEnumRange(100, 6)
    Call List.PopRange(minIndex:=2, maxIndex:=4).DebugPrint("PopRange(�擾�l)=> ")
    Call List.DebugPrint("PopRange(�c�����l)=> ")
End Sub

'�z��ɕϊ����ĕԂ��܂��B �z��̗v�f���̓f�[�^���ɐ؂�l�߂��܂��B
Private Sub ToArrayTest()
    Set List = New List
    For Each elm In List.CreateEnumRange(1, 5).ToArray
        Debug.Print "Toarray=> " & elm
    Next
End Sub

'(0,n)�̔z��ɕϊ����ĕԂ��܂��B �z��̗v�f���̓f�[�^���ɐ؂�l�߂��܂��B �c�����̒l�̓\�t�ɂƂĂ��g���邩�ƁB
Private Sub ToTransposeArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 5).ToTransposeArray
    Rem Range("A1:A5").Value = buffer  �Ⴆ�΂��������g�������ł��܂��B
    For Each elm In buffer
        Debug.Print "ToTransposeArray=> " & elm; "���v�E�H�b�`�E�B���h�E"
    Next
   
End Sub

'2�����z��ɕϊ����ĕԂ��܂��B 2�����ڂ̌���ݒ�ł��܂��B(�񐔂��w�肵�ĕ����ł���`�ł�)
Private Sub To2DArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 10).To2DArray(columnCount:=5)
    For Each elm In buffer
        Debug.Print "To2DArray=> " & elm; "���v�E�H�b�`�E�B���h�E"
    Next
   
End Sub

'2�����z��ɕϊ����ĕԂ��܂��B 2�����ڂ̌���ݒ�ł��܂��B(�񐔂��w�肵�ĕ����ł���`�ł�)
Private Sub To2DTransposeArrayTest()
   
    Set List = New List
    Dim buffer(): buffer = List.CreateEnumRange(1, 10).To2DTransposeArray(rowCount:=5)
    For Each elm In buffer
        Debug.Print "To2DTransposeArray=> " & elm; "���v�E�H�b�`�E�B���h�E"
    Next
   
End Sub

'chunkSize�̌����Ƃɕ������z���z��̒��ɓ���ĕԂ��܂��B ��100�̃f�[�^��40�ŕ������ꍇArray(Array(40),Array(40),Array(40(�㔼20��empty))) �[����empty�ɂȂ�d�l�ł��B�B
Private Sub ToChunkArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToChunkArray(6)
    Debug.Print "Chunk=> "
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab) '�^�u��؂��1�s�ɂ���
Next
   
End Sub

'��̂ɉ�����Chunk���̔z���Transpose����2�����z��ɂ��Ă܂�(For(Each)���񂵂ē\�t������ۂɎg���₷�����Ǝv���܂��B)
Private Sub ToTransposeChunkArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToTransposeChunkArray(6)
    Debug.Print "ToTransposeChunkArray=> ���v�E�H�b�`�E�B���h�E"
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab)
Next
   
End Sub

'chunk�Ƃ͋t�ɁA���ɕ����邩���߂Ă����ɋϓ��ɔz�u���܂��B��100�̃f�[�^��3�ŕ������ꍇArray(Array(33),Array(33 ���X�g��empty),Array(33���X�g��empty))'�[����empty�ɂȂ�d�l�ł��B�B
Private Sub ToDivideArrayTest()
   
    Set List = New List
    Dim chunk(): chunk = List.CreateSeqNumbers(160, 240, 5).ToDivideArray(6)
    Debug.Print "ToDivideArray=> "
   
For Each arr In chunk
    Debug.Print _
        List.Clear.AddRange(arr).ToBuildString(vbTab)
Next
   
End Sub

'Key�������ݒ肳�ꂽDictionary��Ԃ��܂��B�d���͍폜����܂��Bexists�Ƃ���ʂɂ���ꍇ�֗����ƁBWindows�ȊO�ł͑��������܂���B
Private Sub ToHashsetTest()
   
    Set List = New List
    Dim dic: Set dic = List.CreateSeqNumbers(100, 150, 10).ToHashset

    For Each Key In dic.Keys
        Debug.Print "ToHashset=> " & Key
    Next
   
End Sub

'���X�g���m�Řa�W��(�������ă��j�[�N��)���쐬���ĐV�������X�g�ŕԂ��܂�
'���Ƃ̃��X�g�̏d���l�������Ă��܂������Ȃ̂ł��������Ă܂����E�E�E?
Private Sub UnionToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 60, 10)
    Call List2.CreateSeqNumbers(0, 60, 12)
    Call List1.UnionToList(List2).DebugPrint("UnionToList=> ")
   
End Sub

'���X�g���m�ō��W��(���Ƃ��Ƃ��������̂�������ɂ�����̂�������������)���쐬���ĐV�������X�g�ŕԂ��܂��B
Private Sub ExceptToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 15, 2)
    Call List2.CreateSeqNumbers(0, 15, 3)
    Call List1.ExceptToList(List2).DebugPrint("ExceptToList=> ")
   
End Sub

'���X�g���m�ŐϏW��(�����ɂ������)���쐬���ĐV�������X�g�ŕԂ��܂�
Private Sub IntersectToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 24, 2)
    Call List2.CreateSeqNumbers(0, 24, 3)
    Call List1.IntersectToList(List2).DebugPrint("IntersectToList=> ")
   
End Sub

'�I�u�W�F�N�g�𕡐�(�V�����[�R�s�[)���܂��B
Private Sub CloneTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 20)
    Call List.Clone.DebugPrint("Clone=> ")
End Sub

'�w�萔�擪���甲���o���܂��B
Private Sub TakeToListTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 10)
    Call List.TakeToList(5).DebugPrint("TakeToList=> ")
End Sub

'�w�萔���X�L�b�v���Ă���ȍ~�𔲂��o���܂�
Private Sub SkipToListTest()
    Set List = New List
    Call List.CreateSeqNumbers(0, 100, 10)
    Call List.SkipToList(5).DebugPrint("SkipToList=> ")
End Sub

' �l���V���b�t�����܂��B���x�͂��܂�����������܂���B���x���~������΃����Z���k�c�C�X�^�[�Ƃ��ǂ��炵���̂Ŏ������ĉ�����()���Ԃ�̂��l��dll�Ƃ��ɂ��Ă܂��B
Private Sub RandamizeTest()
    Set List = New List
    Call List.CreateSeqNumbers(5, 10)
    Call List.Randamize.DebugPrint("Rndamize=> ")
End Sub

'�����Ń\�[�g���܂��B(�����̓N�C�b�N�\�[�g�ɂȂ��Ă���͂��ł�) �ڋʋ@�\��1�ł��B
Private Sub SortTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Randamize.Sort.DebugPrint("Sort=> ")
End Sub

'�~���Ń\�[�g���܂��B
Private Sub SortByDescendingTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Randamize.SortByDescending.DebugPrint("SortByDescending=> ")
End Sub

'�i�[����Ă���l�𔽓]���܂��B
Private Sub ReverseTest()
    Set List = New List
    Call List.CreateSeqNumbers(1, 5)
    Call List.Reverse.DebugPrint("Reverse=> ")
End Sub

'�l�𕶎���Ƃ��Ă݂Ȃ��āA���̑Ώۂ����X�g�Ɋ܂܂�邩�������܂�
Private Sub StringContainsTest()
    Set List = New List
    Call List.AddRange(Array(1, 2, 3, 4, 5, 487, "��������"))
    Debug.Print "StringContains1 =>  " & List.StringContains("1")
    Debug.Print "StringContains2 =>  " & List.StringContains("7")
End Sub

'���l�𕶎���Ƃ݂Ȃ��āA�Ώە����񂪃��X�g�Ɋ܂܂�邩���K�\���Ō������܂��BWindows�ȊO���Ɠ����Ȃ��Ǝv���܂��B
Private Sub StringContains_RegExpTest()
    Set List = New List
    Call List.AddRange(Array(1, 2, 3, 4, 5, 487, "��������"))
    Debug.Print "StringContains_RegExp1 =>  " & List.StringContains_RegExp(".*7")
    Debug.Print "StringContains_RegExp2 =>  " & List.StringContains_RegExp("[0-9]{3}")
    Debug.Print "StringContains_RegExp3 =>  " & List.StringContains_RegExp("[��-꤂�-��@-��]{3,}")
    Debug.Print "StringContains_RegExp4 =>  " & List.StringContains_RegExp("��+")
    Debug.Print "StringContains_RegExp4 =>  " & List.StringContains_RegExp("^87")
End Sub

'StringBuilder�̂悤�ɃR���N�V������"���ɍ�����"��������String��Ԃ��܂��B(10������x��1000�{��������͂��ł�) �ڋʋ@�\��1�ł�
'�������w�肵���ꍇ�A�R���N�V�����̋�؂�ɂ��̕�����t�^���܂��B(CSV�݂����Ȃ��̂����܂����ACSV(TSV�Ȃǂ��܂�)����郁�\�b�h�͕ʂɂ���܂�)
Private Sub ToBuildStringTest()
    Set List = New List
    Call List.AddRange(Array("������", "������", "�܍�", "�̂��肫��", "�C����", "������", "���s��", "�_����", "������"))
    Debug.Print "ToBuildString1 =>  " & List.ToBuildString()
    Debug.Print "ToBuildString2 =>  " & List.ToBuildString(",")
End Sub

'CSV(Character Separated Value )���쐬���܂��BStringBuilder�̂悤�ɍ����Ɍ������܂��B
'AddRange�Ƒg�ݍ��킹��΃Z������l���擾���āACSV�ɕϊ�����̂�1�s�ŏ����܂��B
Private Sub ToBuildCSVTest()
    Set List = New List
    Call List.AddRange(Array("����", "�N��", "����", "������", 30, "�j", "������", 40, "��"))
    Debug.Print "ToBuildCSV =>"
    Debug.Print List.ToBuildCSV(3, ",", vbCrLf)
End Sub

'�d���̂Ȃ���(���j�[�N�l)�Ƀt�B���^�[���ĐV�������X�g�Ƃ��ĕԂ��܂��B �ڋʋ@�\��1�ł�
Private Sub DistinctToListTest()
   
    Set List1 = New List
    Set List2 = New List
   
    Call List1.CreateSeqNumbers(0, 12, 2)
    Call List2.CreateSeqNumbers(0, 12, 3)
    Call List1.Concat(List2).DistinctToList.Sort.DebugPrint("DistinctToList=> ")
   
End Sub

'�ˉe�������s���܂��B����������Evaluate�֐��ł��邽�߁AExcel�֐����g�p�ł��܂��B�����͂Ȃ��ł��B
Private Sub MapTest()

    Set List = New List
   
      '���̏ꍇ���ꂼ��̒l�Ƀ΂��|����10�̔{���ɐ؂艺���Ă���B
    Call List.CreateSeqNumbers(0, 5) _
        .DebugPrint("Before MAP=> ") _
        .MAP("x", "floor(x*PI()*2,10)") _
        .DebugPrint("After Map=> ")
   
End Sub

'�t�B���^�[�������s���܂��B����������Evaluate�֐��ł��邽�߁AExcel�֐����g�p�ł��܂��B�����͂Ȃ��ł��B
Private Sub FilterTest()

    Set List = New List
   
      '���̏ꍇ20�Ŋ��������܂肪0�ɂȂ���̂��c���܂��B
    Call List.CreateSeqNumbers(0, 50, 10) _
        .DebugPrint("Before filtering =>") _
        .Filter("x", "Mod(x,20)=0") _
        .DebugPrint("After filtering=> ")
   
End Sub

'���������ł��B
'�g�ݍ��݊֐��ōs����͈͂̏ꍇ�͂�����g�������ɏ������܂��B
'2�����z��ɍĔz�u����΁A�g�ݍ��݊֐����K�p�ł���͈͂������܂����A�x���Ȃ�ꍇ�͂�����s���܂���B
'�O��m��(WorksheetFunction�ł�1�����z��ɂ�65535���x�̐������������肷�邪�A2������1048576�������肻��ȏ�\�������肷��)
Private Sub MathematicalFunctionsTest()

    Set List1 = New List

    Call List1.CreateSeqNumbers(1, 2000000) '1���Ƃ����܂葝�₵������Ɨ����邩��
    Debug.Print "Sum =>  " & List1.Math_Sum '���v
    Debug.Print "Average =>  " & List1.Math_Average '����
    Debug.Print "Median =>  " & List1.Math_Median '�����l
    Debug.Print "Max =>  " & List1.Math_Max '�ő�l
    Debug.Print "Min =>  " & List1.Math_Min '�ŏ��l
    Debug.Print "StDevP =>  " & List1.Math_StDevP '�W���΍�
   
    '�ŕp�l�̎擾
    Set List2 = New List
    For i = 1 To 100000
        List2.Add (Int(Rnd() * 2147483647))
    Next
    
    For i = 1 To 65535
        For j = 1 To 10
            List2.Add (j)
        Next j
    Next i

    Debug.Print "ModeSingle =>  " & List2.Math_ModeSingle '�ŕp�l(��������ꍇ1�̂�)
   
    For Each buf In List2.Math_ModeMulti '�ŕp�l(�S��)
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



