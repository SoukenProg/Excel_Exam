Attribute VB_Name = "Module1"
Dim Color(5) As Long
Dim Mark(5) As String
Dim Begin(3) As Integer

Function DefineData()
'�f�[�^�̒�`
Color(0) = RGB(255, 188, 112)
Color(1) = RGB(255, 217, 112)
Color(2) = RGB(112, 255, 214)
Color(3) = RGB(126, 255, 112)
Color(4) = RGB(126, 112, 255)
Mark(0) = "�o"
Mark(1) = "��"
Mark(2) = "��"
Mark(3) = "��"
Mark(4) = "��"
Begin(0) = 10
Begin(1) = 41
Begin(2) = 69
End Function


Sub addUniversity()
Attribute addUniversity.VB_ProcData.VB_Invoke_Func = "J\n14"
'���������̂��߃Z�����N���A
Range("K4:CV26").Clear

Dim Schedule(5) As Date
Call DefineData

For i = 4 To 26 Step 1
    For j = 6 To 10 Step 1
    ' �l������Ȃ���t���擾
        If IsEmpty(Cells(i, j)) = False Then
            Schedule(j - 6) = Cells(i, j)
        End If
        
    Next j

    For j = 0 To 4 Step 1
         Dim m As Integer
         Dim d As Integer
        m = Month(Schedule(j))
        d = Day(Schedule(j))
        
        '�L���͈͓��Ȃ�F�t��
        If m <= 3 And m >= 1 Then
        Dim r As Integer
        r = Begin(m - 1) + d
        
        Cells(i, r).Interior.Color = Color(j)
        Cells(i, r).Value = Mark(j)
        End If
        
    Next j
    '�l��L���͈͊O��
    For j = 0 To 4 Step 1
    Schedule(j) = #12/31/2020#
    Next j
Next i

'�r����`��
Range("I4:CV26").Borders.LineStyle = True


End Sub



