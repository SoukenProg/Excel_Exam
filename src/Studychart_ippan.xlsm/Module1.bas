Attribute VB_Name = "Module1"
Dim Color(5) As Long
Dim Mark(5) As String
Dim Begin(3) As Integer
Dim Code(5, 5) As Date

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
Begin(0) = 19
Begin(1) = 50
Begin(2) = 80
'�������R�[�h�p
For i = 0 To 5 Step 1
    For j = 0 To 5 Step 1
    Code(i, j) = Cells(i + 5, j + 29)
    Next j
    Next i
End Function



Sub addUniversity()
Attribute addUniversity.VB_ProcData.VB_Invoke_Func = "J\n14"
'���������̂��߃Z�����N���A
Range("L4:CV26").Clear

Dim Schedule(5) As Date
Call DefineData

For i = 4 To 26 Step 1
    For j = 6 To 10 Step 1
    ' �l������Ȃ���t���擾
        If IsEmpty(Cells(i, j)) = False Then
            Schedule(j - 6) = Cells(i, j)
            
        Else
        Schedule(j - 6) = #5/21/2019#
        
        
        End If
        
        Next j

    For k = 0 To 4 Step 1
    
         Dim m As Integer
         Dim d As Integer
        m = Month(Schedule(k))
        d = Day(Schedule(k))
        
        '�L���͈͓��Ȃ�F�t��
        If m <= 3 And m >= 1 Then
        Dim r As Integer
        
        r = Begin(m - 1) + d
        
        Cells(i, r).Interior.Color = Color(k)
        Cells(i, r).Value = Mark(k)
        
        ElseIf m = 12 Then
        Cells(i, 11).Interior.Color = Color(k)
        Cells(i, 11).Value = Mark(k)
        End If
        
        Next k
        
    '�l��L���͈͊O��
    For j = 0 To 4 Step 1
    Schedule(j) = #5/21/2020#
    Next j


    '�������R�[�h�̐F�t��
    Dim Code As Integer
    
    For x = 0 To 5 Step 1
    
    If (Cells(i, 11)) = x And (Cells(i, 11)) <> 0 Then
    
    For y = 0 To 5 Step 1
    If IsEmpty(Cells(x + 28, y + 5)) = False Then
    
      m = Month(Cells(x + 28, y + 5))
         d = Day(Cells(x + 28, y + 5))

        
    End If
    r = Begin(m - 1) + d
    
        Cells(i, r).Interior.Color = Color(2)
        Cells(i, r).Value = Mark(2)
    
    Next y
    
    End If
    
    Next x
    
    
    
Next i
'�r����`��
Range("L4:CV26").Borders.LineStyle = True


End Sub


