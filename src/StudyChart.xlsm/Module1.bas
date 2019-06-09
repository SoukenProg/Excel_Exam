Attribute VB_Name = "Module1"
Dim Color(5) As Long
Dim Mark(5) As String
Dim Begin(3) As Integer

Function DefineData()
Color(0) = RGB(255, 188, 112)
Color(1) = RGB(255, 217, 112)
Color(2) = RGB(112, 255, 214)
Color(3) = RGB(126, 255, 112)
Color(4) = RGB(126, 112, 255)
Mark(0) = "èo"
Mark(1) = "í˜"
Mark(2) = "éé"
Mark(3) = "çá"
Mark(4) = "éË"
Begin(0) = 8
Begin(1) = 39
Begin(2) = 67
End Function

Sub addUniversity()
Attribute addUniversity.VB_Description = "ì¸ééèÓïÒÇÃçXêV"
Attribute addUniversity.VB_ProcData.VB_Invoke_Func = "J\n14"
Dim Schedule(5) As Date
Call DefineData

For i = 3 To 26 Step 1
    For j = 4 To 8 Step 1
        If IsEmpty(Cells(i, 4)) = False Then
            Schedule(j - 4) = Cells(i, j)
        End If
        
    Next j

    For j = 0 To 4 Step 1
         Dim m As Integer
         Dim d As Integer
        m = Month(Schedule(j))
        d = Day(Schedule(j))
        
        If m <= 3 And m >= 1 Then
        Dim r As Integer
        r = Begin(m - 1) + j
        
        Cells(i, r).Interior.ColorIndex = Color(j)
        Cells(i, r).Text = Mark(j)
        End If
        
    Next j
Next i

End Sub

