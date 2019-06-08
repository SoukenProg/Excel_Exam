Attribute VB_Name = "Module1"
Sub addUniversity()
Dim Schedule(5) As Date
Dim Mark(5) As String
Dim Color(5)


For i = 3 To 26 Step 1
    For j = 4 To 8 Step 1
        If IsEmpty(Cells(i, 4)) = False Then
            Schedule(j - 3) = Cells(i, j)
        End If
        
    Next j
Next i

For i = 1 To 5 Step 1
    If Month(Schedule(i)) <= 3 Then
    End If
Next i

End Sub
