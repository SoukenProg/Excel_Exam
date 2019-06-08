Attribute VB_Name = "Module1"
Sub addUniversity()
Dim Schedule(4) As Date

For i = 3 To 26 Step 1
    For j = 4 To 8 Step 1
        If IsEmpty(Cells(i, 4)) = False Then
            AppStart = Cells(i, 4)
            
            If Month(AppStart) <= 3 Then
            End If
      
        End If
        
    Next j
Next i

End Sub
