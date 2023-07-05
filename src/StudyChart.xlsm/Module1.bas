Attribute VB_Name = "Module1"
Dim Color(5) As Long
Dim Mark(5) As String
Dim Begin(3) As Integer

Function DefineData()
'データの定義
Color(0) = RGB(255, 188, 112)
Color(1) = RGB(255, 217, 112)
Color(2) = RGB(112, 255, 214)
Color(3) = RGB(126, 255, 112)
Color(4) = RGB(126, 112, 255)
Mark(0) = "出"
Mark(1) = "締"
Mark(2) = "試"
Mark(3) = "合"
Mark(4) = "手"
Begin(0) = 20
Begin(1) = 51
Begin(2) = 80
End Function

Sub addUniversity()
Attribute addUniversity.VB_ProcData.VB_Invoke_Func = "J\n14"


Columns("BS").Hidden = True


'書き換えのためセルをクリア
Range("T4:DG27").Clear

Dim Schedule(5) As Date
Call DefineData

'うるう年の対応
Dim Y As Long
Y = Range("E29").Value
If Y Mod 4 = 0 Then
    Columns("BS").Hidden = False
Else
    If Y Mod 100 = 0 And Y Mod 400 <> 0 Then
        Columns("BS").Hidden = True
    End If
 End If

For i = 4 To 27 Step 1
       For j = 6 To 10 Step 1
    ' 値があるなら日付を取得
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
        
        '有効範囲内なら色付け
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
        
    '値を有効範囲外に
    For j = 0 To 4 Step 1
    Schedule(j) = #5/21/2020#
    Next j
    
    '試験日コードの色付け
    Dim Code As Integer
    
    For x = 0 To 5 Step 1
    
    If (Cells(i, 11)) = x And (Cells(i, 11)) <> 0 Then
    
    For Y = 0 To 5 Step 1
    If IsEmpty(Cells(x + 30, Y + 5)) = False Then
    
      m = Month(Cells(x + 30, Y + 5))
         d = Day(Cells(x + 30, Y + 5))

        
    End If
    r = Begin(m - 1) + d
    
        Cells(i, r).Interior.Color = Color(2)
        Cells(i, r).Value = Mark(2)
    
    Next Y
    
    End If
    
    Next x
    
Next i

'罫線を描画
Range("T4:DG27").Borders.LineStyle = True


End Sub


