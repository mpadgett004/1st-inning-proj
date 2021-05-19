Sub AmericanLeagueLongWay()

Dim LastRow As Long
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
Dim i As Integer

Dim orioles_a_over As Integer
Dim orioles_a_under As Integer
Dim orioles_h_over As Integer
Dim orioles_h_under As Integer
Dim orioles_over As Integer
Dim orioles_under As Integer

Dim red_sox_a_over As Integer
Dim red_sox_a_under As Integer
Dim red_sox_h_over As Integer
Dim red_sox_h_under As Integer
Dim red_sox_over As Integer
Dim red_sox_under As Integer

For i = 2 To LastRow
    If Cells(i, 2).Value = "Baltimore Orioles" Then
        If Cells(i, 3).Value + Cells(i, 5).Value >= 1 Then
            orioles_over = orioles_over + 1
            orioles_a_over = orioles_a_over + 1
        ElseIf Cells(i, 3).Value + Cells(i, 5).Value = 0 Then
            orioles_under = orioles_under + 1
            orioles_a_under = orioles_a_under + 1
        End If
    ElseIf Cells(i, 4).Value = "Baltimore Orioles" Then
        If Cells(i, 3).Value + Cells(i, 5).Value >= 1 Then
            orioles_over = orioles_over + 1
            orioles_h_over = orioles_h_over + 1
        ElseIf Cells(i, 3).Value + Cells(i, 5).Value = 0 Then
            orioles_under = orioles_under + 1
            orioles_h_under = orioles_h_under + 1
        End If
    End If
Next i

For i = 2 To LastRow
    If Cells(i, 2).Value = "Boston Red Sox" Then
        If Cells(i, 3).Value + Cells(i, 5).Value >= 1 Then
            red_sox_over = red_sox_over + 1
            red_sox_a_over = red_sox_a_over + 1
        ElseIf Cells(i, 3).Value + Cells(i, 5).Value = 0 Then
            red_sox_under = red_sox_under + 1
            red_sox_a_under = red_sox_a_under + 1
        End If
    ElseIf Cells(i, 4).Value = "Boston Red Sox" Then
        If Cells(i, 3).Value + Cells(i, 5).Value >= 1 Then
            red_sox_over = red_sox_over + 1
            red_sox_h_over = red_sox_h_over + 1
        ElseIf Cells(i, 3).Value + Cells(i, 5).Value = 0 Then
            red_sox_under = red_sox_under + 1
            red_sox_h_under = red_sox_h_under + 1
        End If
    End If
Next i

Cells(1, 8).Value = "Over/Under"
Cells(1, 9).Value = "Away O/U"
Cells(1, 10).Value = "Home O/U"



Dim orioles_a_record
Dim orioles_h_record
Dim orioles_record
orioles_a_record = CStr(orioles_a_over) + "--" + CStr(orioles_a_under)
orioles_h_record = CStr(orioles_h_over) + "--" + CStr(orioles_h_under)
orioles_record = CStr(orioles_over) + "--" + CStr(orioles_under)
Cells(2, 7).Value = "Baltimore Orioles"
Cells(2, 8).Value = orioles_record
Cells(2, 9).Value = orioles_a_record
Cells(2, 10).Value = orioles_h_record

Dim red_sox_a_record
Dim red_sox_h_record
Dim red_sox_record
red_sox_a_record = CStr(red_sox_a_over) + "--" + CStr(red_sox_a_under)
red_sox_h_record = CStr(red_sox_h_over) + "--" + CStr(red_sox_h_under)
red_sox_record = CStr(red_sox_over) + "--" + CStr(red_sox_under)
Cells(3, 7).Value = "Boston Red Sox"
Cells(3, 8).Value = red_sox_record
Cells(3, 9).Value = red_sox_a_record
Cells(3, 10).Value = red_sox_h_record



End Sub