Sub FinalizedCode()
Dim LastRow As Long
LastRow = Cells(Rows.Count, "A").End(xlUp).Row
Dim z As Integer

Dim a_over As Integer
Dim a_under As Integer
Dim h_over As Integer
Dim h_under As Integer
Dim over As Integer
Dim under As Integer

'declare a variant array
Dim strNames() As String
'initialize the array
ReDim strNames(10 To 39)
'populate the array
strNames(10) = "Baltimore Orioles"
strNames(11) = "Boston Red Sox"
strNames(12) = "New York Yankees"
strNames(13) = "Toronto Blue Jays"
strNames(14) = "Tampa Bay Rays"
strNames(15) = "Chicago White Sox"
strNames(16) = "Cleveland Indians"
strNames(17) = "Detroit Tigers"
strNames(18) = "Kansas City Royals"
strNames(19) = "Minnesota Twins"
strNames(20) = "Los Angeles Angels"
strNames(21) = "Oakland Athletics"
strNames(22) = "Seattle Mariners"
strNames(23) = "Texas Rangers"
strNames(24) = "Houston Astros"
strNames(25) = "Atlanta Braves"
strNames(26) = "Washington Nationals"
strNames(27) = "New York Mets"
strNames(28) = "Philadelphia Phillies"
strNames(29) = "Miami Marlins"
strNames(30) = "Milwaukee Brewers"
strNames(31) = "Chicago Cubs"
strNames(32) = "Cincinnati Reds"
strNames(33) = "Pittsburgh Pirates"
strNames(34) = "St. Louis Cardinals"
strNames(35) = "Los Angeles Dodgers"
strNames(36) = "San Diego Padres"
strNames(37) = "San Francisco Giants"
strNames(38) = "Colorado Rockies"
strNames(39) = "Arizona Diamondbacks"

'declare an integer
Dim i As Integer
'loop from the lower bound of the array to the upper bound of the array - the entire array
'For i = LBound(strNames) To UBound(strNames)
For i = 10 To 39
    For z = 2 To LastRow
        If Cells(z, 2).Value = strNames(i) Then
            If Cells(z, 3).Value + Cells(z, 5).Value >= 1 Then
                over = over + 1
                a_over = a_over + 1
            ElseIf Cells(z, 3).Value + Cells(z, 5).Value = 0 Then
                under = under + 1
                a_under = a_under + 1
            End If
        ElseIf Cells(z, 4).Value = strNames(i) Then
            If Cells(z, 3).Value + Cells(z, 5).Value >= 1 Then
                over = over + 1
                h_over = h_over + 1
            ElseIf Cells(z, 3).Value + Cells(z, 5).Value = 0 Then
                under = under + 1
                h_under = h_under + 1
            End If
        End If
    Next z
'show the name in the immediate window
    Dim a_record
    Dim h_record
    Dim record
    Cells(i, 7).Value = strNames(i)
    a_record = CStr(a_over) + "--" + CStr(a_under)
    h_record = CStr(h_over) + "--" + CStr(h_under)
    record = CStr(over) + "--" + CStr(under)
    Cells(i, 8).Value = record
    Cells(i, 9).Value = a_record
    Cells(i, 10).Value = h_record
    
    a_over = 0
    a_under = 0
    h_over = 0
    h_under = 0
    over = 0
    under = 0
 
Next i
End Sub
