### This VBA routine taken from 
#### http://www.ozgrid.com/News/excel-histogram.htm ==> had to modify because this made certain assumptions about the array
### allows user to calculate histogram in VBA code

``` VBA

Sub Hist(m As Long, rng As Range)

    ' sample of how to load array from range
 
    Dim Arr As Variant
     
    Arr = rng.Value
    maxValue = WorksheetFunction.Max(Arr)
     'pass range values to array
     ' MsgBox (maxValue)

    Dim i As Long, j As Long
    Dim Length As Single
    ReDim breaks(m) As Single
    ReDim freq(m) As Single
    
    'Assign initial value for the frequency array
    For i = 1 To m
        freq(i) = 0
    Next i

    'Linear interpolation
    Length = maxValue / m
    For i = 1 To m
        breaks(i) = Length * i
    Next i
    
    'Counting the number of occurrences for each of the bins
    For i = 1 To UBound(Arr)
        If (Arr(i, 1) <= breaks(1)) Then freq(1) = freq(1) + 1
        If (Arr(i, 1) >= breaks(m - 1)) Then freq(m) = freq(m) + 1
        For j = 2 To m - 1
            If (Arr(i, 1) > breaks(j - 1) And Arr(i, 1) <= breaks(j)) Then freq(j) = freq(j) + 1
        Next j
    Next i
    
    'Sheets("Summary").Select
    'Display the frequency distribution on the summary worksheet
    'For i = 1 To m
    '   Cells(i + 40, 2) = breaks(i)
    '   Cells(i + 40, 3) = freq(i)
    'Next i
    'Sheets("Transformed Data").Select
    
End Sub


```
