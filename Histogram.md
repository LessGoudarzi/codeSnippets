### This VBA routine taken from http://www.ozgrid.com/News/excel-histogram.htm
### allows user to calculate histogram in VBA code

``` VBA

Sub Hist(M As Long, arr() As Single)

   ' sample of how to load array from range
    'Dim vArr As Variant
    'Dim l As Long, m As Long
     
    'vArr = Range("A10:C600").Value
    'pass range values to array


    Dim i As Long, j As Long
    Dim Length As Single
    ReDim breaks(M) As Single
    ReDim freq(M) As Single
    
    'Assign initial value for the frequency array
    For i = 1 To M
        freq(i) = 0
    Next i

    'Linear interpolation
    Length = (arr(UBound(arr)) - arr(1)) / M
    For i = 1 To M
        breaks(i) = arr(1) + Length * i
    Next i
    
    'Counting the number of occurrences for each of the bins
    For i = 1 To UBound(arr)
        If (arr(i) <= breaks(1)) Then freq(1) = freq(1) + 1
        If (arr(i) >= breaks(M - 1)) Then freq(M) = freq(M) + 1
        For j = 2 To M - 1
            If (arr(i) > breaks(j - 1) And arr(i) <= breaks(j)) Then freq(j) = freq(j) + 1
        Next j
    Next i
    
    'Display the frequency distribution on the active worksheet
    For i = 1 To M
        Cells(i, 1) = breaks(i)
        Cells(i, 2) = freq(i)
    Next i
End Sub


```
