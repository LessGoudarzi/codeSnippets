Sort Routines
==============

Note, I don't have a traceback to where I found these.  If anyone recognizes them, please forward me the reference for adding an appropriate source reference.

```VBA

Sub BubbleSort(arr, lngMax)
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
 ' Dim lngMax As Long
  lngMin = LBound(arr)
  'lngMax = UBound(arr)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If arr(i) > arr(j) Then
        strTemp = arr(i)
        arr(i) = arr(j)
        arr(j) = strTemp
      End If
    Next j
  Next i
End Sub

Sub QuickSort(arr, Lo As Long, Hi As Long)
  Dim varPivot As Variant
  Dim varTmp As Variant
  Dim tmpLow As Long
  Dim tmpHi As Long
  tmpLow = Lo
  tmpHi = Hi
  varPivot = arr((Lo + Hi) \ 2)
  Do While tmpLow <= tmpHi
    Do While arr(tmpLow) < varPivot And tmpLow < Hi
      tmpLow = tmpLow + 1
    Loop
    Do While varPivot < arr(tmpHi) And tmpHi > Lo
      tmpHi = tmpHi - 1
    Loop
    If tmpLow <= tmpHi Then
      varTmp = arr(tmpLow)
      arr(tmpLow) = arr(tmpHi)
      arr(tmpHi) = varTmp
      tmpLow = tmpLow + 1
      tmpHi = tmpHi - 1
    End If
  Loop
  If Lo < tmpHi Then QuickSort arr, Lo, tmpHi
  If tmpLow < Hi Then QuickSort arr, tmpLow, Hi
End Sub


```
