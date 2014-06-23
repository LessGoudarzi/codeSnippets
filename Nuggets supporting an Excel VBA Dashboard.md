Building an Excel Dashboard using VBA
======================================

Below are a series of VBA sub and functions used in a draft implementation

This is designed to use normalized data extracted from a GAMS GDX file using the GDXXRW.exe utility routine
see c_reporting.xlsm

This includes the use of:
* indices, 
* bubblesort and
* code generated data dependent checkboxes for UI

Note, there is some dead subs that have not been eliminated at this point

More cleanup and description to follow

```VBA

Public refMax, crudeMax, yearMax, productMax As Long                   'This tracks the last position used in the playIndex array
Public refIndex(10), crudeIndex(10), yearIndex(30), index(30), productIndex(10) As String        'This is the array that holds the index -- max indexes = 2000

Public Function getRefIndex(iPlay)
' this function is to return unique index for each play since the plays are not in a sorted order

    tempIndex = 0
    For i = 1 To refMax
      If iPlay = refIndex(i) Then tempIndex = i
    Next
    
    

    If tempIndex = 0 Then
        tempIndex = refMax + 1
        refMax = refMax + 1
        refIndex(refMax) = iPlay
        'MsgBox (tempIndex & " : " & iPlay)
    End If


    getRefIndex = tempIndex

End Function
Public Function getIndex(iPlay, index, indexMax)
' this function is to return unique index for each play since the plays are not in a sorted order

    tempIndex = 0
    For i = 1 To indexMax
      If iPlay = index(i) Then tempIndex = i
    Next
    

    If tempIndex = 0 Then
        tempIndex = indexMax + 1
        indexMax = indexMax + 1
        index(indexMax) = iPlay
        'MsgBox (tempIndex & " : " & iPlay)
    End If


    getIndex = tempIndex

End Function
Public Function getPosition(iPlay, index, indexMax)
' this function is to return the position in the index array

    i = 0
    For i = 1 To indexMax
      If iPlay = index(i) Then tempIndex = i
    Next

    getPosition = tempIndex

End Function

Sub get_crude_use()

        Application.ScreenUpdating = False

        ' INITIALIZE THE STAGIING SHEET
        '
        
            Sheets("Staging").Select
            Cells.Select
            Selection.ClearContents
            Range("A1").Select
        
            ' GET RID OF CHECK BOXES -- REMEMBER THEIR POSITION WILL BE THE SAME
            Application.ScreenUpdating = True

             Sheets("Dashboard").Select
            'If Sheets("Dashboard").CheckBoxes.Count > 0 Then
           For Each c In Sheets("Dashboard").CheckBoxes
              c.Delete
            Next
           ' End If
           Application.ScreenUpdating = False


        Sheets("crudeInputs").Select

        ' initialize the array to clean it up
        For i = 1 To 30
            index(i) = ""
            Next i
            indexMax = 0
            refMax = 0
            crudeMax = 0
            yearMax = 0
        
        tempr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row ' max row
        'tempC = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column 'max column
        
        Application.StatusBar = "Reading in the crude use data ......"
       

        ' Cycle down rows to collect unique values
        For i = 2 To tempr
            If Cells(i, 3) = "Per1" Then
               rIndex = getIndex(Cells(i, 1), refIndex, refMax)
               cIndex = getIndex(Cells(i, 2), crudeIndex, crudeMax)
               yIndex = getIndex(Cells(i, 4), yearIndex, yearMax)
                
              ' x = x & "*"
             '  Application.StatusBar = "The row is " & x
              End If
            Next i
            
         ReDim crudeValues(yearMax, refMax, crudeMax)
         
         Call BubbleSort(yearIndex, yearMax) ' modified one provided to pass the upper limit of array being sorted
         Call BubbleSort(refIndex, refMax) ' modified one provided to pass the upper limit of array being sorted
         Call BubbleSort(crudeIndex, crudeMax) ' modified one provided to pass the upper limit of array being sorted
         
         ' Create array holding the data in sorted order
         For i = 2 To tempr
          If Cells(i, 3) = "Per1" Then
            yIndex = getPosition(Cells(i, 4), yearIndex, yearMax)
            rIndex = getPosition(Cells(i, 1), refIndex, refMax)
            cIndex = getPosition(Cells(i, 2), crudeIndex, crudeMax)
            crudeValues(yIndex, rIndex, cIndex) = Cells(i, 5)
            End If
            'MsgBox (yIndex & " : " & Cells(i, 4))
         
         Next i
        
        Application.StatusBar = "Filling array with crude use data ......"
        ' output the crudevalues table ===================
        ' by year, refiner stacked crude volumes by type
        Sheets("staging").Select
        sRow = 2
        scolumn = 2
        cRow = sRow
       For i = 1 To crudeMax
        Cells(sRow, scolumn + 1 + i) = crudeIndex(i)
       Next i
        
       For i = 1 To yearMax
            cRow = cRow + 1
            Cells(cRow, scolumn) = yearIndex(i)
            For j = 1 To refMax
            Cells(cRow, scolumn + 1) = refIndex(j)
            
            For k = 1 To crudeMax
             Cells(cRow, scolumn + 1 + k) = crudeValues(i, j, k)
               
            Next k
            cRow = cRow + 1
            Next j
            Next i
            
        lRow = cRow
        lCol = scolumn + 1 + crudeMax
        
        Application.StatusBar = "Linking crude use data to first graph......"
        
        Call crudeInputsGraph(sRow, scolumn, lRow, lCol)
        
        Application.ScreenUpdating = True
        Application.ScreenUpdating = False
        
        ' output crude data table with different layout
        ' this time by refinery, year stacked crude volumes by type
        Sheets("staging").Select
        sRow = 2
        scolumn = lCol + 3
        cRow = sRow
        
       For i = 1 To crudeMax
        Cells(sRow, scolumn + 1 + i) = crudeIndex(i)
       Next i
        
         
         For j = 1 To refMax
            Cells(cRow + 1, scolumn) = refIndex(j)
            
         For i = 1 To yearMax
            cRow = cRow + 1
            Cells(cRow, scolumn + 1) = yearIndex(i)
            
            For k = 1 To crudeMax
             Cells(cRow, scolumn + 1 + k) = crudeValues(i, j, k)
               
            Next k
            
            Next i
            cRow = cRow + 1
            
            Next j
       
            
        lRow = cRow
        lCol = scolumn + 1 + crudeMax
        
        Application.StatusBar = "Linking crude use data to second graph......"
        
        Call crudeInputsGraph_1(sRow, scolumn, lRow, lCol)
        
        Application.ScreenUpdating = True
        Application.ScreenUpdating = False
        
        Application.StatusBar = "Processing RAC pirce data and graph......"
        
        Call get_RAC
        
        Application.StatusBar = "Processing Product pirce data and graph......"
        
        Call get_ProductPrices
        
        Sheets("Dashboard").Select
        Range("A1").Select
        
        Application.ScreenUpdating = True
        Application.StatusBar = False
 '
End Sub
Sub crudeInputsGraph_1(sRow, sCol, lRow, lCol)
'
' crudeInputsGraph Macro
'

'
    Sheets("Staging").Select
    
    rValue = letter(sCol) & sRow & ":" & letter(lCol) & lRow
        
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Chart 4").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Sheets("Staging").Range(rValue)
   
End Sub
Sub crudeInputsGraph(sRow, sCol, lRow, lCol)
'
' crudeInputsGraph Macro
'

'
    Sheets("Staging").Select
    'tempr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    'tempc = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column

    ActiveWorkbook.Names.Add Name:="crudex", RefersToR1C1:= _
        "=Staging!R" & sRow & "C" & sCol & ":R" & lRow & "C" & lCol
        
    'MsgBox ("letter is " & letter(tempc) & " for value " & tempc)
    rValue = letter(sCol) & sRow & ":" & letter(lCol) & lRow
    'MsgBox (rValue)
        
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Chart 5").Activate
    ActiveChart.PlotArea.Select
   ActiveChart.SetSourceData Source:=Sheets("Staging").Range(rValue)
End Sub
Sub get_RAC()

        'Sheets("Staging").Select
        'Cells.Select
        'Selection.ClearContents
        'Range("A1").Select

        Sheets("crudeRAC").Select

        ' initialize the array to clean it up
        For i = 1 To 30
            index(i) = ""
            Next i
            indexMax = 0
            refMax = 0
          '  crudeMax = 0
            yearMax = 0
        
        tempr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row ' max row
        'tempC = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column 'max column
        
       

        ' Cycle down rows to collect unique values
        For i = 2 To tempr
            If Cells(i, 3) = "Per1" Then
               rIndex = getIndex(Cells(i, 1), refIndex, refMax)
             '  cIndex = getIndex(Cells(i, 2), crudeIndex, crudeMax)
               yIndex = getIndex(Cells(i, 4), yearIndex, yearMax)
                
              End If
            Next i
            
         ReDim racValues(yearMax, refMax)
         
         Call BubbleSort(yearIndex, yearMax) ' modified one provided to pass the upper limit of array being sorted
         Call BubbleSort(refIndex, refMax) ' modified one provided to pass the upper limit of array being sorted
        ' Call BubbleSort(crudeIndex, crudeMax) ' modified one provided to pass the upper limit of array being sorted
         
         ' Create array holding the data in sorted order
         For i = 2 To tempr
          If Cells(i, 3) = "Per1" Then
            yIndex = getPosition(Cells(i, 4), yearIndex, yearMax)
            rIndex = getPosition(Cells(i, 1), refIndex, refMax)
          '  cIndex = getPosition(Cells(i, 2), crudeIndex, crudeMax)
            racValues(yIndex, rIndex) = Cells(i, 5)
            End If
            'MsgBox (yIndex & " : " & Cells(i, 4))
         
         Next i
        
        
        ' output the RAC values table ===================
        ' by year and refiner
        Sheets("staging").Select
        sRow = 2
        scolumn = 20
        cRow = sRow
        
       For i = 1 To refMax
        Cells(sRow, scolumn + i) = refIndex(i)
       Next i
        
       For i = 1 To yearMax
            cRow = cRow + 1
            Cells(cRow, scolumn) = yearIndex(i)
            For j = 1 To refMax
            Cells(cRow, scolumn + j) = racValues(i, j)
    
            Next j
            Next i
            
        For i = 1 To refMax
            Call insertCheckBox(i)
        Next i
            
        lRow = cRow
        lCol = scolumn + refMax
        
        'MsgBox (lRow & " : " & lCol)
        Call RAC_Prices_1(sRow, scolumn, lRow, lCol)
        
End Sub
Sub RAC_Prices_1(sRow, sCol, lRow, lCol)
'

    'Sheets("Staging").Select
    
    rValue = letter(sCol) & sRow & ":" & letter(lCol) & lRow
        
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Chart 7").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Sheets("Staging").Range(rValue)
    Range("a1").Select
   
End Sub
Function letter(x)
  letter = Mid("abcdefghijklmnopqrstuvwxyz", x, 1)
End Function
Sub insertCheckBox(i)
'
' insertCheckBox Macro

    Sheets("Dashboard").Select
    y = 900 + (i - 1) * 20
    ActiveSheet.CheckBoxes.Add(925.5, y, 72, 20).Select
    Selection.Characters.Text = refIndex(i)
    Selection.Name = "Set 1 " & refIndex(i)
    Selection.Value = True
    ActiveSheet.Shapes("Set 1 " & refIndex(i)).Select
    Selection.OnAction = "Hide_ref1"

End Sub
Sub Hide_ref1()
    Application.ScreenUpdating = False
    Sheets("Staging").Select
    
    If Sheets("Dashboard").CheckBoxes("Set 1 1_RefReg").Value = Checked Then
        Columns("U:U").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("U:U").Select
        Selection.EntireColumn.Hidden = True
    End If
    
     If Sheets("Dashboard").CheckBoxes("Set 1 2_RefReg").Value = Checked Then
        Columns("V:V").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("V:V").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    If Sheets("Dashboard").CheckBoxes("Set 1 3_RefReg").Value = Checked Then
        Columns("W:W").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("W:W").Select
        Selection.EntireColumn.Hidden = True
    End If
    
     If Sheets("Dashboard").CheckBoxes("Set 1 4_RefReg").Value = Checked Then
        Columns("X:X").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("X:X").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    Sheets("Dashboard").Select
    Application.ScreenUpdating = True
    
End Sub
Sub get_ProductPrices()

        'Sheets("Staging").Select
        'Cells.Select
        'Selection.ClearContents
        'Range("A1").Select

        Sheets("specProdPrices").Select

        ' initialize the array to clean it up
        For i = 1 To 30
            index(i) = ""
            Next i
            
            indexMax = 0
            refMax = 0
            productMax = 0
            yearMax = 0
        
        tempr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row ' max row
        'tempC = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column 'max column
        
       

        ' Cycle down rows to collect unique values
        For i = 2 To tempr
            If Cells(i, 4) = "Per1" Then
               rIndex = getIndex(Cells(i, 1), refIndex, refMax)
               pindex = getIndex(Cells(i, 3), productIndex, productMax)
               yIndex = getIndex(Cells(i, 5), yearIndex, yearMax)
                
              End If
            Next i
            
        'MsgBox (" Max values for year, refinery and product = " & yearMax & " : " & refMax & " : " & productMax)
            
         ReDim productValues(yearMax, refMax, productMax)
         
         Call BubbleSort(yearIndex, yearMax) ' modified one provided to pass the upper limit of array being sorted
         Call BubbleSort(refIndex, refMax) ' modified one provided to pass the upper limit of array being sorted
         Call BubbleSort(productIndex, productMax) ' modified one provided to pass the upper limit of array being sorted
         
         'MsgBox (productMax)
         
         ' Create array holding the data in sorted order
         For i = 2 To tempr
          If Cells(i, 4) = "Per1" Then
            yIndex = getPosition(Cells(i, 5), yearIndex, yearMax)
            rIndex = getPosition(Cells(i, 1), refIndex, refMax)
            pindex = getPosition(Cells(i, 3), productIndex, productMax)
            productValues(yIndex, rIndex, pindex) = Cells(i, 6)
            
            End If
            'MsgBox (yIndex & " : " & Cells(i, 4))
         
         Next i
        
    
        
        ' output the product values table ===================
        ' by year and refiner
        Sheets("staging").Select
        sRow = 2
        scolumn = 27
        cRow = sRow
        
       tCol = scolumn + 1
       For i = 1 To refMax
           tCol = tCol
           Cells(sRow, tCol) = refIndex(i)
        
            For j = 1 To productMax
                Cells(sRow + 1, tCol) = productIndex(j)
                tCol = tCol + 1
            Next j
       
       Next i
        
        cRow = cRow + 1 ' because used extra heading row
        
    '  For y = 1 To yearMax
    '  For r = 1 To refMax
    '  tempstr = ""
    '  For p = 1 To productMax
    '               tempstr = tempstr & productValues(y, r, p) & vbCrLf
    '  Next p
    '    MsgBox tempstr
    '  Next r
    '  Next y
      
        
       For i = 1 To yearMax
           tCol = scolumn + 1
           cRow = cRow + 1
           Cells(cRow, scolumn) = yearIndex(i)
           
          For j = 1 To refMax
               For k = 1 To productMax
               'MsgBox ("values of year, refinery and product are = " & i & " : " & j & " : " & k)
              ' MsgBox productValues(i, j, k)
               Cells(cRow, tCol) = productValues(i, j, k)
                tCol = tCol + 1
                Next k
            Next j
          Next i
            
       For i = 1 To refMax
            Call insertCheckBox2(i)
        Next i
            
        For i = 1 To productMax
            Call insertCheckBox3(i)
        Next i
            
        lRow = cRow
        lCol = scolumn + refMax * productMax
        
        'MsgBox (sRow & " : " & scolumn & " ending at " & lRow & " : " & lCol)
        Call product_Prices_1(sRow, scolumn, lRow, lCol)
        
End Sub
Sub product_Prices_1(sRow, sCol, lRow, lCol)
'

    'Sheets("Staging").Select
    
                 tval = (sCol)
                 If tval <= 26 Then cSelect = letter(tval)
                 If (tval > 26 And tval <= 52) Then cSelect = "A" & letter(tval - 26)
                 If (tval > 52 And tval <= 78) Then cSelect = "B" & letter(tval - 52)
                 
                  tval = (lCol)
                  If tval <= 26 Then cSelect2 = letter(tval)
                  If (tval > 26 And tval <= 52) Then cSelect2 = "A" & letter(tval - 26)
                  If (tval > 52 And tval <= 78) Then cSelect2 = "B" & letter(tval - 52)
    
    rValue = cSelect & sRow & ":" & cSelect2 & lRow
    MsgBox ("hey hey " & rValue)
        
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Chart 8").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Sheets("Staging").Range(rValue)
    Range("a1").Select
   
End Sub

Sub insertCheckBox3(i)
'
' insertCheckBox Macro

    Sheets("Dashboard").Select
    y = 1500 + (i - 1) * 20
    ActiveSheet.CheckBoxes.Add(925.5, y, 72, 20).Select
    Selection.Characters.Text = productIndex(i)
    Selection.Name = productIndex(i)
    Selection.Value = True
    'ActiveSheet.Shapes("My Refinery 2" & i).Select
    Selection.OnAction = "displayProductPrices"

End Sub
Sub insertCheckBox2(i)
'
' insertCheckBox Macro

    Sheets("Dashboard").Select
    y = 1320 + (i - 1) * 20
    ActiveSheet.CheckBoxes.Add(925.5, y, 72, 20).Select
    Selection.Characters.Text = refIndex(i)
    Selection.Name = refIndex(i)
    Selection.Value = True
    'ActiveSheet.Shapes("My Refinery 2" & i).Select
    Selection.OnAction = "displayProductPrices"

End Sub
Sub Hide_ref2()
    Application.ScreenUpdating = False
    Sheets("Staging").Select
    
    If Sheets("Dashboard").CheckBoxes("1_RefReg").Value = Checked Then
        Columns("AB:AI").Select
        Selection.EntireColumn.Hidden = False
       'Call Hide_product1
        Else
        Columns("AB:AI").Select
        Selection.EntireColumn.Hidden = True
         'Call Hide_product1
    End If
    
     If Sheets("Dashboard").CheckBoxes("2_RefReg").Value = Checked Then
        Columns("Aj:AQ").Select
        Selection.EntireColumn.Hidden = False
        ' Call Hide_product1
        Else
        Columns("Aj:AQ").Select
        Selection.EntireColumn.Hidden = True
         'Call Hide_product1
    End If
    
    If Sheets("Dashboard").CheckBoxes("3_RefReg").Value = Checked Then
        Columns("AR:AY").Select
        Selection.EntireColumn.Hidden = False
         'Call Hide_product1
        Else
        Columns("AR:AY").Select
        Selection.EntireColumn.Hidden = True
        'Call Hide_product1
    End If
    
     If Sheets("Dashboard").CheckBoxes("4_RefReg").Value = Checked Then
        Columns("AZ:BG").Select
        Selection.EntireColumn.Hidden = False
         'Call Hide_product1
        Else
        Columns("AZ:BG").Select
        Selection.EntireColumn.Hidden = True
         'Call Hide_product1
    End If
    
    Sheets("Dashboard").Select
    Application.ScreenUpdating = True
    
End Sub
Sub Hide_product1()
    Application.ScreenUpdating = False
    Sheets("Staging").Select
    
    tempstr = "refMax = " & refMax & vbCrLf
    tempstr = tempstr & "productMax = " & productMax & vbCrLf
    
    If Sheets("Dashboard").CheckBoxes(productIndex(1)).Value = Checked Then  ' CBOB
          scolumn = 27
          Call unhideProduct(scolumn)
        Else
          scolumn = 27
           Call hideProduct(scolumn)
    End If
    
    If Sheets("Dashboard").CheckBoxes(productIndex(2)).Value = Checked Then  ' CBOB
          scolumn = 28
          Call unhideProduct(scolumn)
        Else
          scolumn = 28
           Call hideProduct(scolumn)
    End If
    
     If Sheets("Dashboard").CheckBoxes(productIndex(3)).Value = Checked Then  ' CBOB
          scolumn = 29
          Call unhideProduct(scolumn)
        Else
          scolumn = 29
           Call hideProduct(scolumn)
    End If
    
    If Sheets("Dashboard").CheckBoxes(productIndex(4)).Value = Checked Then  ' CBOB
          scolumn = 30
          Call unhideProduct(scolumn)
        Else
          scolumn = 30
           Call hideProduct(scolumn)
    End If
    
     If Sheets("Dashboard").CheckBoxes(productIndex(5)).Value = Checked Then  ' CBOB
          scolumn = 31
          Call unhideProduct(scolumn)
        Else
          scolumn = 31
           Call hideProduct(scolumn)
    End If
    
    If Sheets("Dashboard").CheckBoxes(productIndex(6)).Value = Checked Then  ' CBOB
          scolumn = 32
          Call unhideProduct(scolumn)
        Else
          scolumn = 32
           Call hideProduct(scolumn)
    End If
    
     If Sheets("Dashboard").CheckBoxes(productIndex(7)).Value = Checked Then  ' CBOB
          scolumn = 33
          Call unhideProduct(scolumn)
        Else
          scolumn = 33
           Call hideProduct(scolumn)
    End If
    
    If Sheets("Dashboard").CheckBoxes(productIndex(8)).Value = Checked Then  ' CBOB
          scolumn = 34
          Call unhideProduct(scolumn)
        Else
          scolumn = 34
           Call hideProduct(scolumn)
    End If
    
    
    
    Sheets("Dashboard").Select
    'Application.ScreenUpdating = True
    
End Sub
Sub hideProduct(scolumn)
 Sheets("Staging").Select
 'scolumn = 27
 'refMax = 4
 'productMax = 8
    
            For i = 1 To refMax
                tval = (scolumn + (i - 1) * productMax + 1)
                If tval <= 26 Then cSelect = letter(tval)
                 If (tval > 26 And tval <= 52) Then cSelect = "A" & letter(tval - 26)
                 If (tval > 52 And tval <= 78) Then cSelect = "B" & letter(tval - 52)
                 Columns(cSelect & ":" & cSelect).Select
                 Selection.EntireColumn.Hidden = True
                 tempstr = tempstr & cSelect & vbCrLf
            Next i
           ' MsgBox tempstr

End Sub
Sub unhideProduct(scolumn)
 Sheets("Staging").Select
 'scolumn = 27
 'refMax = 4
 'productMax = 8
    
            For i = 1 To refMax
                tval = (scolumn + (i - 1) * productMax + 1)
                If tval <= 26 Then cSelect = letter(tval)
                 If (tval > 26 And tval <= 52) Then cSelect = "A" & letter(tval - 26)
                 If (tval > 52 And tval <= 78) Then cSelect = "B" & letter(tval - 52)
                 Columns(cSelect & ":" & cSelect).Select
                 Selection.EntireColumn.Hidden = False
                 tempstr = tempstr & cSelect & vbCrLf
            Next i
            'MsgBox tempstr

End Sub
Sub testSub2()
 Application.ScreenUpdating = False
    Sheets("Staging").Select
    
    tempstr = "refMax = " & refMax & vbCrLf
    tempstr = tempstr & "productMax = " & productMax & vbCrLf
    
    If Sheets("Dashboard").CheckBoxes(9).Value = Checked Then  ' CBOB
          scolumn = 27
           ' MsgBox "checked"
        Else
          scolumn = 27
           Call hideProduct(scolumn)

    End If
    
     If Sheets("Dashboard").CheckBoxes(10).Value = Checked Then
        Columns("Aj:AQ").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("Aj:AQ").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    If Sheets("Dashboard").CheckBoxes(11).Value = Checked Then
        Columns("AR:AY").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("AR:AY").Select
        Selection.EntireColumn.Hidden = True
    End If
    
     If Sheets("Dashboard").CheckBoxes(12).Value = Checked Then
        Columns("AZ:BG").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("AZ:BG").Select
        Selection.EntireColumn.Hidden = True
    End If
    
     If Sheets("Dashboard").CheckBoxes(13).Value = Checked Then
        Columns("AB:AI").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("AB:AI").Select
        Selection.EntireColumn.Hidden = True
    End If
    
     If Sheets("Dashboard").CheckBoxes(14).Value = Checked Then
        Columns("Aj:AQ").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("Aj:AQ").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    If Sheets("Dashboard").CheckBoxes(15).Value = Checked Then
        Columns("AR:AY").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("AR:AY").Select
        Selection.EntireColumn.Hidden = True
    End If
    
     If Sheets("Dashboard").CheckBoxes(16).Value = Checked Then
        Columns("AZ:BG").Select
        Selection.EntireColumn.Hidden = False
        Else
        Columns("AZ:BG").Select
        Selection.EntireColumn.Hidden = True
    End If
    
    Sheets("Dashboard").Select
    Application.ScreenUpdating = True

End Sub
Sub displayProductPrices()


    Application.ScreenUpdating = False
    Sheets("Staging").Select

        scolumn = 27
        For r = 1 To refMax
        ' check to see if region is to be displayed
        ' if not, turn all products prices in region off (hide columns)
            If Sheets("Dashboard").CheckBoxes(refIndex(r)).Value <> Checked Then
                For p = 1 To productMax
                        tval = (scolumn + (r - 1) * productMax + p)
                        If tval <= 26 Then cSelect = letter(tval)
                        If (tval > 26 And tval <= 52) Then cSelect = "A" & letter(tval - 26)
                        If (tval > 52 And tval <= 78) Then cSelect = "B" & letter(tval - 52)
                        Columns(cSelect & ":" & cSelect).Select
                        Selection.EntireColumn.Hidden = True
                 Next p
            Else
                For p = 1 To productMax
                        tval = (scolumn + (r - 1) * productMax + p)
                        If tval <= 26 Then cSelect = letter(tval)
                        If (tval > 26 And tval <= 52) Then cSelect = "A" & letter(tval - 26)
                        If (tval > 52 And tval <= 78) Then cSelect = "B" & letter(tval - 52)
                        ' check to see whether to display products
                        If Sheets("Dashboard").CheckBoxes(productIndex(p)).Value <> Checked Then
                                    Columns(cSelect & ":" & cSelect).Select
                                    Selection.EntireColumn.Hidden = True
                                Else
                                    Columns(cSelect & ":" & cSelect).Select
                                    Selection.EntireColumn.Hidden = False
                                End If
                Next p
            
            End If
            
        Next r
        
        
        Sheets("Dashboard").Select
        

End Sub
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

```



