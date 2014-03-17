codeSnippetsVBA
===============

VBA Snippet Archiving

Generate an Index to allow the storing of data that matches the index
see xlsm file read_wellon_debug xxx

```VBA
Public indexMax As Long                   'This tracks the last position used in the playIndex array
Public playIndex(2000) As Double          'This is the array that holds the index -- max indexes = 2000

Public Function getIndex(iPlay)
' this function is to return unique index for each play since the plays are not in a sorted order
'
' loop through the already assigned indexes

    'MsgBox ("i am here : " & iPlay)
    'getIndex = 0
    tempIndex = 0
    For i = 1 To indexMax
      If iPlay = playIndex(i) Then tempIndex = i
    Next
    
    If tempIndex = 0 Then
        tempIndex = indexMax + 1
        indexMax = indexMax + 1
        playIndex(indexMax) = iPlay
    End If
    
    
    getIndex = tempIndex

End Function
```


check out our wiki page [codeSnippetsVBA wiki](https://github.com/LessGoudarzi/codeSnippetsVBA.wiki.git)
