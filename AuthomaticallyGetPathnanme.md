## GetTextFileData

``` VBA
Function GetTextFileData()
  Dim FromMYPath As String
  FromMYPath = Application.CurrentProject.Path

 DoCmd.TransferText transfertype:=acImportDelim, _
  specificationname:="EorpatC", _
  tablename:="Eorpat", _
  FileName:=FromMYPath & "\EORPat.txt", _
  hasfieldnames:=False
End Function
```
