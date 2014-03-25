## Split function to parse long string 

I often have a need to read an ascii file and process it in Excel. 
I have used this funtion to address that problem.
I read a line of the file into a string variable and then call this function to parse the string into "words".
Note the line Const CHARS.  The characters in this variable can be adjusted to identify the delimiting 
character(s) in the string.


```
Public Function Split(ByVal InputText As String, _
         Optional ByVal Delimiter As String) As Variant

' This function splits the sentence in InputText into
' words and returns a string array of the words. Each
' element of the array contains one word.
'
' found on internet at
' http://snipplr.com/view/65504/alternative-split-function/
'

    ' This constant contains punctuation and characters
    ' that should be filtered from the input string.
    Const CHARS = "!?;:""'()[]{}"
    Dim strReplacedText As String
    Dim intIndex As Integer

    ' Replace tab characters with space characters.
    strReplacedText = Trim(Replace(InputText, _
         vbTab, " "))

    ' Filter all specified characters from the string.
    For intIndex = 1 To Len(CHARS)
        strReplacedText = Trim(Replace(strReplacedText, _
            Mid(CHARS, intIndex, 1), " "))
    Next intIndex

    ' Loop until all consecutive space characters are
    ' replaced by a single space character.
    Do While InStr(strReplacedText, "  ")
        strReplacedText = Replace(strReplacedText, _
            "  ", " ")
    Loop

    ' Split the sentence into an array of words and return
    ' the array. If a delimiter is specified, use it.
    'MsgBox "String:" & strReplacedText
    If Len(Delimiter) = 0 Then
        Split = VBA.Split(strReplacedText)
    Else
        Split = VBA.Split(strReplacedText, Delimiter)
    End If
End Function
```
