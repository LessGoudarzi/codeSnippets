### Can't tell you how ofen I need to remember how to do this

**_I got this from Stackoverflow but can't find the link_**

```VBA

tempR = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
tempC = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

```
