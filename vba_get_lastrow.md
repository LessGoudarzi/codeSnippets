## Can't tell you how ofen I need to remember how to do this

```VBA

tempR = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
tempC = Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

```
