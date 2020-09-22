<div align="center">

## Sorting by columnheaders in listviews


</div>

### Description

This allows a user to sort the contents of a listview by clicking on a column header in report-mode. Clicking again will sort descending.
 
### More Info
 
This assumes that a ListView named 'ListView1' exists on a form, and that code exists to put values into it. Also, the ListView style-property must be 'report'.

Paste the code into the 'General Declarations' area of your form (below any other global declarations you may have) and substitute 'ListView1' with the name of your listview.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Hansen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-hansen.md)
**Level**          |Unknown
**User Rating**    |4.2 (162 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-hansen-sorting-by-columnheaders-in-listviews__1-1996/archive/master.zip)





### Source Code

```
Dim iCol As Integer
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As_ MSComctlLib.ColumnHeader)
  ' When a ColumnHeader object is clicked, the ListView control is
  ' sorted by the subitems of that column.
  ' Set the SortKey to the Index of the ColumnHeader - 1
  If ColumnHeader.Index - 1 <> iCol Then
    ListView1.SortOrder = 0
  Else
    ListView1.SortOrder = Abs(ListView1.SortOrder - 1)
  End If
  ListView1.SortKey = ColumnHeader.Index - 1
  ' Set Sorted to True to sort the list.
  ListView1.Sorted = True
  iCol = ColumnHeader.Index - 1
End Sub
```

