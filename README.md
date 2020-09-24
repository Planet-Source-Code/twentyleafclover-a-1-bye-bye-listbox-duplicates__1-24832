<div align="center">

## A\-1 Bye\-Bye ListBox Duplicates


</div>

### Description

For Beginners. Get rid of unwanted listbox duplicates in a few very simple lines of code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[twentyleafclover](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/twentyleafclover.md)
**Level**          |Intermediate
**User Rating**    |4.0 (20 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/twentyleafclover-a-1-bye-bye-listbox-duplicates__1-24832/archive/master.zip)





### Source Code

```
Dim j As Integer
 j = 0
Do While j < List1.ListCount
 List1.Text = List1.List(j)
 If List1.ListIndex <> j Then
  List1.RemoveItem j
 Else
  j = j + 1
 End If
Loop
End Sub
```

