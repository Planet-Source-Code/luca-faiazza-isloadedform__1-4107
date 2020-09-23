<div align="center">

## IsLoadedForm


</div>

### Description

Tells whether a form is loaded or not
 
### More Info
 
ByVal pObjForm As Form

Boolean value


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Luca Faiazza](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/luca-faiazza.md)
**Level**          |Unknown
**User Rating**    |4.4 (22 globes from 5 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/luca-faiazza-isloadedform__1-4107/archive/master.zip)





### Source Code

```
Public Function IsLoadedForm(ByVal pObjForm As Form) As Boolean
  Dim tmpForm As Form
  For Each tmpForm In Forms
    If tmpForm Is pObjForm Then
      IsLoadedForm = True
      Exit For
    End If
  Next
End Function
```

