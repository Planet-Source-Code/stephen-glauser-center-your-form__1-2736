<div align="center">

## \!Center Your Form\!


</div>

### Description

This will Center your form on your screen. It will also higher it a tiny bit, so it looks centered above the TaskBar.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stephen Glauser](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stephen-glauser.md)
**Level**          |Unknown
**User Rating**    |3.1 (25 globes from 8 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stephen-glauser-center-your-form__1-2736/archive/master.zip)





### Source Code

```
Sub FormCenter (Frm As Form)
Rem Written by Stephen Glauser
Rem Last Update: August 1, 1999
Rem Centers your form on the Screen
Rem FormCenter Me
Frm.Top = (Screen.Height * .85) / 2 - Frm.Height / 2
Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub
```

