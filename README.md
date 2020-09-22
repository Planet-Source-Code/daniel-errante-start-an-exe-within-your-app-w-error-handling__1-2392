<div align="center">

## Start an exe within your app w/ error handling\!


</div>

### Description

starts an exe from within your application
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Daniel Errante](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/daniel-errante.md)
**Level**          |Unknown
**User Rating**    |3.3 (43 globes from 13 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/daniel-errante-start-an-exe-within-your-app-w-error-handling__1-2392/archive/master.zip)





### Source Code

```
Sub LoadEXE(Dir As String)
 On Error GoTo err:
 X% = Shell(Dir, 1): NoFreeze% = DoEvents(): Exit Sub
Exit Sub
err:
'make your own error messages like mine below, or use the default:
If err.Number = 6 Then Exit Sub
MsgBox "Please make sure that the application you are trying to launch is located in the correct folder." & vbCrLf & "If not, do this and retry launching the application.", vbExclamation
 'default: MsgBox "Error:" & vbCrLf & err.Description & vbCrLf & err.Number, vbExclamation
End Sub
```

