<div align="center">

## Bulk Un/Register DLLs & OCXs


</div>

### Description

generate two batch files (one to register, one to unregister) that manage dlls/ocxs in a dev branch.

the batch files execute and give the line number of a failed registry event. this is a two minute script that saves me time in trying to see what new dlls/ocxs were added/deleted...hope it is useful to all.
 
### More Info
 
this is a .vbs script...which you can easily convert to a vb exe...


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steven Hoyt](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steven-hoyt.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB Script
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steven-hoyt-bulk-un-register-dlls-ocxs__1-27451/archive/master.zip)





### Source Code

```
Private mfso
Private mfd
Private mRegStream
Private mURegStream
Private mcstrHeader
Private mcstrErrText
Private mcstrFooter
Private mlngPtr
mcstrHeader = "@echo off" & vbCrLf & "cls" & vbCrLf & "echo [REG]..." & vbCrLf & vbCrLf & "set has_err=0" & vbCrLf & "set errors=0" & vbCrLf & "set e_text=0" & vbCrLf & vbCrLf
mcstrErrText = "if errorlevel=1 set has_err=1" & vbCrLf & "if errorlevel=1 set errors=[SEQ]" & vbCrLf & "if errorlevel=0 set e_text=%errors%" & vbCrLf & vbCrLf
mcstrFooter = "set e_text=Error, Line %e_text%" & vbCrLf & vbCrLf & "if %has_err%==0 set e_text=No Errors" & vbCrLf & vbCrLf & "cls" & vbCrLf & "echo [REG]!" & vbCrLf & "echo %e_text%" & vbCrLf & "pause"
mlngPtr = 0
mlnghFileReg = 1
mlnghFileUReg = 2
If MsgBox("Create Un/Register Batch Files?", vbYesNo, "Dll Auto-Register") = vbYes Then
  Set mfso = CreateObject("Scripting.FileSystemObject")
  Set mfd = mfso.GetFolder("C:\")
  Set mRegStream = mfd.CreateTextFile("Register.bat", True, False)
  Set mURegStream = mfd.CreateTextFile("UnRegister.bat", True, False)
  mRegStream.Write Replace(mcstrHeader, "[REG]", "Registering")
  mURegStream.Write Replace(mcstrHeader, "[REG]", "Un-Registering")
  SetDllRegText "C:\Your Project\Dev\"
  mRegStream.Write Replace(mcstrFooter, "[REG]", "Registered")
  mURegStream.Write Replace(mcstrFooter, "[REG]", "Un-Registered")
  mRegStream.Close
  mURegStream.Close
  MsgBox "Done.", vbOKOnly, "Dll Auto-Register"
End If
Private Sub SetDllRegText(ByVal strSearchPath)
  Dim dr
  Dim sfld
  Dim f
  Dim strPrintData
  If Right(strSearchPath, 1) <> "\" Then strSearchPath = strSearchPath & "\"
  Set dr = mfso.GetFolder(strSearchPath)
  For Each f In dr.Files
    If Right(LCase(f.Name), 4) = ".dll" Or Right(LCase(f.Name), 4) = ".ocx" Then
      mlngPtr = mlngPtr + 1
      strPrintData = Replace(mcstrErrText, "[SEQ]", mlngPtr)
      mRegStream.Write "regsvr32.exe " & """" & strSearchPath & f.Name & """ /s" & vbCrLf & strPrintData
      mURegStream.Write "regsvr32.exe /u " & """" & strSearchPath & f.Name & """ /s" & vbCrLf & strPrintData
    End If
  Next
  If dr.SubFolders.Count Then
    For Each sfld In dr.SubFolders
      If Err.Number Then Exit For
      SetDllRegText sfld.Path
    Next
  End If
End Sub
```

