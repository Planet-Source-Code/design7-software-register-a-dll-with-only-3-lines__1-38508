<div align="center">

## Register a DLL with only 3 lines\!


</div>

### Description

Register your DLLs with only 3 lines of code, simple and effective.

Please vote and leave your comments.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Design7 Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/design7-software.md)
**Level**          |Beginner
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/design7-software-register-a-dll-with-only-3-lines__1-38508/archive/master.zip)





### Source Code

```
'Set the directory of your DLL
dllpath="C:\Windows\mydll.dll"
'Call the FileSystemObject for
'retrieving the Windows Folder
Set fso=CreateObject("Scripting.FileSystemObject")
'Call the Reg DLL Server in the Windows dir
'and pass the DLL path as a parameter
Shell(fso.GetSpecialFolder(0) & "\regsvr32.exe " & dllpath, vbHide)
```

