<div align="center">

## Check for existing Directories/Folders


</div>

### Description

Check if that directory exists before running the risk of an error and/or data loss. One of the few that really works. No API, no function calls. Existence check and logic included. Incredibly simple.
 
### More Info
 
None known


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[N/A](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/empty.md)
**Level**          |Beginner
**User Rating**    |4.3 (43 globes from 10 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/check-for-existing-directories-folders__1-2600/archive/master.zip)





### Source Code

```
'By Jim Sivage
'
'ISO Global
'http://www.isoglobal.com
'
'
'Make f$ equal to folder you're testing.
'
f$ = "C:\WINDOWS"
dirFolder = Dir(f$, vbDirectory)
If dirFolder <> "" Then
 strmsg = MsgBox("This folder already exists.", vbCritical):goto optout
End If
'directory exists action here
optout:
```

