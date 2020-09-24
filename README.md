<div align="center">

## Close/open the CD Rom door


</div>

### Description

Open and close the CD rom door from code! Note:comment out the unneeded api declaration (16 or 32 bit) depending on what operating system you are using!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Visual Basic On Line User Group](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/visual-basic-on-line-user-group.md)
**Level**          |Beginner
**User Rating**    |2.8 (11 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/visual-basic-on-line-user-group-close-open-the-cd-rom-door__1-432/archive/master.zip)

### API Declarations

```
'for 16 bit windows
Declare Function mcisendstring Lib "MMSystem" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal wReturnLength As Integer, ByVal hCallback As Integer) As Long
'Note:FOR 95/98/NT using the following:
Public Declare Function
mciSendString Lib "winmm.dll" Alias
"mciSendStringA" (ByVal lpstrCommand As
String, ByVal lpstrReturnString As
String, ByVal uReturnLength As Long,
ByVal hwndCallback As Long) As
Long
```


### Source Code

```

'to open it:
x=
mciSendString("set cd door open", 0&,
0, 0)
'to close it:
x = mciSendString("set
cd door closed", 0&, 0, 0)
```

