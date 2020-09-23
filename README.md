<div align="center">

## Find Any Window


</div>

### Description

Sometimes you need to find a window using the API Call findwindow , but what if this windows caption changes

you can't find that same window all the time. With this function you can find any window just by knowing a few letters

in the caption. This will return the windows' hWnd , also includes a function that will grab the windows caption.

This is something that will be useful to alot of programmers. Updated! 2.23.01
 
### More Info
 
call it like so

call msgbox(FindAnyWindow&(me,"text of window"))

or to get the caption do this

call msgbox(getcaption$(FindAnyWindow&(me,"text of window")))

none(that I know of)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[DoWnLoHo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/downloho.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/downloho-find-any-window__1-2157/archive/master.zip)

### API Declarations

```
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
```


### Source Code

```
Public Function GetCaption(ByVal lhWnd As Long) As String
Dim sA As String, lLen As Long
 lLen& = GetWindowTextLength(lhWnd&)
 sA$ = String(lLen&, 0&)
 Call GetWindowText(lhWnd&, sA$, lLen& + 1)
 GetCaption$ = sA$
End Function
Public Function FindAnyWindow(frm As Form, ByVal WinTitle As String, Optional ByVal CaseSensitive As Boolean = False) As Long
Dim lhWnd As Long, sA As String
lhWnd& = frm.hwnd
Do Until lhWnd& = 0
 DoEvents
 sA$ = GetCaption(lhWnd&)
 If InStr(IIf(CaseSensitive = False, LCase$(sA$), sA$), IIf(CaseSensitive = False, LCase$(WinTitle$), WinTitle$)) Then FindAnyWindow& = lhWnd&: Exit Do Else FindAnyWindow& = 0
 lhWnd& = GetNextWindow(lhWnd&, 2)
Loop
End Function
```

