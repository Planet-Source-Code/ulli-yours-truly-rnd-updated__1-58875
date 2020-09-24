<div align="center">

## Yours Truly \- Rnd \(updated\)


</div>

### Description

This little code snippet returns a truly random sequence of Rnd's
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Intermediate
**User Rating**    |5.0 (50 globes from 10 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ulli-yours-truly-rnd-updated__1-58875/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As TwoLongs) As Long
Private Type TwoLongs
  l1 As Long
  l2 As Long
End Type
Public Function IsCpuSuitable() As Boolean
 Dim c As Currency
  On Error Resume Next
    IsCpuSuitable = CBool(QueryPerformanceFrequency(c))
  On Error GoTo 0
End Function
Public Function TrueRnd() As Single
 'returns a truly random sequence of rnd's
 Dim tl    As TwoLongs
 Dim Seed   As Long
 Dim Tmp    As Long
  Do Until Seed > &H3FFFFFFF
    QueryPerformanceCounter tl
    Tmp = tl.l1 And 1
    QueryPerformanceCounter tl
    If Tmp <> (tl.l1 And 1) Then
      Seed = Seed + Seed + Tmp
    End If
  Loop
  TrueRnd = Rnd(-Seed)
End Function
```

