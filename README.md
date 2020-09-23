<div align="center">

## qsort


</div>

### Description

Want to sort 5,000 10-byte strings in about 1/10th of a second? This will do it (at least on my PII-233!). The insertion sort manages the same task in about 60 seconds (even when optimized it still took about 15 seconds on the same machine).
 
### More Info
 
strList (a string array)

strList (the same array - sorted)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Shaffer](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-shaffer.md)
**Level**          |Unknown
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-shaffer-qsort__1-897/archive/master.zip)





### Source Code

```
Public Function QSort(strList() As String, lLbound As Long, lUbound As Long)
 ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
 ':::                                :::'
 '::: Routine:  QSort                       :::'
 '::: Author:  Mike Shaffer (after Rod Stephens, et al.)     :::'
 '::: Date:   21-May-98                     :::'
 '::: Purpose:  Very fast sort of a string array         :::'
 '::: Passed:  strList  String array              :::'
 ':::       lLbound  Lower bound to sort (usually 1)     :::'
 ':::       lUbound  Upper bound to sort (usually ubound()) :::'
 '::: Returns:  strList  (in sorted order)            :::'
 '::: Copyright: Copyright *c* 1998, Mike Shaffer         :::'
 ':::       ALL RIGHTS RESERVED WORLDWIDE           :::'
 ':::       Permission granted to use in any non-commercial  :::'
 ':::       product with credit where due. For free      :::'
 ':::       commercial license contact mshaffer@nkn.net    :::'
 '::: Revisions: 22-May-98 Added and then dropped revision     :::'
 ':::       using CopyMemory rather than the simple swap   :::'
 ':::       when it was found to not provide much benefit.  :::'
 ':::                                :::'
 ':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
 Dim strTemp As String
 Dim strBuffer As String
 Dim lngCurLow As Long
 Dim lngCurHigh As Long
 Dim lngCurMidpoint As Long
 lngCurLow = lLbound              ' Start current low and high at actual low/high
 lngCurHigh = lUbound
 If lUbound <= lLbound Then Exit Function   ' Error!
 lngCurMidpoint = (lLbound + lUbound) \ 2   ' Find the approx midpoint of the array
 strTemp = strList(lngCurMidpoint)       ' Pick as a starting point (we are making
                        ' an assumption that the data *might* be
                        ' in semi-sorted order already!
 Do While (lngCurLow <= lngCurHigh)
   Do While strList(lngCurLow) < strTemp
      lngCurLow = lngCurLow + 1
      If lngCurLow = lUbound Then Exit Do
   Loop
   Do While strTemp < strList(lngCurHigh)
      lngCurHigh = lngCurHigh - 1
      If lngCurHigh = lLbound Then Exit Do
   Loop
   If (lngCurLow <= lngCurHigh) Then     ' if low is <= high then swap
     strBuffer = strList(lngCurLow)
     strList(lngCurLow) = strList(lngCurHigh)
     strList(lngCurHigh) = strBuffer
     '
     lngCurLow = lngCurLow + 1       ' CurLow++
     lngCurHigh = lngCurHigh - 1      ' CurLow--
   End If
 Loop
 If lLbound < lngCurHigh Then         ' Recurse if necessary
   QSort strList(), lLbound, lngCurHigh
 End If
 If lngCurLow < lUbound Then          ' Recurse if necessary
    QSort strList(), lngCurLow, lUbound
 End If
End Function
```

