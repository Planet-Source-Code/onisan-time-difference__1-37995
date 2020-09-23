<div align="center">

## Time Difference


</div>

### Description

Displays the difference between two times in Hours Minutes and seconds
 
### More Info
 
Start time and End Time

Returns time as HH:NN:SS


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Onisan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/onisan.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/onisan-time-difference__1-37995/archive/master.zip)





### Source Code

```
Sub Test()
          'Start Time, End Time
  MsgBox TimeDiff("11:34:29", "20:32:20")
End Sub
Function TimeDiff(STime As Date, ETime As Date) As Date
Dim TimeSecs, Hrs As Double
  'Get Total Number of seconds difference
  TimeSecs = DateDiff("S", STime, ETime)
    'If Difference is a minus(-), add 24 hours worth of seconds.
    If TimeSecs <> Abs(TimeSecs) Then: TimeSecs = TimeSecs + 86400
    'If there are hours get them here
    If TimeSecs >= 3600 Then: Hrs = Fix(TimeSecs / 3600)
    TimeDiff = TimeSerial(Hrs, 0, TimeSecs - (Hrs * 3600))
End Function
```

