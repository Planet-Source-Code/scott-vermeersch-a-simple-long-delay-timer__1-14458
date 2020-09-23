<div align="center">

## A Simple Long Delay Timer


</div>

### Description

A simple long delay timer. The VB Timer control is limited to 64k milli seconds, this short code extends that to near infinity in 6 lines of code. Heavily commented.
 
### More Info
 
Put a VB Timer control on a form. This assumes the control with a name of "Timer1".


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scott Vermeersch](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scott-vermeersch.md)
**Level**          |Beginner
**User Rating**    |4.8 (38 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scott-vermeersch-a-simple-long-delay-timer__1-14458/archive/master.zip)





### Source Code

```
Private Sub Timer1_Timer()
  'Set the interval of the timer.
  'Each time the timer event is called
  'subtract one from the tick.
  'Assuming 10000 interval with a 1440 tick,
  'this would give a delay of 4 hours.
  '10000 / 1000 = 10 *Milliseconds to Seconds
  '10 * 1440 = 14400 *Multiply by ticks
  '14400 / 60 = 240  *Seconds to minutes
  '240 / 60 = 4    *Minutes to hours
  'Allocate a static variable
  Static siTick As Integer
  'Check if variable greater then desired tick
  If siTick > 1440 Then
    '************ Do some code ************'
    'Start the ticker over
    siTick = 0
  'If not then add one to the tick
  Else
    siTick = siTick + 1
  End If
End Sub
```

