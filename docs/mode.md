# Mode

All of the game logic commands need to be a member of a mode.

## Example  

```
Dim mode_super_pops
Set mode_super_pops = (new Mode)("super_pops", 2000)

With mode_super_pops
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
End With
```

In the example, we created a super pops mode with a priorty of 2000. The mode will start when the ball_started event emits and it will end when the ball_ended event emits.
This mode by it self doesn't do much, you need to define other commands for this mode like counters, timers, shots e.t.c. All commands added to this mode will run at the same priority level.