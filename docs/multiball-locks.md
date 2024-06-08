# Multiball Locks

This command tracks balls locked for a multiball. It is used with the multiball command to start multiballs.

## Example

```
Dim waterfall_mb_locks
Set waterfall_mb_locks = (new MultiballLocks)("waterfall", mode_waterfall_mb) 

With waterfall_mb_locks
    .EnableEvents = Array("enable_waterfall")
    .DisableEvents = Array("disable_waterfall")
    .BallsToLock = 3
    .LockEvents = Array("balldevice_bd_waterfall_vuk_ball_eject_success")
    .ResetEvents = Array("multiball_waterfall_started")
End With
```

In the above example we defined a multiball lock for the waterfall multiball. This is a three ball multiball. When the lock is enabled and the ball device waterfall_vuk successfully ejects a ball, it will count that ball as locked. This could be a physical table locked ball, or a virutal locked ball.