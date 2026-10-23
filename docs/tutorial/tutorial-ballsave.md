# 8. Save one early drain

**Prerequisite:** the core game serves and drains reliably, and `TutorialPlungerEject(ball)` from chapter 3 is present. Keep the ball count at one for this exercise.

Add this sub, then append `CreateTutorialBallSave` to the existing `ConfigureGlfDevices`:

```vbscript
Sub CreateTutorialBallSave()
    With CreateGlfMode("tutorial_save", 150)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .BallSaves("tutorial_save")
            .EnableEvents = Array("mode_tutorial_save_started")
            .TimerStartEvents = Array("balldevice_plunger_ball_eject_success")
            .ActiveTime = 10000
            .HurryUpTime = 2000
            .GracePeriod = 1000
            .BallsToSave = 1
            .AutoLaunch = True
        End With
    End With
End Sub
```

The mode enables the save when the new ball starts. Its timer begins when the plunger device confirms a successful exit. That confirmation follows the configured eject timeout, rather than the exact instant the plunger key is released. Durations here are milliseconds; the grace period extends the nominal active period.

`BallsToSave = 1` limits this save to one drain. `AutoLaunch = True` uses the plunger device's callback for the replacement ball. With only one physical ball, it must first return to the trough before it can be served again. GLF handles that wait; do not create another ball or manipulate the ball-in-play counter.

## Checkpoint

Launch and drain promptly. The replacement should be served and automatically launched without advancing the displayed player or ball number. Drain that replacement: this save has been used, so normal ball-end handling should occur. On the next ball, launch and keep it alive until the save and grace time have elapsed; a subsequent drain must advance to the next turn.

If the replacement sits at the plunger, check the eject callback and its lane direction. If a drain never advances, check that you did not set unlimited saves or restart the save on every hit. Repeat the chapter 6 full-game test with the new behavior understood.

Reference: [Ball Save](../ball-save.md).

[Previous](tutorial-event-player.md) · [Tutorial map](index.md) · [Next](tutorial-shots.md)
