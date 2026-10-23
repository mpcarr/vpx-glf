# 10. Add a timed award opportunity

**Goal:** crossing the original scoring rollover opens a 15-second window; crossing a new goal during that window awards 10,000.

Add Trigger `s_goal` to `glf_switches` and Light `l_window` to `glf_lights`. Add the sub below and append `CreateTutorialHurry` to `ConfigureGlfDevices`. This works alongside the existing scoring mode: `s_score` still gives its ordinary 1,000 points.

```vbscript
Sub CreateTutorialHurry()
    With CreateGlfMode("tutorial_hurry", 300)
        .StartEvents = Array("s_score_active{modes.tutorial_hurry.active == False}")
        .StopEvents = Array("tutorial_hurry_finish", "timer_tutorial_hurry_clock_complete", "ball_will_end", "tilt")
        With .Timers("tutorial_hurry_clock")
            .StartValue = 0
            .EndValue = 15
            .Direction = "up"
            .TickInterval = 1000
            With .ControlEvents()
                .EventName = "mode_tutorial_hurry_started"
                .Action = "restart"
            End With
        End With
        With .LightPlayer()
            With .EventName("mode_tutorial_hurry_started")
                With .Lights("l_window")
                    .Color = "00ff00"
                End With
            End With
        End With
        With .VariablePlayer()
            With .EventName("s_goal_active.1")
                With .Variable("score")
                    .Action = "add"
                    .Int = 10000
                End With
            End With
        End With
        With .EventPlayer()
            .Add "s_goal_active", Array("tutorial_hurry_finish")
        End With
    End With
End Sub
```

The start condition prevents repeated rollover hits from restarting an already active mode. Its timer is named `tutorial_hurry_clock`, so its completion event is exactly `timer_tutorial_hurry_clock_complete`. The configuration API is `.Timers`, and the one-second tick interval makes 15 ticks a 15-second opportunity.

The goal awards points before posting a separate finish event. The `.1` suffix raises that scoring listener's priority; it is not a time delay. Stopping the mode removes the light entry and timer activity. The goal is not scored outside the mode because its variable player is inactive.

## Checkpoint

Before opening a window, `s_goal` gives no hurry-up award. Hit `s_score`: the green window light turns on. Hit `s_score` again during the window: it still gives the ordinary rollover score but does not extend the window. Hit `s_goal` in time: +10,000 and the light turns off. Try again and wait: expiry turns the light off with no hurry-up award. Drain during the window and verify it is gone on the next ball.

Timer control values have action-specific meanings. For example, `add` uses ticks, while `pause` uses a delay. Consult the [Timer API](../timer.md) before adding controls; do not pass 5,000 expecting an addition of five ticks.

[Previous](tutorial-shots.md) · [Tutorial map](index.md) · [Next](tutorial-ball-devices.md)
