# Build a complete timed mode

This walkthrough combines qualification, a timer, progress, scoring and cleanup. It follows the component patterns used by the example table and Dark Chaos's double-scoring feature, using a small new rule: hit one target three times to start a 20-second challenge, then hit a second target five times to earn 100,000 points.

## Prerequisites

Use a working GLF table with game lifecycle and switch events already connected. Replace `s_qualify_active` and `s_goal_active` below with two actual switch events. No custom lights, sounds, player variables or media assets are required by this example.

Add the following sub to your game source and call `CreateChallengeModes` once from your configuration entry point before `Glf_Init`.

```vbscript
Sub CreateChallengeModes
    With CreateGlfMode("challenge_qualify", 300)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("challenge_qualified", "ball_will_end", "tilt")

        With .Counters("challenge_qualifier")
            .EnableEvents = Array("mode_challenge_qualify_started")
            .CountEvents = Array("s_qualify_active")
            .CountCompleteValue = 3
            .DisableOnComplete = True
            .ResetOnComplete = False
            .PersistState = False
            .EventsWhenComplete = Array("challenge_qualified")
        End With
    End With

    With CreateGlfMode("challenge", 400)
        .StartEvents = Array("challenge_qualified")
        .StopEvents = Array("challenge_finish", "challenge_timeout", "ball_will_end", "tilt")

        With .Counters("challenge_hits")
            .EnableEvents = Array("mode_challenge_started")
            .CountEvents = Array("s_goal_active")
            .CountCompleteValue = 5
            .DisableOnComplete = True
            .ResetOnComplete = False
            .PersistState = False
            .EventsWhenComplete = Array("challenge_complete")
        End With

        With .Timers("challenge_clock")
            .StartValue = 0
            .EndValue = 20
            .Direction = "up"
            .TickInterval = 1000
            With .ControlEvents()
                .EventName = "mode_challenge_started"
                .Action = "restart"
            End With
        End With

        With .VariablePlayer()
            With .EventName("challenge_complete.1")
                With .Variable("score")
                    .Action = "add"
                    .Int = 100000
                End With
            End With
        End With

        With .EventPlayer()
            .Add "challenge_complete", Array("challenge_finish")
            .Add "timer_challenge_clock_complete", Array("challenge_timeout")
        End With
    End With
End Sub
```

## Follow the event chain

| Trigger | Result |
| --- | --- |
| `ball_started` | Qualification starts with a fresh counter. |
| Third qualifying hit | Counter posts `challenge_qualified`; qualification stops and challenge starts. |
| `mode_challenge_started` | Enables the goal counter and restarts the timer. |
| Fifth goal hit | Posts `challenge_complete`; variable player awards points, then the event player posts `challenge_finish`. |
| Timer completion | Posts `challenge_timeout`; the mode stops without a completion award. |
| Ball ending or tilt | Both modes stop. |

The `.1` suffix raises that scoring listener's priority relative to the default listener in the same mode. It is not a delay. `challenge_finish` is deliberately separate from `challenge_complete`: award listeners get to process completion before mode teardown removes them.

Counter enable events are explicit because this counter implementation does not automatically begin counting on activation. `PersistState = False` resets progress when the mode is activated again. Qualification starts again on the next `ball_started`, so this version offers one challenge attempt per ball, including an extra ball.

## Add feedback and adapt the rule

Add a [show player](../show-player.md) on `mode_challenge_started` and a completion effect on `challenge_complete`. Register shows before initialization. For displays, use [segment displays](../segment-display-player.md) or [Godot slides](displays.md). Give every repeating effect a defined stop or mode-cleanup path.

For progress that survives between balls, use [player variables](../player-state.md) or a persistent shot/state-machine design. Do not simply add more start events: decide which counters should reset and when qualification becomes available again.

## Test the whole feature

Verify three qualifying hits start exactly one challenge; four goal hits do not award completion; the fifth awards exactly 100,000; timeout awards nothing; tilt and drain stop the timer; and the next player's qualification begins at zero. Also try a goal hit close to the timeout: event order determines which action is processed while the mode is still active.

Source: `scripts/src/commands/counter.vbs`, `timer.vbs`, `mode.vbs`, `event-player.vbs` and `variable-player.vbs`. For a larger feature, compare Dark Chaos's `scripts/src/game/modes/double-scoring.vbs` and the example table's `scripts/src/game/modes/rampshot.vbs`.
