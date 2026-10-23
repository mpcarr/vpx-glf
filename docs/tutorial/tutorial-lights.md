# 6. Add GI and finish a complete game

**Goal:** light the playfield during a ball and verify the complete game loop before adding more features.

Add a Light named `l_gi`, initially off, and include it in `glf_lights`. Set its BlinkPattern field to `GI`. GLF uses this field for comma-separated tags; additional GI lights can share that tag. Keep the tag exactly consistent with the rule.

![GI controlled by a mode](../images/tutorial-lights.gif)

## Add the GI mode

```vbscript
Sub CreateTutorialLights()
    With CreateGlfMode("tutorial_lights", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .LightPlayer()
            With .EventName("mode_tutorial_lights_started")
                With .Lights("GI")
                    .Color = "ffffff"
                End With
            End With
        End With
    End With
End Sub
```

Append `CreateTutorialLights` to `ConfigureGlfDevices`. The complete core configuration entry point is now:

```vbscript
Sub ConfigureGlfDevices()
    ConfigureTutorialPlunger
    ConfigureTutorialFlippers
    CreateTutorialScoring
    CreateTutorialLights
End Sub
```

The light player responds to its own mode's started event, when its listeners are ready. Stopping the mode removes its light entry; with no other light rule active, the GI returns to its off baseline.

## Acceptance test: one complete game

1. Start the table: one physical ball is in the trough, GI is off and the display says `PRESS START`.
2. Press and release Start: a ball is served, GI turns on and the display shows player 1, ball 1, score 0.
3. Launch, flip and cross `s_score` twice: score becomes 2,000.
4. Drain: the ball returns through the trough; ball 2 is served and the score remains 2,000.
5. Drain balls 2 and 3. With the table's `Balls Per Game` option set to `3 Balls`, no fourth ball is served. GI turns off and `PRESS START` returns.
6. Start a fresh game: score and ball number reset. Confirm there is still only one physical ball.
7. Start another game, add player 2 with Start during ball 1, and verify separate scores and alternating turns after drains.

The table option can select five balls, so check it if your game continues after the third drain. Physical ball count (`tnob`) is independent of turns per game.

Compare your script with the [completed core script](../assets/tutorial/minimal-game.vbs). It includes chapters 1–6 only. Do not append it below your existing script; use it to compare or replace the whole tutorial script after saving a copy.

You now have the core loop. Add features only after these checks pass; otherwise debugging a new rule becomes mixed up with physical ball-routing problems.

[Previous](tutorial-scoring.md) · [Tutorial map](index.md) · [Next](tutorial-event-player.md)
