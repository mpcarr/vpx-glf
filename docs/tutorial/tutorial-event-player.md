# 7. Separate switch events from awards

**Prerequisite:** finish the core game through chapter 6. This chapter changes the existing scoring sub; it does not create a second scoring mode.

Replace `CreateTutorialScoring` with the following version. Keep its existing call in `ConfigureGlfDevices` exactly once.

```vbscript
Sub CreateTutorialScoring()
    With CreateGlfMode("tutorial_scoring", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .EventPlayer()
            .Add "s_score_active", Array("tutorial_score_1000")
        End With
        With .VariablePlayer()
            With .EventName("tutorial_score_1000")
                With .Variable("score")
                    .Action = "add"
                    .Int = 1000
                End With
            End With
        End With
    End With
End Sub
```

Now the path is `s_score_active` → `tutorial_score_1000` → score increment. Other devices can post `tutorial_score_1000` without repeating the scoring rule. Later, another player can also listen to the same award event for sound or display feedback.

Remove the old `s_score_active` variable-player entry when applying this replacement. Keeping both gives 2,000 points per rollover. Do not create another `CreateTutorialScoring` sub or add the same event key twice to one event player.

## Checkpoint

A single rollover still awards exactly 1,000. Drain, start another ball and check the route still works. With debug logging enabled, follow the switch event and then the award event. See [Events](../events.md) for conditions, event arguments and priority offsets after this simple path works.

[Previous](tutorial-lights.md) · [Tutorial map](index.md) · [Next](tutorial-ballsave.md)
