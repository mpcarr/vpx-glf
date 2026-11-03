# 5. Add scoring and a visible display

**Goal:** make one physical rollover award exactly 1,000 points and show the current player, ball and score.

Add a Trigger named `s_score` somewhere a launched ball can cross it and put it in `glf_switches`. Use a rollover rather than a standup target for this first rule; specialized target integration comes later.

## Register a scoring mode

```vbscript
Sub CreateTutorialScoring()
    With CreateGlfMode("tutorial_scoring", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .VariablePlayer()
            With .EventName("s_score_active")
                With .Variable("score")
                    .Action = "add"
                    .Int = 1000
                End With
            End With
        End With
    End With
End Sub
```

Append `CreateTutorialScoring` to `ConfigureGlfDevices`. This mode starts every ball and stops before ball-end processing. Its variable player listens to the generated `s_score_active` event. GLF creates the built-in `score` variable for each new player; do not initialize or reset it on every ball.

## Show the score without external media

Create a visible backglass TextBox named `ScoreText` and a Timer named `TutorialDisplayTimer`, Enabled with Interval 100. Add its callback:

```vbscript
Sub TutorialDisplayTimer_Timer()
    If glf_gameStarted And Not IsNull(glf_currentPlayer) Then
        ScoreText.Text = "PLAYER " & CStr(GetPlayerState("number")) & _
            "   BALL " & CStr(GetPlayerState("ball")) & _
            "   SCORE " & CStr(GetPlayerState("score"))
    Else
        ScoreText.Text = "PRESS START"
    End If
End Sub
```

Test in desktop view with the backglass TextBox visible. This small UI timer only reads state; it does not run the game or score points. `Glf_GameTimer` must still run independently.

## Checkpoint

Start a new game and verify `PLAYER 1`, `BALL 1`, `SCORE 0`. Cross the rollover once: the score becomes 1,000. Cross again after leaving the trigger: 2,000. Drain and confirm the next ball retains the score. Do not leave the rollover inside the shooter-lane trigger or another overlapping sensor.

If a hit scores twice, look for duplicate hit subs or a second scoring rule. If there is no score, confirm collection membership and that `mode_tutorial_scoring_started` occurred. If points change but the text does not, check the display timer and backglass visibility separately.

We will replace this direct switch-to-score mapping with an event player in the optional [event-routing chapter](tutorial-event-player.md). For now, keep one award source. Reference: [Variable Player](../variable-player.md).

[Previous](tutorial-flippers.md) · [Tutorial map](index.md) · [Next](tutorial-lights.md)
