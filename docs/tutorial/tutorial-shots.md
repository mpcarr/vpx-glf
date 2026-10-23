# 9. Add a shot with visible progress

**Goal:** one new target lights an insert and awards 5,000 points only on its first hit each ball.

Create a rollover Trigger called `s_target` and a Light called `l_target` beside it. Add the trigger to `glf_switches` and the light to `glf_lights`. Keep this separate from the `s_score` rollover so the two awards are easy to distinguish.

Add this sub and append `CreateTutorialTarget` to `ConfigureGlfDevices`:

```vbscript
Sub CreateTutorialTarget()
    With CreateGlfMode("tutorial_target_mode", 200)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .ShotProfiles("tutorial_target_profile")
            .AdvanceOnHit = True
            With .States("unlit")
                .Show = "off"
            End With
            With .States("lit")
                .Show = "on"
            End With
        End With
        With .Shots("tutorial_target")
            .Switch = "s_target"
            .Profile = "tutorial_target_profile"
            .Persist = False
            With .Tokens()
                .Add "lights", "l_target"
            End With
        End With
        With .VariablePlayer()
            With .EventName("tutorial_target_unlit_hit")
                With .Variable("score")
                    .Action = "add"
                    .Int = 5000
                End With
            End With
        End With
    End With
End Sub
```

The profile defines states in order: index 0 is `unlit`, index 1 is `lit`. Profiles do not loop in this implementation; `ProfileLoop` is read-only, so do not assign it. GLF provides the `off` and `on` shows, and the shot supplies their required `lights` token. The API is `.Shots("name")`, not `.Shot`.

On a hit, GLF advances the profile and emits events describing the state **before** the hit. Therefore `tutorial_target_unlit_hit` awards completion on the first hit, while the insert changes to lit. Further hits emit `tutorial_target_lit_hit` and do not match this award. `Persist = False` lets this standalone tutorial shot start fresh on its next activation.

`HitEvents` is an input that tells a shot to register a hit from another event; it is not a list of awards to emit. Put scoring in a variable player responding to the shot's emitted events.

## Checkpoint

At ball start the insert is off. First target pass: insert on and +5,000. Second pass: no additional target award. Drain and start another ball: insert off and the first target hit can earn the award again. Confirm player 2 starts unlit as well.

Once this works, continue with [shot groups](../shot-group.md) for sets of objectives and [sequence shots](../sequence-shot.md) for ramps/orbits that require ordered switches. Those need extra playfield objects and are not prerequisites for the core game.

[Previous](tutorial-ballsave.md) · [Tutorial map](index.md) · [Next](tutorial-timer.md)
