# 12. Optional autofire: a simple kickback

**Goal:** show how an autofire device connects a switch event to a physical action. This example uses a kicker so the callback can be complete without importing a sling animation or physics system.

Add a capturing Kicker named `s_kickback` in a return lane where its kick will send the ball back onto the playfield. Add it to `glf_switches`. Do not also register it as a GLF ball device: this example uses immediate autofire behavior.

Add these subs and append `ConfigureTutorialKickback` to `ConfigureGlfDevices`:

```vbscript
Sub ConfigureTutorialKickback()
    With CreateGlfAutoFireDevice("tutorial_kickback")
        .Switch = "s_kickback"
        .ActionCallback = "TutorialKickbackAction"
    End With
End Sub

Sub TutorialKickbackAction(args)
    If args(0) = 1 And s_kickback.BallCntOver > 0 Then
        s_kickback.Kick 0, 20
    End If
End Sub
```

Autofire action callbacks receive an array: `args(0)` is the enabled/action state and `args(1)` is the active ball, when available. This differs from the flipper callback's Boolean and ball-device eject callback's ball argument. The callback above acts only on activation and checks physical occupancy.

Defaults enable the device on `ball_started` and disable it on `ball_will_end` / service entry. The kick direction and speed must match your return lane. This basic exercise does not implement kickback qualification or a one-use limit.

## Checkpoint

During a ball, send the ball into the kicker and confirm it returns to play once, without changing the player/ball number. Make sure its unhit event does not kick again. At ball end the autofire device should disable.

Slingshots and bumpers use the same GLF device type, but their callbacks must drive their own VPX mechanics/animations. Add real slingshots to `glf_slingshots`; do not copy undefined `LeftSlingshotAction` or enable/disable callback names from another table. For the full pattern, inspect the example table's `_configuration.vbs` and matching VPX physics files.

Your next steps are [a complete game mode](../guides/complete-mode.md), [multiball](../guides/multiball.md), [tilt](../tilt.md), [sound](../sound-system.md), [displays](../guides/displays.md) and [high scores](../high-score.md). Keep the core acceptance test working as each feature is added.

[Previous](tutorial-ball-devices.md) · [Tutorial map](index.md)
