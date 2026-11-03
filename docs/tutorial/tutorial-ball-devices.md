# 11. Optional scoop: capture and return a ball

**Prerequisite:** the core game is working. This optional device does not depend on the shots or timer extensions.

Add a capturing Kicker named `s_scoop` on the playfield and put it in `glf_switches`. Place it so a kick toward the upper playfield clears nearby walls. Keep its hit/unhit subs free for GLF's generated switch handlers.

Add these subs and append `ConfigureTutorialScoop` to `ConfigureGlfDevices`:

```vbscript
Sub ConfigureTutorialScoop()
    With CreateGlfBallDevice("tutorial_scoop")
        .BallSwitches = Array("s_scoop")
        .EjectTimeout = 2000
        .EjectCallback = "TutorialScoopEject"
    End With
End Sub

Sub TutorialScoopEject(ball)
    If IsObject(ball) Then s_scoop.Kick 0, 15
End Sub
```

Adjust the kick angle and strength to your scoop geometry. The callback returns the existing ball; it must not destroy and recreate it.

With no hold or lock claiming it, an unexpected ball is automatically ejected by the device after entry handling. `balldevice_tutorial_scoop_ball_enter` is the claim relay used by holds/locks; `balldevice_tutorial_scoop_ball_entered` is the notification event suitable for an award. The latter can be mapped to `tutorial_score_1000` inside the scoring mode's event player if you completed chapter 7.

## Checkpoint

A ball entering the scoop should settle, eject and continue playing without changing the current ball number. The device should recognize its exit. If the ball stays captured, check the callback, unclaimed-ball handling and physical kick path. A ball-save should not be needed for an ordinary scoop return.

To hold the ball during a selection or presentation, add a [ball hold](../ball-hold.md) with an explicit release event. A scoop does not automatically wait for an arbitrary award animation to finish.

[Previous](tutorial-timer.md) · [Tutorial map](index.md) · [Next](tutorial-autofire.md)
