# 4. Make both flippers playable

**Goal:** control both flippers during an active ball and release them at ball end.

Name the VPX objects `LeftFlipper` and `RightFlipper`, position them above the drain, and mirror their start/end angles so they flip toward the playfield. GLF supplies input events named `s_left_flipper` and `s_right_flipper`; these are keyboard-derived switches, not extra VPX Trigger objects.

## Add both callbacks

```vbscript
Sub ConfigureTutorialFlippers()
    With CreateGlfFlipper("left")
        .Switch = "s_left_flipper"
        .ActionCallback = "TutorialLeftFlipper"
    End With
    With CreateGlfFlipper("right")
        .Switch = "s_right_flipper"
        .ActionCallback = "TutorialRightFlipper"
    End With
End Sub

Sub TutorialLeftFlipper(enabled)
    If enabled Then
        LeftFlipper.RotateToEnd
    Else
        LeftFlipper.RotateToStart
    End If
End Sub

Sub TutorialRightFlipper(enabled)
    If enabled Then
        RightFlipper.RotateToEnd
    Else
        RightFlipper.RotateToStart
    End If
End Sub
```

Append `ConfigureTutorialFlippers` after `ConfigureTutorialPlunger` inside the existing `ConfigureGlfDevices`. Keep the same two key handlers from chapter 1. Do not add direct flipper rotation there as well.

The flipper device defaults enable it on `ball_started` and disable it on `ball_will_end` or `service_mode_entered`. Keeping those defaults avoids relying on a custom `kill_flippers` event that this game never posts. These simple callbacks demonstrate GLF wiring; advanced flipper physics can replace their internals later.

## Checkpoint

Before Start, the flippers should remain at rest. Start and launch a ball: both buttons should move their respective flippers and release correctly. Drain while holding a button and confirm both return to rest. Test another ball to confirm they enable again. Correct mirrored angles or collision geometry in the editor if a callback runs but the ball cannot be played.

Reference: [Flipper API](../flipper.md).

[Previous](tutorial-plunger.md) · [Tutorial map](index.md) · [Next](tutorial-scoring.md)
