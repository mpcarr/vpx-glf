# Tutorial - Flippers

[See Flipper Config Reference](../../flipper)

## Introduction
Flippers are a critical component of any pinball machine. This tutorial will guide you through the setup, configuration, and advanced customization of flippers in your system.

## **Configuration**
To configure your flippers, add the following code to the `ConfigureGlfDevices` subroutine:

```
Sub ConfigureGlfDevices
  
    ' Flippers
    With CreateGlfFlipper("left")
        .Switch = "s_left_flipper"
        .ActionCallback = "LeftFlipperAction"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started", "enable_flippers")
    End With

    With CreateGlfFlipper("right")
        .Switch = "s_right_flipper"
        .ActionCallback = "RightFlipperAction"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started", "enable_flippers")
    End With

End Sub
```

Your action callback could look like this:

```
Sub LeftFlipperAction(Enabled)
	If Enabled Then
		LeftFlipper.RotateToEnd
	Else
		LeftFlipper.RotateToStart
	End If
End Sub
```
