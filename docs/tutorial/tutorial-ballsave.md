# Tutorial - Ball Save

Lets add a ball save to our table. A ball save is a logic block that monitors if the balls drains while it is active. If it does, the ball will not end and a new ball will be released from the trough.

### Download the VPX file
[Tutorial - Ball Save VPX File](https://github.com/mpcarr/vpx-glf/raw/main/tutorial/glf_tutorial_ballsave.vpx)

### Configuration

To configure your ball save we need a mode for it live under. Create the following config to create a mode and a ball save logic block

#### Mode Config

```
Dim mode_ballsave
Set mode_ballsave = (new Mode)("ballsave", 1000)

With mode_ballsave
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
End With
```

The mode above will start every time a new ball is started and it will end when a ball ends.

#### Ball Save Config

```
Dim ball_save
Set ball_save = (new BallSave)("base", mode_ballsave)

With ball_save
    .EnableEvents = Array("mode_ballsave_started")
    .TimerStartEvents = Array("balldevice_plunger_ball_eject_success")
    .ActiveTime = 15
    .HurryUpTime = 5
    .GracePeriod = 3
    .BallsToSave = -1
    .AutoLaunch = True
End With
```

As you can see from the settings above the ball save will be active for 15 seconds once the ball has successfully been ejected from the plunger lane, it has a grace period of 3 seconds meaning that it will actually be active for 18 seconds in total and it will save an unlimited about of balls as long as it is active.

#### Update Plunger config

We have configured our ball save to auto launch the new ball. In the previous tutorial we only set out plunger as a manual plunger. We need to update the config to support auto launching a ball.

```
With ball_device_plunger
    .BallSwitches = Array("sw01")
	.EjectTargets = Array("sw02")
    .EjectStrength = 100
    .MechcanicalEject = True
    .DefaultDevice = True
End With
```

Above we have also changed the **EjectTimeout** for **EjectTargets**. As we have a gate at the end of our plunger lane, we can use this to determine if our ball made it out of the plunger lane successfully rather than a timeout.

