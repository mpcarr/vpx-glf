# Tutorial - Ball Save

Lets add a ball save to our table. A ball save is a logic block that monitors if the balls drains while it is active. If it does, the ball will not end and a new ball will be released from the trough.

### Configuration

To configure your ball save we need a mode for it live under. Create the following config to create a mode and a ball save logic block

#### Ball Save Config

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        With BallSaves("base")
            .EnableEvents = Array("mode_base_started")
            .TimerStartEvents = Array("balldevice_plunger_ball_eject_success")
            .ActiveTime = 15000
            .HurryUpTime = 5000
            .GracePeriod = 3000
            .BallsToSave = -1
            .AutoLaunch = True
        End With
    End With
End Sub
```

As you can see from the settings above the ball save will be active for 15 seconds once the ball has successfully been ejected from the plunger lane, it has a grace period of 3 seconds meaning that it will actually be active for 18 seconds in total and it will save an unlimited about of balls as long as it is active.

## Enabling the Mode

Finally we need to enable this mode in our game by calling the ```CreateBaseMode``` sub in our ```CongifureGlfDevices``` Sub.

```
Sub ConfigureGlfDevices
   
    CreateBaseMode() '<---Add this to enable the mode
    
    'Other device config....
End Sub
```