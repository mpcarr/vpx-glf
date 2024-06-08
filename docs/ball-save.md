# Ball Save

A ball save prevents a drained ball from ending the player's turn. It automatically returns the ball to the playfield, giving the player another chance to continue their game.

## Example

```
Dim ball_save_base
Set ball_save_base = (new BallSave)("base", mode_base)

With ball_save_base
    .EnableEvents = Array("mode_base_started")
    .TimerStartEvents = Array("ball_device_plunger_eject_success")
    .ActiveTime = 15
    .HurryUpTime = 5
    .GracePeriod = 3
    .BallsToSave = -1
    .AutoLaunch = True
End With
```

In the above example we created a ball save called **ball_save_base** that belongs to the mode called **mode_base**. The ball save is enabled when the mode starts and the timer will start when the a ball is successfully ejected from the plunger device. A hurry up event will be emitted When the timer gets to 5 seconds remaining. The ball save has a grace period of 3 seconds which means it will stay enabled for 3 seconds after the timer has finished.

## Required Setings

### Name
```String```

The name of this device. Events emitted from the device will be in the format **name_ball_save**

### Mode
```String```

This is the mode the ball save belongs to. 

### Active Time
```Number```

### Hurry Up Time
```Number```

### Grace Period
```Number```

### Balls To Save
```Number```

### Auto Launch
```Boolean```

