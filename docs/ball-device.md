# Ball Device

A ball device is anything on your table that can hold a ball and then release it.

## Scoop Example

```
Dim ball_device_scoop
Set ball_device_scoop = (new BallDevice)("scoop")

With ball_device_scoop
    .BallSwitches = Array("sw39")
    .EjectTimeout = 2
    .EjectCallback = "ScoopKickBall"
    .EjectAllEvents = Array("ball_ended")
End With
```

In the above example we created a ball device called **ball_device_scoop**, it has one switch so holds one ball and an eject timeout of 2 seconds. When the ball ends, if there are any balls in this device, it will eject them as we set the eject all events to ball_ended.

## Plunger Example

```
Dim ball_device_scoop
Set ball_device_scoop = (new BallDevice)("plunger")

With ball_device_scoop
    .BallSwitches = Array("sw10")
    .EjectTimeout = 2
    .EjectCallback = "AutoPlunge"
    .MechcanicalEject = True
    .DefaultDevice = True
End With
```

In the above example we created a plunger ball device, This plunger has a real plunger so we set the mechcanical eject to true and also set the default device property to true as this is the device that is fed from the trough. This device also has a auto plunge so we set a callback of *AutoPlunge* so that sub will be called when the ball is auto plunged.

## Required Setings

### Name
```String```

The name of this device. Events emitted from the device will be in the format **name_ball_device**

### Ball Switches
```Array```

This is an array of switch / trigger names used to monitor what balls are in the device.

### Eject Timeout
```Number```

A value in seconds that the ball device will wait until it deems the ball has been ejected

### Eject Callback
```String```

The sub or function to call when the ball is ejected. The ball device doesn't actually do anything to the ball, this sub must **kick** the ball
out of the device.

## Optional Settings

### Default Device
```Boolean```

The default device is the plunger. Set this to True for the device that the trough ejects to.

### Eject All Events
```Array```

A list of event names that will cause the ball device to eject all it's balls. e.g. **ball_ended**

### Player Controlled Eject Events
```Array```

A list of event names that will cause the ball device to eject one ball.

### Mechcanical Eject
```Boolean```

If this device doesn't have an eject callback. e.g. a plunger moves the ball.