# Ball Device Configuration

The ball device configuration allows you to set up and customize ball devices in your pinball machine. Ball devices are used for scoops, physical ball locks, staging areas. Anywhere you want to hold a ball on the table. 

## Configuration Options

### Basic Configuration
```vbscript
' Basic ball device configuration
With CreateGlfBallDevice("scoop")
    .BallSwitches = Array("s_scoop1")
    .EjectTimemout = 2000
    .EjectCallback = "ScoopEjectCallback"              ' Name of the callback function for ball device
End With
```

Each device will have a ```Eject Callback``` setting. This calls your defined function / sub when the ball should be ejected from the device. In a real pinball machine this would be the coil firing to eject the ball. In VPX, you need to implement this function to virutally *kick* the ball out from the device. Below is an example of how to you could do that.

```
Sub ScoopEjectCallback(ball)
	Dim ang, vel
	If s_scoop1.BallCntOver > 0 Then 'Check if the ball is positioned on the switch / kicker
		ang = 14.8 + ScoopAngleTol*2*(rnd-0.5) 'generate some randomness to the angle
		vel = 70.0 + ScoopStrengthTol*2*(rnd-0.5) 'generare some randomness to the velocity
		KickBall ball, ang, vel, 0, 0 'actually kick the ball
		SoundSaucerKick 1, s_Scoop 'play a sound to simulate the coil firing
	Else
		SoundSaucerKick 0, s_Scoop 'play a sound if even if the ball wasn't on the switch (failed kick)
	End If
	DOF 109, DOFPulse 'Optionally call a DOF event.
End Sub
```

## Property Descriptions

### Basic Settings
- `BallSwitches`: Array of switch names that detect balls in the device (Default: Empty Array)
- `EjectCallback`: String name of the callback function that executes when the ball device ejects a ball (Default: Null)

### Ejection Control
- `EjectTimeout`: Integer specifying the timeout in milliseconds for ejection (Default: 1000)
- `EjectEnableTime`: Integer specifying the enable time in milliseconds for ejection (Default: 0).
- `MechanicalEject`: Boolean indicating whether the device uses mechanical ejection (Default: False)
- `DefaultDevice`: Boolean indicating whether this is the default ball device (Default: False)
- `PlayerControlledEjectEvent`: String name of the event that triggers player-controlled ejection (Default: Empty)

### Event Control
- `EjectAllEvents`: Array of event names that will trigger ejection of all balls (Default: Empty Array)
- `EjectTargets`: Array of target names that will trigger ejection timeout when activated (Default: Empty Array)

### Ball Tracking
- `EntranceCountDelay`: Integer specifying the delay in milliseconds before counting a ball as entered (Default: 500)
- `ExcludeFromBallSearch`: Boolean indicating whether to exclude this device from ball search operations (Default: False)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Events

The ball device system generates several events that you can listen to:

- `balldevice_{device_name}_ball_enter`: Relay event triggered when a ball is entering the device (can be claimed by other systems)
- `balldevice_{device_name}_ball_entered`: Triggered when a ball has successfully entered the device
- `balldevice_{device_name}_ball_exiting`: Triggered when a ball is exiting the device
- `balldevice_{device_name}_ball_eject_success`: Triggered when a ball has been successfully ejected from the device

Replace `{device_name}` with your actual ball device name.



## Notes

- The ball device system automatically manages ball tracking and ejection based on configured switches
- Eject callbacks receive a ball parameter (can be Null for enable time callbacks)
- Debug logging can be enabled to track ball device state changes
- Ball devices can be configured to eject balls during specific game modes using EjectAllEvents
- The system prevents multiple ejection operations from overlapping
- Ball tracking is always enabled and managed automatically
- The entrance count delay helps prevent false ball detection
- Mechanical ejection mode uses timeouts to track successful ball exits
- Player controlled ejection allows manual triggering via events
- Eject targets can be used to trigger ejection timeouts when specific targets are hit 