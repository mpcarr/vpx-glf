# Ball Device Configuration

The ball device configuration allows you to set up and customize ball devices in your pinball machine. Ball devices are used to manage balls, including tracking their location, ejecting them, and handling ball-related events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic ball device configuration
With CreateGlfBallDevice("device_name")
    .BallSwitches = Array("s_switch1", "s_switch2")  ' Array of switches that detect balls
    .EjectCoil = "c_eject_coil"                      ' Name of the coil to eject balls
    .ActionCallback = "BallDeviceAction"              ' Name of the callback function for ball device actions
End With
```

### Advanced Configuration
```vbscript
' Advanced ball device configuration
With CreateGlfBallDevice("device_name")
    ' Basic settings
    .BallSwitches = Array("s_switch1", "s_switch2")
    .EjectCoil = "c_eject_coil"
    .ActionCallback = "BallDeviceAction"
    
    ' Ejection control
    .EjectTimeout = 5000                              ' Timeout in milliseconds for ejection
    .MechanicalEject = False                          ' Whether the device uses mechanical ejection
    .DefaultDevice = "default_device"                 ' Default device to send balls to
    
    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the ball device
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the ball device
    .EjectAllEvents = Array("event1", "event2")              ' Events that trigger ejection of all balls
    .EjectCallback = "EjectCallback"                         ' Callback for ejection events
    
    ' Ball tracking
    .TrackBalls = True                                ' Whether to track balls in this device
    .BallCount = 2                                    ' Maximum number of balls this device can hold
    
    ' Debug settings
    .Debug = True                                     ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `BallSwitches`: Array of switch names that detect balls in the device (Default: Empty Array)
- `EjectCoil`: String name of the coil to eject balls (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the ball device state changes (Default: Empty)

### Ejection Control
- `EjectTimeout`: Integer specifying the timeout in milliseconds for ejection (Default: 5000)
- `MechanicalEject`: Boolean indicating whether the device uses mechanical ejection (Default: False)
- `DefaultDevice`: String name of the default device to send balls to (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the ball device (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the ball device (Default: Array("ball_will_end", "service_mode_entered"))
- `EjectAllEvents`: Array of event names that will trigger ejection of all balls (Default: Empty Array)
- `EjectCallback`: String name of the callback function for ejection events (Default: Empty)

### Ball Tracking
- `TrackBalls`: Boolean indicating whether to track balls in this device (Default: True)
- `BallCount`: Integer specifying the maximum number of balls this device can hold (Default: 1)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Ball Device Example
```vbscript
' Basic ball device configuration
With CreateGlfBallDevice("plunger")
    .BallSwitches = Array("s_plunger")
    .EjectCoil = "c_plunger"
    .ActionCallback = "PlungerAction"
End With
```

### Advanced Ball Device Example
```vbscript
' Advanced ball device configuration
With CreateGlfBallDevice("scoop")
    .BallSwitches = Array("s_scoop_entry", "s_scoop_exit")
    .EjectCoil = "c_scoop_eject"
    .ActionCallback = "ScoopAction"
    .EjectTimeout = 3000
    .MechanicalEject = False
    .DefaultDevice = "playfield"
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .EjectAllEvents = Array("scoop_eject_all")
    .EjectCallback = "ScoopEjectCallback"
    .TrackBalls = True
    .BallCount = 3
    .Debug = True
End With
```

## Callback Examples

### Basic Ball Device Callback
```vbscript
' Basic ball device callback
Sub BallDeviceAction(State)
    ' Code to handle ball device action
    If State = 1 Then
        ' Ball device is being activated
        PlaySound "device_activate"
    Else
        ' Ball device is being deactivated
        PlaySound "device_deactivate"
    End If
End Sub
```

### Advanced Ball Device Callback
```vbscript
' Advanced ball device callback with additional logic
Sub BallDeviceAction(State)
    If State = 1 Then
        ' Ball device is being activated
        PlaySound "device_activate"
        LightOn "l_device_active"
        
        ' Add score
        AddScore 1000
    Else
        ' Ball device is being deactivated
        PlaySound "device_deactivate"
        LightOff "l_device_active"
    End If
End Sub
```

## Events

The ball device system generates several events that you can listen to:

- `{device_name}_ball_added`: Triggered when a ball is added to the device
- `{device_name}_ball_removed`: Triggered when a ball is removed from the device
- `{device_name}_ejecting`: Triggered when the device is ejecting a ball
- `{device_name}_ejected`: Triggered when the device has ejected a ball
- `{device_name}_enabled`: Triggered when the device is enabled
- `{device_name}_disabled`: Triggered when the device is disabled

Replace `{device_name}` with your actual ball device name (prefixed with "ball_device_" internally).

## Default Behavior

By default, ball devices are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Eject timeout of 5000 milliseconds
- Mechanical ejection disabled
- Ball tracking enabled
- Maximum ball count of 1
- Debug logging disabled

## Notes

- The ball device system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for activate, 0 for deactivate)
- Debug logging can be enabled to track ball device state changes
- Ball devices can be configured to eject balls during specific game modes
- The system prevents multiple ejection operations from overlapping
- Ball tracking can be enabled or disabled based on your needs
- The maximum ball count can be adjusted to accommodate multiple balls 