# Magnet Configuration

The magnet configuration allows you to set up and customize magnets in your pinball machine. Magnets are used to grab, hold, and release balls with precise timing control.

## Configuration Options

### Basic Configuration
```vbscript
' Basic magnet configuration
With CreateGlfMagnet("magnet_name")
    .GrabSwitch = "s_magnet_switch"       ' Name of the switch that triggers ball grabbing
    .ActionCallback = "MagnetAction"      ' Name of the callback function for magnet actions
    .GrabTime = 1000                      ' Time in milliseconds to hold the ball
End With
```

### Advanced Configuration
```vbscript
' Advanced magnet configuration
With CreateGlfMagnet("magnet_name")
    ' Basic settings
    .GrabSwitch = "s_magnet_switch"
    .ActionCallback = "MagnetAction"
    .GrabTime = 1000

    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the magnet
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the magnet
    .GrabBallEvents = Array("event1", "event2")              ' Events that trigger ball grabbing
    .ReleaseBallEvents = Array("event3", "event4")           ' Events that trigger ball release
    .ResetEvents = Array("event5", "event6")                 ' Events that reset the magnet
    .FlingBallEvents = Array("event7", "event8")             ' Events that trigger ball flinging

    ' Timing settings
    .FlingDropTime = 250                                    ' Time to drop ball during fling
    .FlingRegrabTime = 50                                   ' Time to regrab after fling
    .ReleaseTime = 500                                      ' Time to release the ball

    ' Debug settings
    .Debug = True                                           ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `GrabSwitch`: String name of the switch that triggers ball grabbing (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the magnet state changes (Default: Empty)
- `GrabTime`: Integer specifying how long to hold the ball in milliseconds (Default: 1500)

### Event Control
- `EnableEvents`: Array of event names that will enable the magnet (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the magnet (Default: Array("ball_will_end", "service_mode_entered"))
- `GrabBallEvents`: Array of event names that will trigger ball grabbing (Default: Empty Array)
- `ReleaseBallEvents`: Array of event names that will trigger ball release (Default: Empty Array)
- `ResetEvents`: Array of event names that will reset the magnet (Default: Array("machine_reset_phase_3", "ball_starting"))
- `FlingBallEvents`: Array of event names that will trigger ball flinging (Default: Empty Array)

### Timing Settings
- `FlingDropTime`: Integer specifying how long to drop the ball during fling (Default: 250)
- `FlingRegrabTime`: Integer specifying how long to wait before regrabbing after fling (Default: 50)
- `ReleaseTime`: Integer specifying how long to take to release the ball (Default: 500)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Magnet Example
```vbscript
' Basic magnet configuration
With CreateGlfMagnet("mag1")
    .GrabSwitch = "s_TargetMystery5"
    .ReleaseBallEvents = Array("magnet_mag1_grabbed_ball")
    .GrabTime = 1000
    .ActionCallback = "GrabMagnetAction"
End With
```

### Advanced Magnet Example
```vbscript
' Advanced magnet configuration
With CreateGlfMagnet("mag2")
    .GrabSwitch = "s_magnet2"
    .ActionCallback = "Magnet2Action"
    .GrabTime = 1500
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .GrabBallEvents = Array("magnet_ready")
    .ReleaseBallEvents = Array("magnet_release")
    .FlingBallEvents = Array("magnet_fling")
    .FlingDropTime = 300
    .FlingRegrabTime = 100
    .ReleaseTime = 800
    .Debug = True
End With
```

## VPX Magnet Object

To create a VPX magnet object, use the following code:

```vbscript
' Create VPX magnet object
Dim MagnetName
Set MagnetName = New cvpmMagnet
With MagnetName
    .InitMagnet MagnetObject, 30           ' Initialize magnet with object and strength
    .GrabCenter = False                    ' Whether to grab at center of magnet
    .Strength = 15                         ' Magnet strength (higher = stronger pull)
    .CreateEvents "MagnetName"             ' Create events with this name prefix
End With
```

### VPX Magnet Properties
- `InitMagnet`: Initializes the magnet with a target object and strength
- `GrabCenter`: Boolean indicating whether to grab at the center of the magnet (Default: False)
- `Strength`: Integer specifying the magnet's strength (higher values = stronger pull)
- `MagnetOn`: Boolean property to control magnet state (true = on, false = off)

## Callback Examples

### Basic Magnet Callback
```vbscript
' Basic magnet callback
Sub MagnetAction(Enabled)
    MagnetName.MagnetOn = Enabled
End Sub
```

### Advanced Magnet Callback
```vbscript
' Advanced magnet callback with additional logic
Sub MagnetAction(Enabled)
    MagnetName.MagnetOn = Enabled
    
    If Enabled Then
        ' Magnet is being activated
        PlaySound "magnet_on"
        LightOn "l_magnet_active"
    Else
        ' Magnet is being deactivated
        PlaySound "magnet_off"
        LightOff "l_magnet_active"
    End If
End Sub
```

## Events

The magnet system generates several events that you can listen to:

- `{magnet_name}_grabbing_ball`: Triggered when the magnet starts grabbing a ball
- `{magnet_name}_grabbed_ball`: Triggered when the magnet has grabbed a ball
- `{magnet_name}_releasing_ball`: Triggered when the magnet starts releasing a ball
- `{magnet_name}_released_ball`: Triggered when the magnet has released a ball
- `{magnet_name}_flinging_ball`: Triggered when the magnet starts flinging a ball
- `{magnet_name}_flinged_ball`: Triggered when the magnet has completed flinging a ball

Replace `{magnet_name}` with your actual magnet name (prefixed with "magnet_" internally).

## Default Behavior

By default, magnets are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Reset on `machine_reset_phase_3` and `ball_starting` events
- Fling drop time of 250 milliseconds
- Fling regrab time of 50 milliseconds
- Grab time of 1500 milliseconds
- Release time of 500 milliseconds
- Debug logging disabled

## Notes

- The magnet system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for grab/activate, 0 for release/deactivate)
- Debug logging can be enabled to track magnet state changes
- Magnets can be configured to fling balls with precise timing control
- The system prevents multiple grab/release operations from overlapping
- VPX magnet objects provide additional control over magnet behavior and strength 