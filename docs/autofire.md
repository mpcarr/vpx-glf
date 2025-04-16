# Auto-Fire Configuration

The auto-fire configuration allows you to set up and customize auto-firing devices in your pinball machine. Auto-fire devices are used to automatically fire coils or other devices when triggered by switches or events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic auto-fire configuration
With CreateGlfAutoFire("autofire_name")
    .Switch = "s_autofire_switch"       ' Name of the switch that triggers the auto-fire
    .Coil = "c_autofire_coil"           ' Name of the coil to fire
    .ActionCallback = "AutoFireAction"  ' Name of the callback function for auto-fire actions
End With
```

### Advanced Configuration
```vbscript
' Advanced auto-fire configuration
With CreateGlfAutoFire("autofire_name")
    ' Basic settings
    .Switch = "s_autofire_switch"
    .Coil = "c_autofire_coil"
    .ActionCallback = "AutoFireAction"
    
    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the auto-fire
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the auto-fire
    .FireEvents = Array("event1", "event2")                  ' Events that trigger firing
    .ResetEvents = Array("event3", "event4")                 ' Events that reset the auto-fire
    
    ' Ball search settings
    .BallSearchEvents = Array("ball_search_started")         ' Events that trigger ball search
    .BallSearchDelay = 500                                   ' Delay before ball search firing
    
    ' Debug settings
    .Debug = True                                           ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: String name of the switch that triggers the auto-fire (Default: Empty)
- `Coil`: String name of the coil to fire (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the auto-fire is triggered (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the auto-fire (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the auto-fire (Default: Array("ball_will_end", "service_mode_entered"))
- `FireEvents`: Array of event names that will trigger firing (Default: Empty Array)
- `ResetEvents`: Array of event names that will reset the auto-fire (Default: Array("machine_reset_phase_3", "ball_starting"))

### Ball Search Settings
- `BallSearchEvents`: Array of event names that will trigger ball search (Default: Array("ball_search_started"))
- `BallSearchDelay`: Integer specifying the delay in milliseconds before firing during ball search (Default: 500)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Auto-Fire Example
```vbscript
' Basic auto-fire configuration
With CreateGlfAutoFire("pop1")
    .Switch = "s_pop1"
    .Coil = "c_pop1"
    .ActionCallback = "PopBumperAction"
End With
```

### Advanced Auto-Fire Example
```vbscript
' Advanced auto-fire configuration
With CreateGlfAutoFire("slingshot1")
    .Switch = "s_slingshot1"
    .Coil = "c_slingshot1"
    .ActionCallback = "SlingshotAction"
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .FireEvents = Array("slingshot_ready")
    .BallSearchEvents = Array("ball_search_started", "ball_search_ended")
    .BallSearchDelay = 750
    .Debug = True
End With
```

## Callback Examples

### Basic Auto-Fire Callback
```vbscript
' Basic auto-fire callback
Sub AutoFireAction(Enabled)
    ' Code to handle auto-fire action
    If Enabled Then
        ' Auto-fire is being triggered
        PlaySound "autofire_sound"
    End If
End Sub
```

### Advanced Auto-Fire Callback
```vbscript
' Advanced auto-fire callback with additional logic
Sub AutoFireAction(Enabled)
    If Enabled Then
        ' Auto-fire is being triggered
        PlaySound "autofire_sound"
        LightOn "l_autofire_active"
        
        ' Add score
        AddScore 1000
    Else
        ' Auto-fire is being reset
        LightOff "l_autofire_active"
    End If
End Sub
```

## Events

The auto-fire system generates several events that you can listen to:

- `{autofire_name}_firing`: Triggered when the auto-fire is triggered
- `{autofire_name}_fired`: Triggered when the auto-fire has completed firing
- `{autofire_name}_enabled`: Triggered when the auto-fire is enabled
- `{autofire_name}_disabled`: Triggered when the auto-fire is disabled

Replace `{autofire_name}` with your actual auto-fire name (prefixed with "autofire_" internally).

## Default Behavior

By default, auto-fire devices are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Reset on `machine_reset_phase_3` and `ball_starting` events
- Ball search triggered on `ball_search_started` event
- Ball search delay of 500 milliseconds
- Debug logging disabled

## Notes

- The auto-fire system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for fire/activate, 0 for reset/deactivate)
- Debug logging can be enabled to track auto-fire state changes
- Auto-fire devices can be configured to fire during ball search with a configurable delay
- The system prevents multiple firing operations from overlapping 