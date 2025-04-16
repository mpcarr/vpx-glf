# Diverter Configuration

The diverter configuration allows you to set up and customize diverters in your pinball machine. Diverters are used to change the path of balls by activating mechanical devices.

## Configuration Options

### Basic Configuration
```vbscript
' Basic diverter configuration
With CreateGlfDiverter("diverter_name")
    .ActionCallback = "DiverterAction"    ' Name of the callback function for diverter actions
    .ActivationTime = 1000                ' Time in milliseconds to hold the diverter active
End With
```

### Advanced Configuration
```vbscript
' Advanced diverter configuration
With CreateGlfDiverter("diverter_name")
    ' Basic settings
    .ActionCallback = "DiverterAction"
    .ActivationTime = 1000

    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the diverter
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the diverter
    .ActivateEvents = Array("event1", "event2")              ' Events that activate the diverter
    .DeactivateEvents = Array("event3", "event4")            ' Events that deactivate the diverter

    ' Switch settings
    .ActivationSwitches = Array("switch1", "switch2")        ' Switches that activate the diverter

    ' Ball search settings
    .BallSearchHoldTime = 1000                               ' Time to hold during ball search
    .ExcludeFromBallSearch = False                           ' Whether to exclude from ball search

    ' Debug settings
    .Debug = True                                            ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `ActionCallback`: String name of the callback function that executes when the diverter activates or deactivates
- `ActivationTime`: Integer specifying how long to hold the diverter active in milliseconds

### Event Control
- `EnableEvents`: Array of event names that will enable the diverter
- `DisableEvents`: Array of event names that will disable the diverter
- `ActivateEvents`: Array of event names that will activate the diverter
- `DeactivateEvents`: Array of event names that will deactivate the diverter

### Switch Settings
- `ActivationSwitches`: Array of switch names that will activate the diverter when triggered

### Ball Search Settings
- `BallSearchHoldTime`: Integer specifying how long to hold the diverter during ball search
- `ExcludeFromBallSearch`: Boolean indicating whether to exclude this device from ball search operations

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device

## Example Configurations

### Ramp Diverter Example
```vbscript
' Ramp diverter configuration
With CreateGlfDiverter("ramp_diverter")
    .ActionCallback = "RampDiverterAction"
    .ActivationTime = 500
    .EnableEvents = Array("ball_started")
    .DisableEvents = Array("ball_will_end")
    .ActivateEvents = Array("ramp_switch")
    .DeactivateEvents = Array("ball_exit")
    .ActivationSwitches = Array("s_diverter")
    .BallSearchHoldTime = 1000
    .Debug = True
End With
```

### Multiball Diverter Example
```vbscript
' Multiball diverter configuration
With CreateGlfDiverter("multiball_diverter")
    .ActionCallback = "MultiballDiverterAction"
    .ActivationTime = 2000
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .ActivateEvents = Array("multiball_ready")
    .DeactivateEvents = Array("multiball_complete")
    .ExcludeFromBallSearch = True
End With
```

## Events

The diverter system generates several events that you can listen to:

- `{diverter_name}_activating`: Triggered when the diverter is activating
- `{diverter_name}_deac# Diverter Configuration

The diverter configuration allows you to set up and customize diverters in your pinball machine. Diverters are used to change the path of balls, typically to direct them to different areas of the playfield.

## Configuration Options

### Basic Configuration
```vbscript
' Basic diverter configuration
With CreateGlfDiverter("diverter_name")
    .Switch = "s_diverter_switch"       ' Name of the switch that triggers the diverter
    .Coil = "c_diverter_coil"           ' Name of the coil to activate the diverter
    .ActionCallback = "DiverterAction"  ' Name of the callback function for diverter actions
End With
```

### Advanced Configuration
```vbscript
' Advanced diverter configuration
With CreateGlfDiverter("diverter_name")
    ' Basic settings
    .Switch = "s_diverter_switch"
    .Coil = "c_diverter_coil"
    .ActionCallback = "DiverterAction"
    
    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the diverter
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the diverter
    .ActivateEvents = Array("event1", "event2")              ' Events that trigger activation
    .DeactivateEvents = Array("event3", "event4")            ' Events that trigger deactivation
    .ResetEvents = Array("event5", "event6")                 ' Events that reset the diverter
    
    ' Switch settings
    .ActivateSwitch = "s_activate"                           ' Switch that activates the diverter
    .DeactivateSwitch = "s_deactivate"                       ' Switch that deactivates the diverter
    
    ' Ball search settings
    .BallSearchEvents = Array("ball_search_started")         ' Events that trigger ball search
    .BallSearchDelay = 500                                   ' Delay before ball search activation
    
    ' Debug settings
    .Debug = True                                           ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: String name of the switch that triggers the diverter (Default: Empty)
- `Coil`: String name of the coil to activate the diverter (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the diverter is triggered (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the diverter (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the diverter (Default: Array("ball_will_end", "service_mode_entered"))
- `ActivateEvents`: Array of event names that will trigger activation (Default: Empty Array)
- `DeactivateEvents`: Array of event names that will trigger deactivation (Default: Empty Array)
- `ResetEvents`: Array of event names that will reset the diverter (Default: Array("machine_reset_phase_3", "ball_starting"))

### Switch Settings
- `ActivateSwitch`: String name of the switch that activates the diverter (Default: Empty)
- `DeactivateSwitch`: String name of the switch that deactivates the diverter (Default: Empty)

### Ball Search Settings
- `BallSearchEvents`: Array of event names that will trigger ball search (Default: Array("ball_search_started"))
- `BallSearchDelay`: Integer specifying the delay in milliseconds before activation during ball search (Default: 500)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Diverter Example
```vbscript
' Basic diverter configuration
With CreateGlfDiverter("ramp1")
    .Switch = "s_ramp1"
    .Coil = "c_ramp1"
    .ActionCallback = "RampDiverterAction"
End With
```

### Advanced Diverter Example
```vbscript
' Advanced diverter configuration
With CreateGlfDiverter("multiball1")
    .Switch = "s_multiball1"
    .Coil = "c_multiball1"
    .ActionCallback = "MultiballDiverterAction"
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .ActivateEvents = Array("diverter_ready")
    .DeactivateEvents = Array("diverter_reset")
    .ActivateSwitch = "s_activate_multiball"
    .DeactivateSwitch = "s_deactivate_multiball"
    .BallSearchEvents = Array("ball_search_started", "ball_search_ended")
    .BallSearchDelay = 750
    .Debug = True
End With
```

## Callback Examples

### Basic Diverter Callback
```vbscript
' Basic diverter callback
Sub DiverterAction(Enabled)
    ' Code to handle diverter action
    If Enabled Then
        ' Diverter is being activated
        PlaySound "diverter_sound"
    End If
End Sub
```

### Advanced Diverter Callback
```vbscript
' Advanced diverter callback with additional logic
Sub DiverterAction(Enabled)
    If Enabled Then
        ' Diverter is being activated
        PlaySound "diverter_sound"
        LightOn "l_diverter_active"
        
        ' Add score
        AddScore 1000
    Else
        ' Diverter is being deactivated
        LightOff "l_diverter_active"
    End If
End Sub
```

## Events

The diverter system generates several events that you can listen to:

- `{diverter_name}_activating`: Triggered when the diverter is being activated
- `{diverter_name}_activated`: Triggered when the diverter has been activated
- `{diverter_name}_deactivating`: Triggered when the diverter is being deactivated
- `{diverter_name}_deactivated`: Triggered when the diverter has been deactivated
- `{diverter_name}_enabled`: Triggered when the diverter is enabled
- `{diverter_name}_disabled`: Triggered when the diverter is disabled

Replace `{diverter_name}` with your actual diverter name (prefixed with "diverter_" internally).

## Default Behavior

By default, diverters are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Reset on `machine_reset_phase_3` and `ball_starting` events
- Ball search triggered on `ball_search_started` event
- Ball search delay of 500 milliseconds
- Debug logging disabled

## Notes

- The diverter system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for activate, 0 for deactivate)
- Debug logging can be enabled to track diverter state changes
- Diverters can be configured to activate during ball search with a configurable delay
- The system prevents multiple activation/deactivation operations from overlapping
- Diverters can be controlled by both events and switches tivating`: Triggered when the diverter is deactivating

Replace `{diverter_name}` with your actual diverter name (prefixed with "diverter_" internally).

## Default Behavior

By default, diverters are configured with:
- No enable/disable events
- No activation/deactivation events
- No activation switches
- Activation time of 0 milliseconds
- Ball search hold time of 1000 milliseconds
- Debug logging disabled
- Not excluded from ball search

## Notes

- The diverter system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for activate, 0 for deactivate)
- Debug logging can be enabled to track diverter state changes
- Ball search functionality is available for troubleshooting
- Diverters can be configured to automatically deactivate after a specified time 