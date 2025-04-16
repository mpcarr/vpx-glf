# Flipper Configuration

The flipper configuration allows you to set up and customize flippers in your pinball machine. Flippers are the primary control mechanism for players to interact with the ball.

## Configuration Options

### Basic Configuration
```vbscript
' Basic flipper configuration
With CreateGlfFlipper("flipper_name")
    .Switch = "s_flipper_switch"       ' Name of the switch that controls the flipper
    .Coil = "c_flipper_coil"           ' Name of the coil that powers the flipper
    .ActionCallback = "FlipperAction"  ' Name of the callback function for flipper actions
End With
```

### Advanced Configuration
```vbscript
' Advanced flipper configuration
With CreateGlfFlipper("flipper_name")
    ' Basic settings
    .Switch = "s_flipper_switch"
    .Coil = "c_flipper_coil"
    .ActionCallback = "FlipperAction"
    
    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the flipper
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the flipper
    .ResetEvents = Array("event1", "event2")                 ' Events that reset the flipper
    
    ' Debug settings
    .Debug = True                                           ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: String name of the switch that controls the flipper (Default: Empty)
- `Coil`: String name of the coil that powers the flipper (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the flipper is activated (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the flipper (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the flipper (Default: Array("ball_will_end", "service_mode_entered"))
- `ResetEvents`: Array of event names that will reset the flipper (Default: Array("machine_reset_phase_3", "ball_starting"))

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Flipper Example
```vbscript
' Basic flipper configuration
With CreateGlfFlipper("left")
    .Switch = "s_left_flipper"
    .Coil = "c_left_flipper"
    .ActionCallback = "LeftFlipperAction"
End With
```

### Advanced Flipper Example
```vbscript
' Advanced flipper configuration
With CreateGlfFlipper("right")
    .Switch = "s_right_flipper"
    .Coil = "c_right_flipper"
    .ActionCallback = "RightFlipperAction"
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .ResetEvents = Array("flipper_reset")
    .Debug = True
End With
```

## Callback Examples

### Basic Flipper Callback
```vbscript
' Basic flipper callback
Sub FlipperAction(Enabled)
    ' Code to handle flipper action
    If Enabled Then
        ' Flipper is being activated
        PlaySound "flipper_up"
    Else
        ' Flipper is being deactivated
        PlaySound "flipper_down"
    End If
End Sub
```

### Advanced Flipper Callback
```vbscript
' Advanced flipper callback with additional logic
Sub FlipperAction(Enabled)
    If Enabled Then
        ' Flipper is being activated
        PlaySound "flipper_up"
        LightOn "l_flipper_active"
        
        ' Add score
        AddScore 100
    Else
        ' Flipper is being deactivated
        PlaySound "flipper_down"
        LightOff "l_flipper_active"
    End If
End Sub
```

## Events

The flipper system generates several events that you can listen to:

- `{flipper_name}_activated`: Triggered when the flipper is activated
- `{flipper_name}_deactivated`: Triggered when the flipper is deactivated
- `{flipper_name}_enabled`: Triggered when the flipper is enabled
- `{flipper_name}_disabled`: Triggered when the flipper is disabled

Replace `{flipper_name}` with your actual flipper name (prefixed with "flipper_" internally).

## Default Behavior

By default, flippers are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Reset on `machine_reset_phase_3` and `ball_starting` events
- Debug logging disabled

## Notes

- The flipper system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for activate, 0 for deactivate)
- Debug logging can be enabled to track flipper state changes
- Flippers can be configured to activate during specific game modes
- The system prevents multiple activation/deactivation operations from overlapping 