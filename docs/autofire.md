# Auto-Fire Configuration

The auto-fire configuration allows you to set up and customize auto-firing devices in your pinball machine. Auto-fire devices are used to automatically fire coils or other devices when triggered by switches or events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic auto-fire configuration
With CreateGlfAutoFire("autofire_name")
    .Switch = "s_autofire_switch"       ' Name of the switch that triggers the auto-fire
    .ActionCallback = "AutoFireAction"  ' Name of the callback function for auto-fire actions
End With
```

### Advanced Configuration
```vbscript
' Advanced auto-fire configuration
With CreateGlfAutoFire("autofire_name")
    ' Basic settings
    .Switch = "s_autofire_switch"
    .ActionCallback = "AutoFireAction"
    
    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the auto-fire
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the auto-fire
    
    ' Debug settings
    .Debug = True                                           ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: String name of the switch that triggers the auto-fire (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the auto-fire is triggered (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the auto-fire (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the auto-fire (Default: Array("ball_will_end", "service_mode_entered"))

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Auto-Fire Example
```vbscript
' Basic auto-fire configuration
With CreateGlfAutoFire("pop1")
    .Switch = "s_pop1"
    .ActionCallback = "PopBumperAction"
End With
```

### Advanced Auto-Fire Example
```vbscript
' Advanced auto-fire configuration
With CreateGlfAutoFire("slingshot1")
    .Switch = "s_slingshot1"
    .ActionCallback = "SlingshotAction"
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .Debug = True
End With
```

## Callback Examples

The callback will received an array with two params (enabled, ball). This is useful if you want to play a sound at the balls location or if using advanced physics, you need to correct the balls velocity at the slingshots.

### Basic Auto-Fire Callback
```vbscript
' Basic auto-fire callback
Sub AutoFireAction(args)
    Dim enabled, ball : enabled = args(0)
    ' Code to handle auto-fire action
    If enabled Then
        ' Auto-fire is being triggered
        PlaySound "autofire_sound"
    End If
End Sub
```

## Events

## Default Behavior

By default, auto-fire devices are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Debug logging disabled
