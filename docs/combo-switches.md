# Combo Switches Configuration

The combo switches configuration allows you to set up and customize combinations of switches that trigger events when activated in specific patterns. Combo switches are useful for creating complex gameplay mechanics that require multiple switches to be activated in sequence or simultaneously.

## Configuration Options

### Basic Configuration
```vbscript
' Basic combo switch configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ComboSwitches("combo_name")
        .Switch1 = "switch1"              ' First switch in the combo
        .Switch2 = "switch2"              ' Second switch in the combo
        .EventsWhenBoth = Array("event1") ' Events triggered when both switches are active
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced combo switch configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ComboSwitches("combo_name")
        ' Basic settings
        .Switch1 = "switch1"              ' First switch in the combo
        .Switch2 = "switch2"              ' Second switch in the combo
        
        ' Timing settings
        .HoldTime = 500                   ' Time in milliseconds a switch must be held to be considered active
        .MaxOffsetTime = 1000             ' Maximum time difference between switch activations to trigger "both" state
        .ReleaseTime = 300                ' Time in milliseconds a switch must be released to be considered inactive
        
        ' Event settings
        .EventsWhenBoth = Array("event1", "event2")       ' Events triggered when both switches are active
        .EventsWhenInactive = Array("event3", "event4")   ' Events triggered when no switches are active
        .EventsWhenOne = Array("event5", "event6")        ' Events triggered when only one switch is active
        .EventsWhenSwitch1 = Array("event7", "event8")    ' Events triggered when only switch1 is active
        .EventsWhenSwitch2 = Array("event9", "event10")   ' Events triggered when only switch2 is active
        
        ' Debug settings
        .Debug = True                     ' Enable debug logging for this combo switch
    End With
End With
```

## Property Descriptions

### Basic Settings
- `Switch1`: Name of the first switch in the combo (Default: Empty)
- `Switch2`: Name of the second switch in the combo (Default: Empty)

### Timing Settings
- `HoldTime`: Time in milliseconds a switch must be held to be considered active (Default: 0)
- `MaxOffsetTime`: Maximum time difference between switch activations to trigger "both" state (Default: -1, meaning no limit)
- `ReleaseTime`: Time in milliseconds a switch must be released to be considered inactive (Default: 0)

### Event Control
- `EventsWhenBoth`: Array of event names triggered when both switches are active (Default: Empty Array)
- `EventsWhenInactive`: Array of event names triggered when no switches are active (Default: Empty Array)
- `EventsWhenOne`: Array of event names triggered when only one switch is active (Default: Empty Array)
- `EventsWhenSwitch1`: Array of event names triggered when only switch1 is active (Default: Empty Array)
- `EventsWhenSwitch2`: Array of event names triggered when only switch2 is active (Default: Empty Array)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this combo switch (Default: False)

## Example Configurations

### Basic Combo Switch Example
```vbscript
' Basic combo switch configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ComboSwitches("ramp_combo")
        .Switch1 = "ramp1"
        .Switch2 = "ramp2"
        .EventsWhenBoth = Array("ramp_combo_complete")
    End With
End With
```

### Advanced Combo Switch Example
```vbscript
' Advanced combo switch configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ComboSwitches("multiball_combo")
        ' Basic settings
        .Switch1 = "left_ramp"
        .Switch2 = "right_ramp"
        
        ' Timing settings
        .HoldTime = 300
        .MaxOffsetTime = 2000
        .ReleaseTime = 200
        
        ' Event settings
        .EventsWhenBoth = Array("multiball_start")
        .EventsWhenInactive = Array("multiball_ended")
        .EventsWhenOne = Array("combo_partial")
        .EventsWhenSwitch1 = Array("left_ramp_complete")
        .EventsWhenSwitch2 = Array("right_ramp_complete")
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## States

The combo switch system manages several states:

- `both`: Both switches are active
- `one`: Only one switch is active
- `inactive`: No switches are active
- `switches_1`: Only switch1 is active
- `switches_2`: Only switch2 is active

## Events

The combo switch system generates several events that you can listen to:

- `{combo_name}_switch1_active`: Triggered when switch1 becomes active
- `{combo_name}_switch1_inactive`: Triggered when switch1 becomes inactive
- `{combo_name}_switch2_active`: Triggered when switch2 becomes active
- `{combo_name}_switch2_inactive`: Triggered when switch2 becomes inactive

Replace `{combo_name}` with your actual combo switch name (prefixed with "combo_switch_" internally).

## Default Behavior

By default, combo switches are configured with:
- No switches defined
- No hold time, max offset time, or release time
- No events for any state
- Debug logging disabled

## Notes

- Combo switches are managed within the context of a mode
- The combo switch system automatically handles switch state transitions
- Debug logging can be enabled to track combo switch operations
- The system prevents multiple state transitions from overlapping
- Combo switches can be configured to work with specific game modes or features
- Hold time and release time can be used to create more reliable combo detection
- Max offset time can be used to create time-sensitive combo requirements 