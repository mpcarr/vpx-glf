# Timed Switches Configuration

The timed switches configuration allows you to set up and customize switches that activate after a specified delay in your pinball machine. Timed switches are useful for creating delayed reactions to switch hits.

## Configuration Options

### Basic Configuration
```vbscript
' Basic timed switches configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .TimedSwitches("timed_switch_name")
        .Time = 1000  ' 1000 milliseconds (1 second) delay
        .Switches = Array("sw01", "sw02")
        .EventsWhenActive = Array("timed_switch_active")
        .EventsWhenReleased = Array("timed_switch_released")
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced timed switches configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .TimedSwitches("timed_switch_name")
        ' Configure delay time
        .Time = 2000  ' 2000 milliseconds (2 seconds) delay
        
        ' Configure switches to monitor
        .Switches = Array("sw01", "sw02", "sw03")
        
        ' Configure events
        .EventsWhenActive = Array("timed_switch_active", "multiball_start")
        .EventsWhenReleased = Array("timed_switch_released", "multiball_end")
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Property Descriptions

### Timed Switch Settings
- `Time`: The delay in milliseconds before a switch is considered active (Default: 0)
- `Switches`: Array of switch names to monitor
- `EventsWhenActive`: Array of events to fire when any of the switches become active
- `EventsWhenReleased`: Array of events to fire when all switches are released
- `Debug`: Boolean to enable debug logging for this timed switch (Default: False)

## Example Configurations

### Basic Timed Switch Example
```vbscript
' Basic timed switch configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .TimedSwitches("ramp_switch")
        .Time = 1000  ' 1000 milliseconds (1 second) delay
        .Switches = Array("sw_ramp")
        .EventsWhenActive = Array("ramp_active")
        .EventsWhenReleased = Array("ramp_released")
    End With
End With
```

### Advanced Timed Switch Example
```vbscript
' Advanced timed switch configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .TimedSwitches("jackpot_switch")
        ' Configure delay time
        .Time = 2000  ' 2000 milliseconds (2 seconds) delay
        
        ' Configure switches to monitor
        .Switches = Array("sw_jackpot1", "sw_jackpot2", "sw_jackpot3")
        
        ' Configure events
        .EventsWhenActive = Array("jackpot_active", "jackpot_ready")
        .EventsWhenReleased = Array("jackpot_released", "jackpot_reset")
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Timed Switch System

The timed switch system manages delayed switch activation with the following features:

- Switches can be configured to activate after a specified delay
- Multiple switches can be monitored simultaneously
- Events can be fired when switches become active or are released
- The system tracks which switches are currently active
- Timed switches are managed within the context of a mode

## Events

The timed switch system generates the following events:

- Events defined in `EventsWhenActive` when any of the switches become active
- Events defined in `EventsWhenReleased` when all switches are released

## Default Behavior

By default, timed switches are configured with:
- No delay (Time = 0)
- No switches to monitor
- No events when active or released
- Debug logging disabled

## Notes

- Timed switches are managed within the context of a mode
- The timed switch system automatically handles switch activation and release
- Debug logging can be enabled to track timed switch operations
- Timed switches can be configured to work with specific game modes or features
- The timed switch system prevents switch conflicts
- The delay is applied to each switch individually
- A switch is considered active only after the delay has passed
- All switches must be released before the release events are fired
- The system automatically removes any pending delays when a switch is released 