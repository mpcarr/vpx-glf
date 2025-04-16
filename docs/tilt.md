# Tilt Configuration

The tilt configuration allows you to set up and customize tilt behavior in your pinball machine. The tilt system manages tilt warnings, tilt events, and slam tilts.

## Configuration Options

### Basic Configuration
```vbscript
' Basic tilt configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Tilt
        .WarningsToTilt = 3
        .SettleTime = 5000  ' 5000 milliseconds (5 seconds) settle time
        .MultipleHitWindow = 1000  ' 1000 milliseconds (1 second) multiple hit window
        .ResetWarningEvents = Array("reset_warnings_event")
        .TiltWarningEvents = Array("tilt_warning_event")
        .TiltEvents = Array("tilt_event")
        .TiltSlamTiltEvents = Array("slam_tilt_event")
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced tilt configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Tilt
        ' Configure tilt settings
        .WarningsToTilt = 3
        .SettleTime = 5000  ' 5000 milliseconds (5 seconds) settle time
        .MultipleHitWindow = 1000  ' 1000 milliseconds (1 second) multiple hit window
        
        ' Configure events
        .ResetWarningEvents = Array("reset_warnings_event1", "reset_warnings_event2")
        .TiltWarningEvents = Array("tilt_warning_event1", "tilt_warning_event2")
        .TiltEvents = Array("tilt_event1", "tilt_event2")
        .TiltSlamTiltEvents = Array("slam_tilt_event1", "slam_tilt_event2")
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Property Descriptions

### Tilt Settings
- `WarningsToTilt`: Number of warnings before a tilt occurs (Default: 0)
- `SettleTime`: Time in milliseconds to wait before clearing a tilt (Default: 0)
- `MultipleHitWindow`: Time window in milliseconds for multiple hits to count as one warning (Default: 0)
- `ResetWarningEvents`: Array of events that reset the tilt warnings
- `TiltWarningEvents`: Array of events that trigger a tilt warning
- `TiltEvents`: Array of events that trigger a tilt
- `TiltSlamTiltEvents`: Array of events that trigger a slam tilt
- `Debug`: Boolean to enable debug logging for this tilt system (Default: False)

## Example Configurations

### Basic Tilt Example
```vbscript
' Basic tilt configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Tilt
        .WarningsToTilt = 3
        .SettleTime = 5000  ' 5000 milliseconds (5 seconds) settle time
        .MultipleHitWindow = 1000  ' 1000 milliseconds (1 second) multiple hit window
        .ResetWarningEvents = Array("reset_warnings")
        .TiltWarningEvents = Array("tilt_warning")
        .TiltEvents = Array("tilt")
        .TiltSlamTiltEvents = Array("slam_tilt")
    End With
End With
```

### Advanced Tilt Example
```vbscript
' Advanced tilt configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Tilt
        ' Configure tilt settings
        .WarningsToTilt = 3
        .SettleTime = 5000  ' 5000 milliseconds (5 seconds) settle time
        .MultipleHitWindow = 1000  ' 1000 milliseconds (1 second) multiple hit window
        
        ' Configure events
        .ResetWarningEvents = Array("reset_warnings", "ball_started")
        .TiltWarningEvents = Array("tilt_warning_switch", "tilt_warning_event")
        .TiltEvents = Array("tilt_switch", "tilt_event")
        .TiltSlamTiltEvents = Array("slam_tilt_switch", "slam_tilt_event")
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Tilt System

The tilt system manages tilt behavior with the following features:

- Tilt warnings can be configured with a specific count before a tilt occurs
- Tilt warnings can be reset by specific events
- Tilt events can be triggered by specific events
- Slam tilt events can be triggered by specific events
- The system tracks the number of tilt warnings for each player
- The system manages the tilt settle time
- The system prevents multiple tilt warnings within a short time window
- The tilt system is managed within the context of a mode

## Events

The tilt system generates the following events:

- `tilt_warning`: Fired when a tilt warning occurs
- `tilt_warning_X`: Fired when the Xth tilt warning occurs
- `tilt`: Fired when a tilt occurs
- `slam_tilt`: Fired when a slam tilt occurs
- `tilt_clear`: Fired when a tilt is cleared

## Default Behavior

By default, the tilt system is configured with:
- No warnings to tilt (WarningsToTilt = 0)
- No settle time (SettleTime = 0)
- No multiple hit window (MultipleHitWindow = 0)
- No reset warning events
- No tilt warning events
- No tilt events
- No slam tilt events
- Debug logging disabled

## Notes

- The tilt system is managed within the context of a mode
- The tilt system automatically handles tilt warnings and tilts
- Debug logging can be enabled to track tilt operations
- The tilt system can be configured to work with specific game modes or features
- The tilt system prevents tilt conflicts
- Tilt warnings are tracked per player
- Tilt warnings are reset when a new ball starts
- The settle time is used to prevent immediate re-tilting
- The multiple hit window prevents rapid-fire tilt warnings
- When a tilt occurs, the current ball ends and all balls drain
- When all balls have drained, the tilt is cleared after the settle time 