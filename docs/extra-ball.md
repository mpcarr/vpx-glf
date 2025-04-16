# Extra Ball Configuration

The extra ball configuration allows you to set up and customize extra ball awards in your pinball machine. Extra balls are special rewards that give players additional balls to play with, enhancing gameplay and scoring opportunities.

## Configuration Options

### Basic Configuration
```vbscript
' Basic extra ball configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ExtraBall("extra_ball_name")
        .AwardEvents = Array("event1") ' Event that triggers the extra ball award
        .MaxPerGame = 3               ' Maximum number of extra balls that can be awarded per game
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced extra ball configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ExtraBall("extra_ball_name")
        ' Award settings
        .AwardEvents = Array("event1", "event2", "event3") ' Multiple events that can trigger the extra ball award
        
        ' Limit settings
        .MaxPerGame = 5               ' Maximum number of extra balls that can be awarded per game
        
        ' Debug settings
        .Debug = True                 ' Enable debug logging for this extra ball
    End With
End With
```

## Property Descriptions

### Award Settings
- `AwardEvents`: Array of event names that trigger the extra ball award (Default: Empty Array)

### Limit Settings
- `MaxPerGame`: Maximum number of extra balls that can be awarded per game (Default: 0, meaning unlimited)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this extra ball (Default: False)

## Example Configurations

### Basic Extra Ball Example
```vbscript
' Basic extra ball configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ExtraBall("ramp_extra")
        .AwardEvents = Array("ramp_complete")
        .MaxPerGame = 2
    End With
End With
```

### Advanced Extra Ball Example
```vbscript
' Advanced extra ball configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ExtraBall("multiball_extra")
        ' Award settings
        .AwardEvents = Array("left_ramp_complete", "right_ramp_complete", "center_ramp_complete")
        
        ' Limit settings
        .MaxPerGame = 3
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Extra Ball Tracking

The extra ball system tracks:
- The total number of extra balls awarded to the player
- The number of extra balls awarded for each specific extra ball configuration

## Events

The extra ball system generates several events that you can listen to:

- `{extra_ball_name}_awarded`: Triggered when an extra ball is awarded
- `extra_ball_awarded`: Generic event triggered when any extra ball is awarded

Replace `{extra_ball_name}` with your actual extra ball name (prefixed with "extra_ball_" internally).

## Default Behavior

By default, extra balls are configured with:
- No award events
- No maximum per game limit (0)
- Debug logging disabled

## Notes

- Extra balls are managed within the context of a mode
- The extra ball system automatically tracks the number of extra balls awarded
- Debug logging can be enabled to track extra ball operations
- Extra balls can be configured to work with specific game modes or features
- The extra ball system prevents exceeding the maximum per game limit
- Extra balls are awarded by incrementing the player's "extra_balls" state
- Extra balls can be awarded through multiple different events 