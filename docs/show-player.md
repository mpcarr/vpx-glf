# Show Player Configuration

The show player configuration allows you to set up and customize show playback in your pinball machine. Show players are used to control animations, lighting effects, and other visual elements.

## Configuration Options

### Basic Configuration
```vbscript
' Basic show player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShowPlayer("show_name")
        .Show = "show_name"
        .Loops = 1
        .Speed = 1.0
        .SyncMs = 0
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced show player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShowPlayer("show_name")
        ' Configure show settings
        .Show = "show_name"
        .Loops = 2
        .Speed = 1.5
        .SyncMs = 100
        
        ' Configure show tokens
        With .Tokens
            .Add "token1", "value1"
            .Add "token2", "value2"
        End With
        
        ' Configure priority
        .Priority = 10
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Property Descriptions

### Show Settings
- `Show`: Name of the show to play
- `Loops`: Number of times to loop the show (Default: 1)
- `Speed`: Speed multiplier for the show (Default: 1.0)
- `SyncMs`: Synchronization time in milliseconds (Default: 0)

### Show Tokens
- `Tokens`: Dictionary of tokens to pass to the show

### Priority Settings
- `Priority`: Priority level for the show (Default: 0)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this show player (Default: False)

## Example Configurations

### Basic Show Player Example
```vbscript
' Basic show player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShowPlayer("ramp_light")
        .Show = "ramp_light_show"
        .Loops = 1
        .Speed = 1.0
        .SyncMs = 0
    End With
End With
```

### Advanced Show Player Example
```vbscript
' Advanced show player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShowPlayer("jackpot_light")
        ' Configure show settings
        .Show = "jackpot_light_show"
        .Loops = 2
        .Speed = 1.5
        .SyncMs = 100
        
        ' Configure show tokens
        With .Tokens
            .Add "color", "red"
            .Add "intensity", "high"
        End With
        
        ' Configure priority
        .Priority = 10
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Show Player System

The show player system manages show playback with the following features:

- Shows can be played with configurable loops, speed, and synchronization
- Shows can have custom tokens passed to them
- Shows can have different priority levels
- Shows are managed within the context of a mode

## Events

The show player system doesn't generate its own events, but it plays the shows defined in your configuration.

## Default Behavior

By default, show players are configured with:
- No show defined
- One loop
- Normal speed (1.0)
- No synchronization
- No tokens
- Default priority (0)
- Debug logging disabled

## Notes

- Show players are managed within the context of a mode
- The show player system automatically handles show playback
- Debug logging can be enabled to track show player operations
- Show players can be configured to work with specific game modes or features
- The show player system prevents shows from conflicting with each other
- Show tokens can be used to customize shows for different situations
- Show synchronization can be used to coordinate multiple shows
- The loop feature can be used to create continuous animations
- Priority levels can be used to control which shows take precedence
