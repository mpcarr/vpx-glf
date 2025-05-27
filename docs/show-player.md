# Show Player Configuration

The show player configuration allows you to set up and customize show playback in your game. Show players are used to control lighting effects, and other visual elements.

## Configuration Options

### Basic Configuration
```vbscript
' Basic show player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")

    With .ShowPlayer()
        With .EventName("s_left_sling_active")
            .Key = "key_s_left_sling_active"
            .Show = "flicker_color"
            .Speed = 1
            With .Tokens()
                .Add "lights", "f140"
                .Add "color", "C81616"
            End With
        End With
        With .EventName("s_right_sling_active")
            .Key = "key_s_right_sling_active"
            .Show = "flicker_color"
            .Speed = 1
            With .Tokens()
                .Add "lights", "f141"
                .Add "color", "1616FF"
            End With
        End With
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

