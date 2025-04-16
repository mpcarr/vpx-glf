# Light Player Configuration

The light player configuration allows you to set up and customize light patterns and sequences in your pinball machine. Light players are useful for creating dynamic lighting effects that respond to game events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic light player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .LightPlayer
        With .EventName("event1")
            With .Lights("light1")
                .Color = "ff0000"      ' Red color in RGB hex format
            End With
        End With
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced light player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .LightPlayer
        ' Configure multiple events with multiple lights
        With .EventName("event1")
            With .Lights("light1")
                .Color = "ff0000"      ' Red color in RGB hex format
                .Fade = 500             ' Fade time in milliseconds
                .Priority = 10          ' Priority level (higher numbers take precedence)
            End With
            With .Lights("light2")
                .Color = "00ff00"      ' Green color in RGB hex format
                .Fade = 300             ' Fade time in milliseconds
            End With
        End With
        
        With .EventName("event2")
            With .Lights("light3")
                .Color = "0000ff"      ' Blue color in RGB hex format
                .Fade = 1000            ' Fade time in milliseconds
                .Priority = 20          ' Priority level
            End With
        End With
        
        ' Debug settings
        .Debug = True                  ' Enable debug logging for this light player
    End With
End With
```

## Property Descriptions

### Light Settings
- `Color`: Color of the light in RGB hex format (Default: "ffffff" - white)
- `Fade`: Fade time in milliseconds (Default: Empty, meaning no fade)
- `Priority`: Priority level for the light (Default: 0)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this light player (Default: False)

## Example Configurations

### Basic Light Player Example
```vbscript
' Basic light player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .LightPlayer
        With .EventName("ramp_complete")
            With .Lights("ramp_light")
                .Color = "ff0000"      ' Red color
            End With
        End With
    End With
End With
```

### Advanced Light Player Example
```vbscript
' Advanced light player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .LightPlayer
        ' Configure multiple events with multiple lights
        With .EventName("multiball_start")
            With .Lights("left_ramp")
                .Color = "ff0000"      ' Red color
                .Fade = 500             ' Fade time in milliseconds
                .Priority = 10          ' Priority level
            End With
            With .Lights("right_ramp")
                .Color = "00ff00"      ' Green color
                .Fade = 500             ' Fade time in milliseconds
                .Priority = 10          ' Priority level
            End With
            With .Lights("center_ramp")
                .Color = "0000ff"      ' Blue color
                .Fade = 500             ' Fade time in milliseconds
                .Priority = 10          ' Priority level
            End With
        End With
        
        With .EventName("jackpot_ready")
            With .Lights("jackpot_light")
                .Color = "ffff00"      ' Yellow color
                .Fade = 300             ' Fade time in milliseconds
                .Priority = 20          ' Priority level
            End With
        End With
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Light Stack System

The light player system uses a stack-based approach to manage light states:

- Each light has a stack of color states
- When a new color is set, it's pushed onto the stack
- When a light is turned off, the top color is popped from the stack
- The light displays the color at the top of the stack
- Colors with higher priority take precedence

## Events

The light player system doesn't generate its own events, but it responds to events defined in your configuration. When an event occurs, the light player will automatically set the lights associated with that event.

## Default Behavior

By default, light players are configured with:
- No events or lights
- White color for lights
- No fade time
- Priority level of 0
- Debug logging disabled

## Notes

- Light players are managed within the context of a mode
- The light player system automatically handles light state transitions
- Debug logging can be enabled to track light player operations
- Light players can be configured to work with specific game modes or features
- The light stack system prevents multiple light changes from conflicting
- Light players support tag-based lighting for controlling multiple lights with a single configuration
- Fade effects can be customized for smooth light transitions
- Priority levels can be used to create complex light hierarchies 