# Slide Player Configuration

The slide player configuration allows you to set up and customize slide animations in your pinball machine. Slide players are used to play slide animations in response to events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic slide player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SlidePlayer
        With .EventName("event_name")
            .Slide = "slide_name"
            .Action = "play"
        End With
    End With
End With
```

## Property Descriptions

### Slide Settings
- `Slide`: The name of the slide to play
- `Action`: Action to perform on the slide ("play") (Default: "play")
- `Expire`: Time in milliseconds before the slide expires (Default: Empty)
- `Priority`: Priority level for the slide (Default: Empty)
- `Debug`: Boolean to enable debug logging for this slide player (Default: False)

## Slide Player System

The slide player system manages slide animations with the following features:

- Slides can be played in response to events
- Slide players are managed within the context of a mode

## Events

The slide player system doesn't generate its own events, but it responds to events defined in your configuration.

## Default Behavior

By default, slide players are configured with:
- No events defined
- Debug logging disabled

## Notes

- Slide players are managed within the context of a mode
- The slide player system automatically handles slide playback
- Debug logging can be enabled to track slide player operations
- Slide players can be configured to work with specific game modes or features