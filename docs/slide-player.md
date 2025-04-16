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

### Advanced Configuration
```vbscript
' Advanced slide player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SlidePlayer
        ' Configure debug
        .Debug = True
        
        ' Configure slide events
        With .EventName("target_hit")
            .Slide = "target_slide"
            .Action = "play"
            .Expire = 5000  ' 5000 milliseconds (5 seconds) expire time
            .MaxQueueTime = 1000  ' 1000 milliseconds (1 second) max queue time
            .Method = "queue"
            .Priority = 10
            .Target = "display1"
            .Tokens = Array("token1", "token2")
        End With
        
        With .EventName("multiball_start")
            .Slide = "multiball_slide"
            .Action = "play"
            .Method = "immediate"
            .Priority = 20
            .Target = "display1"
        End With
    End With
End With
```

## Property Descriptions

### Slide Settings
- `Slide`: The name of the slide to play
- `Action`: Action to perform on the slide ("play") (Default: "play")
- `Expire`: Time in milliseconds before the slide expires (Default: Empty)
- `MaxQueueTime`: Maximum time in milliseconds to queue the slide (Default: Empty)
- `Method`: Method to use for playing the slide ("queue" or "immediate") (Default: Empty)
- `Priority`: Priority level for the slide (Default: Empty)
- `Target`: Target display for the slide (Default: Empty)
- `Tokens`: Array of tokens to pass to the slide (Default: Empty)
- `Debug`: Boolean to enable debug logging for this slide player (Default: False)

## Example Configurations

### Basic Slide Player Example
```vbscript
' Basic slide player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SlidePlayer
        With .EventName("target_hit")
            .Slide = "target_slide"
            .Action = "play"
        End With
    End With
End With
```

### Advanced Slide Player Example
```vbscript
' Advanced slide player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SlidePlayer
        ' Configure debug
        .Debug = True
        
        ' Configure slide events
        With .EventName("jackpot_hit")
            .Slide = "jackpot_slide"
            .Action = "play"
            .Method = "immediate"
            .Priority = 30
            .Target = "display1"
            .Tokens = Array("score", "10000")
        End With
        
        With .EventName("multiball_start")
            .Slide = "multiball_slide"
            .Action = "play"
            .Method = "queue"
            .Priority = 20
            .Target = "display1"
            .Expire = 10000  ' 10000 milliseconds (10 seconds) expire time
        End With
    End With
End With
```

## Slide Player System

The slide player system manages slide animations with the following features:

- Slides can be played in response to events
- Slides can be configured with various settings like expire time, queue time, method, priority, target, and tokens
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
- The slide player system prevents slide conflicts
- Slides must be defined in the slide system before they can be played
- The `Method` property determines how the slide is played:
  - `queue`: The slide is queued and played when its turn comes
  - `immediate`: The slide is played immediately, interrupting any currently playing slides
- The `Priority` property determines the order in which queued slides are played
- The `Target` property specifies which display the slide is played on
- The `Tokens` property can be used to pass data to the slide
- The `Expire` property determines how long the slide can be queued before it expires
- The `MaxQueueTime` property determines the maximum time a slide can be queued 