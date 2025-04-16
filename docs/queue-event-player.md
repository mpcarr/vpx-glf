# Queue Event Player Configuration

The queue event player configuration allows you to set up and customize queued event playback in your pinball machine. Queue event players are useful for creating sequences of events that play in order, with configurable timing between events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic queue event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .QueueEventPlayer
        With .EventName("event1")
            .Events = Array("response1", "response2")
            .Delay = 500    ' Delay in milliseconds between events
        End With
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced queue event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .QueueEventPlayer
        ' Configure multiple event queues
        With .EventName("event1")
            .Events = Array("response1", "response2", "response3")
            .Delay = 500    ' Delay in milliseconds between events
            .Randomize = True    ' Randomize the order of events
            .Repeat = 2     ' Repeat the sequence 2 times
        End With
        
        With .EventName("event2")
            .Events = Array("response4", "response5")
            .Delay = 1000   ' Longer delay between events
            .Randomize = False   ' Play events in order
            .Repeat = 1     ' Play sequence once
        End With
        
        ' Debug settings
        .Debug = True      ' Enable debug logging for this queue event player
    End With
End With
```

## Property Descriptions

### Queue Settings
- `Events`: Array of event names to play in sequence
- `Delay`: Delay in milliseconds between events (Default: 0)
- `Randomize`: Boolean to randomize the order of events (Default: False)
- `Repeat`: Number of times to repeat the sequence (Default: 1)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this queue event player (Default: False)

## Example Configurations

### Basic Queue Event Player Example
```vbscript
' Basic queue event player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .QueueEventPlayer
        With .EventName("ramp_complete")
            .Events = Array("ramp_light_on", "ramp_light_off")
            .Delay = 500    ' Half second between events
        End With
    End With
End With
```

### Advanced Queue Event Player Example
```vbscript
' Advanced queue event player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .QueueEventPlayer
        ' Configure multiple event queues
        With .EventName("multiball_start")
            .Events = Array("left_ramp_on", "right_ramp_on", "center_ramp_on")
            .Delay = 500    ' Half second between events
            .Randomize = True    ' Randomize the order of events
            .Repeat = 2     ' Repeat the sequence 2 times
        End With
        
        With .EventName("jackpot_ready")
            .Events = Array("jackpot_light_on", "jackpot_light_off")
            .Delay = 1000   ' One second between events
            .Randomize = False   ' Play events in order
            .Repeat = 1     ' Play sequence once
        End With
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Queue System

The queue event player system manages event sequences with the following features:

- Events are played in sequence with configurable delays
- Sequences can be randomized for variety
- Sequences can be repeated multiple times
- Each queue is independent and can be configured differently
- Queues are cleared when the mode stops

## Events

The queue event player system doesn't generate its own events, but it plays the events defined in your configuration. When a trigger event occurs, the queue event player will automatically play the associated sequence of events.

## Default Behavior

By default, queue event players are configured with:
- No events or queues
- No delay between events
- Events played in order (not randomized)
- Sequences played once
- Debug logging disabled

## Notes

- Queue event players are managed within the context of a mode
- The queue system automatically handles event timing and sequencing
- Debug logging can be enabled to track queue event player operations
- Queue event players can be configured to work with specific game modes or features
- The queue system prevents events from overlapping or conflicting
- Randomization can be used to create variety in event sequences
- Repeat counts can be used to create longer sequences
- Delays can be customized to create rhythm and timing in event sequences 