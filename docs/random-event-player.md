# Random Event Player Configuration

The random event player configuration allows you to set up and customize random event generation in your pinball machine. Random event players are used to generate random events based on configured probabilities and conditions.

## Configuration Options

### Basic Configuration
```vbscript
' Basic random event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .RandomEventPlayer
        With .EventName("random_event")
            .Events = Array("event1", "event2", "event3")
            .Probability = 0.5  ' 50% chance of triggering
        End With
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced random event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .RandomEventPlayer
        ' Configure debug
        .Debug = True
        
        ' Configure random events
        With .EventName("bonus_event")
            .Events = Array("bonus_1000", "bonus_2000", "bonus_3000")
            .Probability = 0.3  ' 30% chance of triggering
            .MinTime = 1000  ' Minimum 1 second between events
            .MaxTime = 5000  ' Maximum 5 seconds between events
            .Conditions = Array("multiball_active", "jackpot_ready")
        End With
        
        With .EventName("special_event")
            .Events = Array("special_1", "special_2")
            .Probability = 0.1  ' 10% chance of triggering
            .MinTime = 5000  ' Minimum 5 seconds between events
            .MaxTime = 15000  ' Maximum 15 seconds between events
            .Conditions = Array("multiball_active")
        End With
    End With
End With
```

## Property Descriptions

### Random Event Settings
- `Events`: Array of possible events to trigger
- `Probability`: Probability of triggering the random event (0.0 to 1.0) (Default: 0.5)
- `MinTime`: Minimum time in milliseconds between events (Default: Empty)
- `MaxTime`: Maximum time in milliseconds between events (Default: Empty)
- `Conditions`: Array of conditions that must be met for the event to trigger (Default: Empty)
- `Debug`: Boolean to enable debug logging for this random event player (Default: False)

## Example Configurations

### Basic Random Event Player Example
```vbscript
' Basic random event player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .RandomEventPlayer
        With .EventName("bonus_event")
            .Events = Array("bonus_1000", "bonus_2000")
            .Probability = 0.5
        End With
    End With
End With
```

### Advanced Random Event Player Example
```vbscript
' Advanced random event player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .RandomEventPlayer
        ' Configure debug
        .Debug = True
        
        ' Configure random events
        With .EventName("jackpot_event")
            .Events = Array("jackpot_10000", "jackpot_20000", "jackpot_30000")
            .Probability = 0.2
            .MinTime = 2000
            .MaxTime = 8000
            .Conditions = Array("multiball_active", "jackpot_ready")
        End With
        
        With .EventName("special_event")
            .Events = Array("special_1", "special_2", "special_3")
            .Probability = 0.1
            .MinTime = 5000
            .MaxTime = 15000
            .Conditions = Array("multiball_active")
        End With
    End With
End With
```

## Random Event Player System

The random event player system manages random event generation with the following features:

- Random events can be configured with probabilities and conditions
- Events can be timed with minimum and maximum intervals
- Random event players are managed within the context of a mode

## Events

The random event player system generates events based on the configured probabilities and conditions.

## Default Behavior

By default, random event players are configured with:
- No events defined
- Debug logging disabled

## Notes

- Random event players are managed within the context of a mode
- The random event player system automatically handles event generation
- Debug logging can be enabled to track random event player operations
- Random event players can be configured to work with specific game modes or features
- The `Probability` property determines how often the random event will trigger
- The `MinTime` and `MaxTime` properties determine the timing between events
- The `Conditions` property can be used to restrict when random events can trigger
- Random events are only generated when all conditions are met
- The system uses a random number generator to determine when to trigger events
- Events are selected randomly from the configured array of possible events 