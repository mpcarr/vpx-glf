# Timer Configuration

The timer configuration allows you to set up and customize countdown or count-up timers in your pinball machine. Timers are useful for creating time-limited modes, bonus rounds, and other time-based events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic timer configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("timer_name")
        .StartEvents = Array("start_event")
        .StopEvents = Array("stop_event")
        .Direction = "down"
        .StartValue = 10
        .EndValue = 0
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced timer configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("timer_name")
        ' Configure start and stop events
        .StartEvents = Array("start_event1", "start_event2")
        .StopEvents = Array("stop_event1", "stop_event2")
        
        ' Configure timer direction and values
        .Direction = "down"
        .StartValue = 30
        .EndValue = 0
        
        ' Configure tick interval
        .TickInterval = 1000    ' Tick every 1000 milliseconds (1 second)
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Property Descriptions

### Timer Settings
- `StartEvents`: Array of events that will start the timer
- `StopEvents`: Array of events that will stop the timer
- `Direction`: Direction of the timer ("up" or "down") (Default: "down")
- `StartValue`: Initial value of the timer (Default: 10)
- `EndValue`: Final value of the timer (Default: 0)
- `TickInterval`: Interval between timer ticks in milliseconds (Default: 1000)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this timer (Default: False)

## Example Configurations

### Basic Timer Example
```vbscript
' Basic timer configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("bonus_timer")
        .StartEvents = Array("bonus_start")
        .StopEvents = Array("bonus_end")
        .Direction = "down"
        .StartValue = 10
        .EndValue = 0
    End With
End With
```

### Advanced Timer Example
```vbscript
' Advanced timer configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("jackpot_timer")
        ' Configure start and stop events
        .StartEvents = Array("jackpot_start")
        .StopEvents = Array("jackpot_end", "jackpot_complete")
        
        ' Configure timer direction and values
        .Direction = "down"
        .StartValue = 30
        .EndValue = 0
        
        ' Configure tick interval
        .TickInterval = 1000    ' Tick every 1000 milliseconds (1 second)
        
        ' Configure debug
        .Debug = True
    End With
End With
```

## Timer System

The timer system manages countdown or count-up timers with the following features:

- Timers can count up or down
- Timers can be started and stopped by events
- Timers can have configurable tick intervals
- Timers can trigger events when they complete
- Timers are managed within the context of a mode

## Events

The timer system generates the following events:

- `timer_name_started`: Fired when the timer starts
- `timer_name_stopped`: Fired when the timer stops
- `timer_name_complete`: Fired when the timer reaches its end value
- `timer_name_tick`: Fired on each timer tick

## Default Behavior

By default, timers are configured with:
- No start or stop events
- Down direction
- Start value of 10
- End value of 0
- Tick interval of 1000 milliseconds (1 second)
- Debug logging disabled

## Notes

- Timers are managed within the context of a mode
- The timer system automatically handles timer ticks and completion
- Debug logging can be enabled to track timer operations
- Timers can be configured to work with specific game modes or features
- The timer system prevents timers from conflicting with each other
- Countdown timers are useful for creating urgency in game modes
- Count-up timers are useful for tracking elapsed time
- Tick events can be used to update displays or trigger other events
- Timer completion events can be used to end modes or trigger rewards