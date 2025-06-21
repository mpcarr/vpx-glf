# Timer Configuration

The timer configuration allows you to set up and customize countdown or count-up timers in your pinball machine. Timers are useful for creating time-limited modes, bonus rounds, and other time-based events.

## Configuration Overview

Timers are always managed within the context of a mode. The mode itself is responsible for starting and stopping, and you can configure the timer to start running automatically or control it via events and control actions.

## Basic Configuration
```vbscript
' Basic timer configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("timer_name")
        .StartRunning = True         ' Start timer automatically when mode starts
        .Direction = "down"         ' Count down
        .StartValue = 10            ' Start at 10
        .EndValue = 0               ' End at 0
        .TickInterval = 1000        ' Tick every 1 second
    End With
End With
```

## Controlling the Timer with Events
You can control the timer dynamically using control events. For example, you can add time, pause, or change the tick interval:

```vbscript
With .Timer("timer_name")
    ' ... basic config ...
    With .ControlEvents
        .EventName = "add_time"
        .Action = "add"
        .Value = 5000   ' Add 5000 ticks (or ms, depending on your logic)
    End With
    With .ControlEvents
        .EventName = "pause_timer"
        .Action = "pause"
        .Value = 2000   ' Pause for 2 seconds
    End With
End With
```

## Property Descriptions

### Timer Settings
- `StartRunning`: Boolean, if true, timer starts automatically when the mode starts
- `Direction`: Direction of the timer ("up" or "down") (Default: "down")
- `StartValue`: Initial value of the timer (Default: 10)
- `EndValue`: Final value of the timer (Default: 0)
- `TickInterval`: Interval between timer ticks in milliseconds (Default: 1000)

### Control Events
- `ControlEvents`: Configure actions (add, subtract, pause, etc.) that can be triggered by events

### Debug Settings
- `Debug`: Boolean to enable debug logging for this timer (Default: False)

## Example Configuration
```vbscript
With CreateGlfMode("hurry_up", 10)
    .StartEvents = Array("hurry_up_start")
    .StopEvents = Array("timer_hurry_up_complete")

    With .Timer("hurry_up_timer")
        .StartRunning = True
        .Direction = "down"
        .StartValue = 15
        .EndValue = 0
        .TickInterval = 1000
    End With
End With
```

## Timer System Features
- Timers can count up or down
- Timers can be started automatically or controlled by events/actions
- Timers can have configurable tick intervals
- Timers can trigger events when they complete
- Timers are managed within the context of a mode

## Timer Events
The timer system generates the following events:
- `timer_name_started`: Fired when the timer starts
- `timer_name_stopped`: Fired when the timer stops
- `timer_name_complete`: Fired when the timer reaches its end value
- `timer_name_tick`: Fired on each timer tick

## Default Behavior
By default, timers are configured with:
- Not running until started by the mode or a control event
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