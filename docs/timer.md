# Timer Configuration

The timer configuration allows you to set up and customize countdown or count-up timers. Timers are useful for creating time-limited modes, bonus rounds, and other time-based events.

## Configuration Overview

Timers are always managed within the context of a mode. The mode itself is responsible for starting and stopping, and you can configure the timer to start running automatically or control it via events and control actions.

## Basic Configuration
```vbscript
' Basic timer configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("timer_name")
        .StartRunning = True        ' Start timer automatically when mode starts
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

' Example event player to dispatch custom events.
With .EventPlayer
    .Add("s_switch01_active", Array("add_ticks_to_my_timer")
    .Add("s_switch02_active", Array("pause_my_timer")
End With

With .Timer("timer_name")
    With .ControlEvents
        .EventName = "add_ticks_to_my_timer"
        .Action = "add"
        .Value = 5   ' Add 5 ticks,
    End With
    With .ControlEvents
        .EventName = "pause_my_timer"
        .Action = "pause"
        .Value = 2000   ' Pause for 2 seconds
    End With
End With
```

## Property Descriptions

### StartRunning:

type: `boolean` (`True`/`False`). Default: `False`

Controls whether this timer starts running, or whether it
needs to be started with one of the start control events.

### Direction:

type: `string` (`up`/`down`). Default: `up`

Controls which direction this timer runs in. Options are `up` or `down`.

### start_value:

Single value, type: `integer` or `template`
([Instructions for entering templates](dynamic_values.md)). Defaults to `0`.

The initial value of the timer. When a timer is restarted, this is the
value it will always start from. If you ever need to change the value,
you can use a jump control action to set it to whatever value you want.

### end_value:

Single value, type: `integer` or `template`
([Instructions for entering templates](dynamic_values.md)). Defaults to empty.

Specifies what the final value for this timer will be. When the timer
value equals or exceeds this (for timers counting up), or when it equals
or is lower than this (for timers counting down), the
*timer_\(name\)_complete* event is posted and the timer is stopped.
(If the `restart_on_complete:` setting is true, then the timer is also
reset back to its `start_value:` and started again.)

### tick_interval:

Single value, type: `integer (milliseconds)`

### Control Events
- `ControlEvents`: Configure actions (add, subtract, jump, start, stop, reset, restart,pause, set_tick_internal, change_tick_interval, reset_tick_interval) that can be triggered by events

### Debug Settings
- `Debug`: Boolean to enable debug logging for this timer (Default: False)

## Timer Events
The timer system generates the following events

- `timer_NAME_started`: Fired when the timer starts
- `timer_NAME_stopped`: Fired when the timer stops
- `timer_NAME_complete`: Fired when the timer reaches its end value
- `timer_NAME_tick`: Fired on each timer tick
