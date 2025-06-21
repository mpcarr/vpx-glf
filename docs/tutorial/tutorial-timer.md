# Timer Tutorial

Timers are essential for creating time-based events in your pinball game, such as mode durations, countdowns, and timed opportunities. The timer in the VPX Game Logic Framework is flexible and can be configured to count up or down, trigger events, and be controlled by other events.

## What is a Timer?
A timer is a component that counts up or down from a specified value, firing events as it runs. You can use timers to:
- Create time-limited modes (e.g., "Complete the objective in 30 seconds!")
- Track elapsed time for scoring or achievements
- Add urgency to gameplay

## Basic Timer Configuration

Here is a straightforward example of how to configure a timer for a "hurry up" mode:

```vbscript
With CreateGlfMode("hurry_up", 10)
    .StartEvents = Array("hurry_up_start")
    .StopEvents = Array("timer_hurry_up_complete")

    With .Timer("hurry_up_timer")
        .StartRunning = True
        .Direction = "down"      ' Count down
        .StartValue = 15         ' Start at 15
        .EndValue = 0            ' End at 0
        .TickInterval = 1000     ' Tick every 1 second (1000 ms)
    End With
End With
```

**How it works:**
- The timer is named `hurry_up_timer` and is part of the `hurry_up` mode.
- The mode starts when the `hurry_up_start` event occurs and stops when the timer completes (`timer_hurry_up_complete`).
- The timer starts running immediately, counts down from 15 to 0, ticking every second.
- When the timer reaches 0, it automatically fires a `timer_hurry_up_complete` event, which ends the mode.

## Timer Events
Timers automatically generate events you can use in your game logic:
- `<timer_name>_started` — when the timer starts
- `<timer_name>_stopped` — when the timer stops
- `<timer_name>_complete` — when the timer reaches its end value
- `<timer_name>_tick` — on every tick

You can use these events to trigger actions, such as ending a mode, playing a sound, or updating the display.

## Controlling the Timer
You can also control the timer using control events. For example, you can add time, pause, or change the tick interval dynamically by configuring `.ControlEvents`:

```vbscript
With .Timer("hurry_up_timer")
    ' ... basic config ...
    With .ControlEvents
        .EventName = "add_time"
        .Action = "add"
        .Value = 5000
    End With
    With .ControlEvents
        .EventName = "pause_timer"
        .Action = "pause"
        .Value = 2000  ' Pause for 2 seconds
    End With
End With
```

- Triggering the `add_time` event will add 5 to the timer.
- Triggering the `pause_timer` event will pause the timer for 2 seconds.

## Summary
- Timers are configured inside a mode using `.Timer("name")`.
- Set the direction, start/end values, and tick interval.
- Use events to start/stop the timer and respond to timer events.
- Use control events to dynamically modify the timer during gameplay.

This covers the basics of using timers in your game. For more advanced usage, see the [Timer Command Reference](../../timer). 