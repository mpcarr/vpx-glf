# Tutorial - Timer

Welcome to the Timer tutorial for Galaxy Quest Pinball! In this tutorial, we'll learn how to use the Timer component to create time-based events in your pinball game. Timers are essential for creating urgency, managing mode durations, and tracking elapsed time.

## What is a Timer?

A Timer in the VPX Game Logic Framework is a component that can count up or down from a specified value. Timers are useful for:

- Creating time-limited modes (like "Shoot the ramp within 10 seconds!")
- Managing bonus rounds with countdown timers
- Tracking elapsed time for scoring or achievements
- Creating urgency in gameplay

## Basic Timer Setup

Let's start with a simple example of how to create a timer in your Galaxy Quest Pinball game:

```vbscript
' Basic timer configuration within a mode
With CreateGlfMode("galaxy_quest_base", 10)
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

This configuration:
1. Creates a timer called "bonus_timer" in the "galaxy_quest_base" mode
2. Starts the timer when the "bonus_start" event occurs
3. Stops the timer when the "bonus_end" event occurs
4. Sets the timer to count down from 10 to 0

When the timer reaches 0, it will automatically fire a `bonus_timer_complete` event that you can use to end the bonus round or trigger other actions.

## Timer Events

Timers generate several events that you can listen for:

- `timer_name_started`: Fired when the timer starts
- `timer_name_stopped`: Fired when the timer stops
- `timer_name_complete`: Fired when the timer reaches its end value
- `timer_name_tick`: Fired on each timer tick

Let's use these events to create a more interactive bonus round:

```vbscript
' Bonus round with timer
With CreateGlfMode("galaxy_quest_bonus", 20)
    .StartEvents = Array("bonus_start")
    .StopEvents = Array("bonus_end", "bonus_timer_complete")
    
    With .Timer("bonus_timer")
        .StartEvents = Array("bonus_start")
        .StopEvents = Array("bonus_end", "bonus_timer_complete")
        .Direction = "down"
        .StartValue = 30
        .EndValue = 0
        .TickInterval = 1000  ' Tick every second
    End With
    
    ' Use the Event Player to respond to timer events
    With .EventPlayer
        ' When the timer starts, play a show and enable special shots
        .Add "bonus_timer_started", Array("play_show_bonus_start", "enable_bonus_shots")
        
        ' When the timer completes, end the bonus round
        .Add "bonus_timer_complete", Array("play_show_bonus_end", "bonus_end")
    End With
End With
```

In this example:
1. We create a bonus round mode with a 30-second timer
2. When the timer starts, we play a show and enable special shots
3. When the timer completes, we play an end show and end the bonus round

### Timer Control

You can also control timers with events:

```vbscript
' Timer with control events
With CreateGlfMode("galaxy_quest_base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("jackpot_timer")
        ' Configure basic settings
        .Direction = "down"
        .StartValue = 10
        .EndValue = 0
        .TickInterval = 1000
        
        ' Configure control events
        With .ControlEvents
            .EventName = "add_time"
            .Action = "add"
            .Value = 5
        End With
        
        With .ControlEvents
            .EventName = "pause_timer"
            .Action = "pause"
            .Value = 2000  ' Pause for 2 seconds
        End With
        
        With .ControlEvents
            .EventName = "speed_up_timer"
            .Action = "change_tick_interval"
            .Value = 0.5  ' Halve the tick interval (speed up)
        End With
    End With
End With
```

In this example:
1. The `add_time` event will add 5 seconds to the timer
2. The `pause_timer` event will pause the timer for 2 seconds
3. The `speed_up_timer` event will make the timer tick twice as fast

### Restart on Complete

You can also configure a timer to restart automatically when it completes:

```vbscript
' Timer that restarts on complete
With CreateGlfMode("galaxy_quest_base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("cycling_timer")
        .StartEvents = Array("cycling_start")
        .StopEvents = Array("cycling_end")
        .Direction = "down"
        .StartValue = 5
        .EndValue = 0
        .RestartOnComplete = True  ' Restart when complete
    End With
End With
```

This timer will automatically restart when it reaches 0, creating a cycling effect that can be used for rotating shot lights or other cycling features.

## Practical Example: Time Attack Mode

Let's put it all together with a practical example for a "Time Attack" mode in Galaxy Quest Pinball:

```vbscript
' Time Attack mode with timer
With CreateGlfMode("galaxy_quest_time_attack", 30)
    .StartEvents = Array("time_attack_start")
    .StopEvents = Array("time_attack_end", "time_attack_timer_complete")
    
    With .Timer("time_attack_timer")
        .StartEvents = Array("mode_galaxy_quest_time_attack_started")
        .Direction = "down"
        .StartValue = 60  ' 60 seconds
        .EndValue = 0
        .TickInterval = 1000  ' Tick every second
        .RestartOnComplete = False  ' Don't restart when complete
    End With
    
    ' Use the Event Player to respond to timer events
    With .EventPlayer
        ' When the mode starts, play a show and enable special shots
        .Add "mode_galaxy_quest_time_attack_started", Array("play_show_time_attack_start", "enable_time_attack_shots")
        
        ' When the timer completes, end the mode
        .Add "time_attack_timer_complete", Array("play_show_time_attack_end", "time_attack_end")
        
        ' Add time when specific shots are hit
        .Add "left_ramp_hit", Array("add_time:value=5")
        .Add "right_ramp_hit", Array("add_time:value=5")
        .Add "center_ramp_hit", Array("add_time:value=10")
    End With
End With
```

This example demonstrates:
1. A 60-second time attack mode
2. Display updates on each tick
3. Urgency checks (you could flash lights or play sounds when time is running low)
4. The ability to add time by hitting specific shots
5. Automatic mode ending when time runs out

## Timer with Shot Integration

Timers work great with shots to create time-limited shot opportunities:

```vbscript
' Timer with shot integration
With CreateGlfMode("galaxy_quest_base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Timer("opportunity_timer")
        .StartEvents = Array("opportunity_start")
        .StopEvents = Array("opportunity_end", "opportunity_timer_complete")
        .Direction = "down"
        .StartValue = 5
        .EndValue = 0
        .TickInterval = 1000
    End With
    
    With .Shot("left_ramp_opportunity")
        .Switch = "sw_left_ramp"
        .Profile = "opportunity_profile"
        .EnableEvents = Array("opportunity_start")
        .DisableEvents = Array("opportunity_end", "opportunity_timer_complete")
    End With
    
    ' Use the Event Player to manage the opportunity
    With .EventPlayer
        .Add "opportunity_start", Array("play_show_opportunity_start")
        .Add "opportunity_timer_complete", Array("play_show_opportunity_end", "opportunity_end")
        .Add "left_ramp_opportunity_hit", Array("add_score:points=1000", "opportunity_end")
    End With
End With
```

In this example:
1. We create a 5-second opportunity timer
2. We configure a shot that is only enabled during the opportunity
3. When the opportunity starts, we play a show
4. When the opportunity ends (either by time running out or hitting the shot), we play an end show
5. If the player hits the shot during the opportunity, they score 1000 points

## Conclusion

Timers are a versatile and powerful component that can add excitement and urgency to your pinball game. By using timers, you can:

1. Create time-limited modes and opportunities
2. Track elapsed time for scoring or achievements
3. Create cycling effects for lights or shots
4. Add time-based bonuses and rewards
5. Create urgency in gameplay

Remember that timers are managed within the context of a mode, so you can have different timer configurations for different modes. This allows you to create complex and sophisticated time-based gameplay with minimal code.

## What's Next

In the next tutorial, we'll explore the Multiball component, which allows you to create exciting multiball modes with multiple balls in play. We'll learn how to configure multiball modes, manage ball locks, and create jackpot opportunities. 