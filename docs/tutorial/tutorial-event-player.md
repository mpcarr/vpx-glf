# Tutorial - Event Player

The Event Player is a powerful component in the VPX Game Logic Framework that allows you to create complex event chains where one event can trigger multiple other events. This tutorial will show you how to use the Event Player to create sophisticated game logic with minimal code.

## What is the Event Player?

The Event Player listens for specific events and automatically dispatches other events in response. This is useful for:

- Creating reusable event patterns
- Simplifying complex event handling logic
- Responding to game state changes with multiple actions
- Creating conditional event chains based on game conditions

## Basic Event Playing

```vbscript
' Mode-specific event player configuration
With CreateGlfMode("shoot_here", 20)
    .StartEvents = Array("mode_shoot_here_started")
    .StopEvents = Array("mode_shoot_here_stopped")
    
    With .EventPlayer
        ' When the mode starts, reset the upper target
        .Add "mode_shoot_here_started", Array("cmd_upper_target_reset")
    End With
End With
```

In this example, when the "shoot_here" mode starts, the Event Player will automatically reset the upper target.

## Conditional Event Playing

The Event Player can also use conditions to determine when to dispatch events. This allows for more precise control over when events are played:

```vbscript
' Conditional event player configuration
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .EventPlayer
        ' Only start super bonus round and play rich show if score is over 10000
        .Add "mode_base_started{current_player.score>10000}", Array("start_mode_superbonusround", "play_show_richy_rich")
        
        ' Choose different battle modes based on achievement state
        .Add "start_mode_battle{device.achievements.ironthrone.state!='completed'}", Array("start_mode_choose_battle")
        .Add "start_mode_battle{device.achievements.ironthrone.state=='completed'}", Array("start_mode_victory_lap")
    End With
End With
```

In this example:
- The `start_mode_superbonusround` and `play_show_richy_rich` events will only be dispatched if the player's score is over 10,000 when the base mode starts
- The `start_mode_choose_battle` event will be dispatched if the "ironthrone" achievement is not completed
- The `start_mode_victory_lap` event will be dispatched if the "ironthrone" achievement is completed

You can also apply conditions to individual events within a list:

```vbscript
' Conditional events within a list
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .EventPlayer
        .Add "reenable_nonrecruit_modes", Array( _
            "start_mode_shadowbroker_base", _
            "start_mode_n7_assignments", _
            "start_mode_overlordlight{device.achievements.collectorship.state!='complete'}", _
            "start_mode_arrival{device.achievements.collectorship.state=='complete'}", _
            "start_mode_shopping{current_player.cash>=1000}" _
        )
    End With
End With
```

In this example:
- `start_mode_shadowbroker_base` and `start_mode_n7_assignments` will always be dispatched
- Either `start_mode_overlordlight` or `start_mode_arrival` will be dispatched, depending on the state of the "collectorship" achievement
- `start_mode_shopping` will only be dispatched if the player's cash is at least 1000

## Dynamic Values in Events

The Event Player supports dynamic values in event names and arguments, allowing for even more flexibility:

### Dynamic Event Names

You can include dynamic values in event names using parentheses:

```vbscript
' Dynamic event names
With CreateGlfMode("dynamo", 30)
    .StartEvents = Array("mode_dynamo_started")
    .StopEvents = Array("mode_dynamo_stopped")
    
    With .EventPlayer
        ' Use player variables in event names
        .Add "mode_dynamo_started", Array("play_dynamo_show_phase_(current_player.phase_name)")
        
        ' Use device states in event names
        .Add "mode_dynamo_started", Array("dynamo_started_with_state_(device.achievements.dynamo.state)")
        
        ' Use conditional expressions in event names
        .Add "mode_dynamo_started", Array("player_score_is_('high' if current_player.score > 10000 else 'low')")
    End With
End With
```

In this example:
- If `current_player.phase_name` is "attackwave", the event `play_dynamo_show_phase_attackwave` will be dispatched
- If the "dynamo" achievement is completed, the event `dynamo_started_with_state_completed` will be dispatched
- If the player's score is over 10,000, the event `player_score_is_high` will be dispatched, otherwise `player_score_is_low` will be dispatched

### Dynamic Event Arguments

You can also include dynamic values as arguments to events:

```vbscript
' Dynamic event arguments
With CreateGlfMode("dynamo", 30)
    .StartEvents = Array("mode_dynamo_started")
    .StopEvents = Array("mode_dynamo_stopped")
    
    With .EventPlayer
        ' Set environment sounds with a specific name
        .Add "mode_dynamo_started", Array("set_environment_sounds:env_name=driving")
        
        ' Set initial laps count
        .Add "mode_dynamo_started", Array("set_initial_laps_count:count=10")
        
        ' Use dynamic values for event arguments
        .Add "mode_dynamo_started", Array("set_dynamo_phase:phase_name=(current_player.dynamo_phase)")
        
        ' Specify the type of dynamic values
        .Add "mode_dynamo_started", Array("set_dynamo_round:round_number=(device.counters.dynamo_rounds.value):type=int")
    End With
End With
```

In this example:
- The `set_environment_sounds` event will be dispatched with the argument `env_name=driving`
- The `set_initial_laps_count` event will be dispatched with the argument `count=10`
- The `set_dynamo_phase` event will be dispatched with the argument `phase_name` set to the value of `current_player.dynamo_phase`
- The `set_dynamo_round` event will be dispatched with the argument `round_number` set to the value of `device.counters.dynamo_rounds.value` as an integer

## Event Priority

You can specify the priority of events to control the order in which they are dispatched:

```vbscript
' Event priority
With CreateGlfMode("dynamo", 30)
    .StartEvents = Array("mode_dynamo_started")
    .StopEvents = Array("mode_dynamo_stopped")
    
    With .EventPlayer
        ' Reset player variable with high priority
        .Add "mode_dynamo_started", Array("reset_pv_tokens_collected_to_0:priority=50")
        
        ' Play slide with lower priority (after the player variable is reset)
        .Add "mode_dynamo_started", Array("play_slide{current_player.pv_tokens_collected <= 5}:priority=5:slide=dynamo_collect_more_tokens_slide")
    End With
End With
```

In this example:
- The `reset_pv_tokens_collected_to_0` event will be dispatched with a priority of 50
- The `play_slide` event will be dispatched with a priority of 5, ensuring it happens after the player variable is reset

## Practical Example: Multiball Mode

Let's put it all together with a practical example for a multiball mode:

```vbscript
' Multiball mode with event player
With CreateGlfMode("multiball", 40)
    .StartEvents = Array("multiball_start")
    .StopEvents = Array("multiball_end", "ball_lost")
    
    With .EventPlayer
        ' When multiball starts, enable all flippers and autofire coils
        .Add "multiball_start", Array("cmd_flippers_enable", "cmd_autofire_coils_enable")
        
        ' Play a show based on the current multiball phase
        .Add "multiball_start", Array("play_show_multiball_phase_(current_player.multiball_phase)")
        
        ' Set the jackpot value based on the number of balls in play
        .Add "multiball_start", Array("set_jackpot_value:value=(1000 * device.ball_devices.playfield.balls):type=int")
        
        ' When a jackpot is collected, increase the player's score and play a show
        .Add "jackpot_collected", Array("add_score:points=(current_player.jackpot_value)", "play_show_jackpot_collected")
        
        ' When multiball ends, disable all flippers and autofire coils
        .Add "multiball_end", Array("cmd_flippers_disable", "cmd_autofire_coils_disable")
        
        ' When a ball is lost during multiball, check if multiball should end
        .Add "ball_lost{device.ball_devices.playfield.balls <= 1}", Array("multiball_end")
    End With
End With
```

This example demonstrates:
- Basic event playing with multiple response events
- Dynamic event names based on player variables
- Dynamic event arguments with type specification
- Conditional event playing based on device state
- Event priority for controlling the order of operations

## Conclusion

The Event Player is a versatile and powerful component that can simplify your game logic and make your code more maintainable. By using the Event Player, you can:

1. Create reusable event patterns
2. Respond to game events with multiple actions
3. Create conditional event chains based on game conditions
4. Use dynamic values in event names and arguments
5. Control the priority of event dispatch

Remember that the Event Player is managed within the context of a mode, so you can have different event player configurations for different modes. This allows you to create complex and sophisticated game logic with minimal code. 