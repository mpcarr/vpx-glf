# Shot Configuration

> For a step-by-step guide on how to configure shots, see the [Shots Tutorial](../tutorial/tutorial-shots.md).

The shot configuration allows you to set up and customize individual shots in your pinball machine. Shots are the primary interactive elements that players can hit to score points and trigger game events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic shot configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Shots("shot_name")
        .Switch = "switch_name"
        .Profile = "profile_name"
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced shot configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Shots("shot_name")
        ' Configure switches
        .Switches = Array("switch1", "switch2", "switch3")
        
        ' Configure profile
        .Profile = "profile_name"
        
        ' Configure enable events
        .EnableEvents = Array("enable_event1", "enable_event2")
        
        ' Configure disable events
        .DisableEvents = Array("disable_event1", "disable_event2")
        
        ' Configure advance events
        .AdvanceEvents = Array("advance_event1", "advance_event2")
        
        ' Configure hit events
        .HitEvents = Array("hit_event1", "hit_event2")
        
        ' Configure reset events
        .ResetEvents = Array("reset_event1", "reset_event2")
        
        ' Configure restart events
        .RestartEvents = Array("restart_event1", "restart_event2")
        
        ' Configure control events
        With .ControlEvents
            .Events = Array("control_event1", "control_event2")
            .State = 1
            .Force = True
            .ForceShow = False
        End With
        
        ' Configure show tokens
        With .Tokens
            .Add "token1", "value1"
            .Add "token2", "value2"
        End With
        
        ' Configure start enabled
        .StartEnabled = True
        
        ' Configure persist
        .Persist = True
        
        ' Debug settings
        .Debug = True      ' Enable debug logging for this shot
    End With
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: Single switch name that triggers the shot
- `Switches`: Array of switch names that trigger the shot
- `Profile`: Name of the shot profile to use (Default: "default")
- `StartEnabled`: Boolean to determine if the shot starts enabled (Default: Empty)
- `Persist`: Boolean to determine if the shot state persists between balls (Default: True)

### Event Control
- `EnableEvents`: Array of events that will enable the shot
- `DisableEvents`: Array of events that will disable the shot
- `AdvanceEvents`: Array of events that will advance the shot
- `HitEvents`: Array of events that will trigger the shot
- `ResetEvents`: Array of events that will reset the shot
- `RestartEvents`: Array of events that will restart the shot
- `ControlEvents`: Object that defines events that will control the shot
  - `Events`: Array of events that will control the shot
  - `State`: State to jump to when control events are triggered
  - `Force`: Boolean to determine if the shot should be forced to the state
  - `ForceShow`: Boolean to determine if the show should be forced to play

### Show Settings
- `Tokens`: Dictionary of tokens to pass to the show

### Debug Settings
- `Debug`: Boolean to enable debug logging for this shot (Default: False)

## Example Configurations

### Basic Shot Example
```vbscript
' Basic shot configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Shots("left_ramp")
        .Switch = "sw_left_ramp"
        .Profile = "ramp_profile"
    End With
End With
```

### Advanced Shot Example
```vbscript
' Advanced shot configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Shots("jackpot")
        ' Configure switches
        .Switches = Array("sw_left_ramp", "sw_right_ramp", "sw_center_ramp")
        
        ' Configure profile
        .Profile = "jackpot_profile"
        
        ' Configure enable events
        .EnableEvents = Array("multiball_start")
        
        ' Configure disable events
        .DisableEvents = Array("multiball_end", "ball_lost")
        
        ' Configure advance events
        .AdvanceEvents = Array("advance_jackpot")
        
        ' Configure hit events
        .HitEvents = Array("hit_jackpot")
        
        ' Configure reset events
        .ResetEvents = Array("reset_jackpot")
        
        ' Configure restart events
        .RestartEvents = Array("restart_jackpot")
        
        ' Configure control events
        With .ControlEvents
            .Events = Array("control_jackpot")
            .State = 2
            .Force = True
            .ForceShow = False
        End With
        
        ' Configure show tokens
        With .Tokens
            .Add "color", "red"
            .Add "intensity", "high"
        End With
        
        ' Configure start enabled
        .StartEnabled = True
        
        ' Configure persist
        .Persist = True
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Shot System

The shot system manages individual shots with the following features:

- Shots can be enabled, disabled, reset, or restarted
- Shots can advance to the next state or jump to a specific state
- Shots can be controlled by events
- Shots can have multiple switches
- Shots can have custom show tokens
- Shots can persist their state between balls
- Shots are managed within the context of a mode

## Events

The shot system generates the following events:

- `shot_name_hit`: Fired when the shot is hit
- `shot_name_profile_hit`: Fired when the shot is hit with the profile name
- `shot_name_profile_state_hit`: Fired when the shot is hit with the profile name and state
- `shot_name_state_hit`: Fired when the shot is hit with the state name

## Default Behavior

By default, shots are configured with:
- No switches
- Default profile
- No enable or disable events
- No advance events
- No hit events
- No reset or restart events
- No control events
- No show tokens
- Start enabled determined by enable events
- Persist enabled
- Debug logging disabled

## Notes

- Shots are managed within the context of a mode
- The shot system automatically handles state transitions
- Debug logging can be enabled to track shot operations
- Shots can be configured to work with specific game modes or features
- The shot system prevents events from conflicting with each other
- Control events can be used to create complex shot behavior
- Show tokens can be used to customize shows for different shots
- The persist feature can be used to maintain shot state between balls
- Multiple switches can be used to create complex shot triggers 