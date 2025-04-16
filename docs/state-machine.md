# State Machine Configuration

The state machine configuration allows you to set up and customize state-based behavior in your pinball machine. State machines are used to manage complex game logic with multiple states and transitions.

## Configuration Options

### Basic Configuration
```vbscript
' Basic state machine configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .StateMachine("state_machine_name")
        ' Configure starting state
        .StartingState = "idle"
        
        ' Configure states
        With .States("idle")
            .Label = "Idle State"
        End With
        
        With .States("active")
            .Label = "Active State"
        End With
        
        ' Configure transitions
        With .Transitions
            .Source = Array("idle")
            .Target = "active"
            .Events = Array("start_event")
        End With
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced state machine configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .StateMachine("state_machine_name")
        ' Configure persistence
        .PersistState = True
        
        ' Configure starting state
        .StartingState = "idle"
        
        ' Configure debug
        .Debug = True
        
        ' Configure states
        With .States("idle")
            .Label = "Idle State"
            .EventsWhenStarted = Array("idle_state_started")
            .EventsWhenStopped = Array("idle_state_stopped")
            With .ShowWhenActive
                .Show = "idle_show"
                .Loops = 1
                .Speed = 1.0
            End With
        End With
        
        With .States("active")
            .Label = "Active State"
            .EventsWhenStarted = Array("active_state_started")
            .EventsWhenStopped = Array("active_state_stopped")
            With .ShowWhenActive
                .Show = "active_show"
                .Loops = 0
                .Speed = 1.0
            End With
        End With
        
        With .States("complete")
            .Label = "Complete State"
            .EventsWhenStarted = Array("complete_state_started")
            .EventsWhenStopped = Array("complete_state_stopped")
            With .ShowWhenActive
                .Show = "complete_show"
                .Loops = 1
                .Speed = 1.0
            End With
        End With
        
        ' Configure transitions
        With .Transitions
            .Source = Array("idle")
            .Target = "active"
            .Events = Array("start_event")
            .EventsWhenTransitioning = Array("transitioning_to_active")
        End With
        
        With .Transitions
            .Source = Array("active")
            .Target = "complete"
            .Events = Array("complete_event")
            .EventsWhenTransitioning = Array("transitioning_to_complete")
        End With
        
        With .Transitions
            .Source = Array("complete")
            .Target = "idle"
            .Events = Array("reset_event")
            .EventsWhenTransitioning = Array("transitioning_to_idle")
        End With
    End With
End With
```

## Property Descriptions

### State Machine Settings
- `StartingState`: The state to start in when the state machine is activated (Default: "start")
- `PersistState`: Whether to persist the state between balls (Default: False)
- `Debug`: Boolean to enable debug logging for this state machine (Default: False)

### State Settings
- `Label`: A human-readable label for the state
- `EventsWhenStarted`: Array of events to fire when the state starts
- `EventsWhenStopped`: Array of events to fire when the state stops
- `ShowWhenActive`: Show to play when the state is active
- `InternalCacheId`: Internal cache ID for the state

### Transition Settings
- `Source`: Array of source states that can transition to the target state
- `Target`: The target state to transition to
- `Events`: Array of events that trigger the transition
- `EventsWhenTransitioning`: Array of events to fire when transitioning

## Example Configurations

### Basic State Machine Example
```vbscript
' Basic state machine configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .StateMachine("ramp_state")
        ' Configure starting state
        .StartingState = "idle"
        
        ' Configure states
        With .States("idle")
            .Label = "Idle State"
        End With
        
        With .States("active")
            .Label = "Active State"
        End With
        
        ' Configure transitions
        With .Transitions
            .Source = Array("idle")
            .Target = "active"
            .Events = Array("ramp_enter")
        End With
        
        With .Transitions
            .Source = Array("active")
            .Target = "idle"
            .Events = Array("ramp_exit")
        End With
    End With
End With
```

### Advanced State Machine Example
```vbscript
' Advanced state machine configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .StateMachine("jackpot_state")
        ' Configure persistence
        .PersistState = True
        
        ' Configure starting state
        .StartingState = "idle"
        
        ' Configure debug
        .Debug = True
        
        ' Configure states
        With .States("idle")
            .Label = "Idle State"
            .EventsWhenStarted = Array("jackpot_idle")
            With .ShowWhenActive
                .Show = "jackpot_idle_show"
                .Loops = 1
                .Speed = 1.0
            End With
        End With
        
        With .States("collecting")
            .Label = "Collecting Jackpot"
            .EventsWhenStarted = Array("jackpot_collecting")
            With .ShowWhenActive
                .Show = "jackpot_collecting_show"
                .Loops = 0
                .Speed = 1.0
            End With
        End With
        
        With .States("complete")
            .Label = "Jackpot Complete"
            .EventsWhenStarted = Array("jackpot_complete")
            With .ShowWhenActive
                .Show = "jackpot_complete_show"
                .Loops = 1
                .Speed = 1.0
            End With
        End With
        
        ' Configure transitions
        With .Transitions
            .Source = Array("idle")
            .Target = "collecting"
            .Events = Array("jackpot_start")
            .EventsWhenTransitioning = Array("starting_jackpot")
        End With
        
        With .Transitions
            .Source = Array("collecting")
            .Target = "complete"
            .Events = Array("jackpot_collected")
            .EventsWhenTransitioning = Array("completing_jackpot")
        End With
        
        With .Transitions
            .Source = Array("complete")
            .Target = "idle"
            .Events = Array("jackpot_reset")
            .EventsWhenTransitioning = Array("resetting_jackpot")
        End With
    End With
End With
```

## State Machine System

The state machine system manages state-based behavior with the following features:

- States can be defined with labels, events, and shows
- Transitions can be defined between states
- States can fire events when started or stopped
- Transitions can fire events when transitioning
- Shows can be played when a state is active
- State machines can persist their state between balls
- State machines are managed within the context of a mode

## Events

The state machine system generates the following events:

- Events defined in `EventsWhenStarted` when a state starts
- Events defined in `EventsWhenStopped` when a state stops
- Events defined in `EventsWhenTransitioning` when transitioning between states

## Default Behavior

By default, state machines are configured with:
- Starting state of "start"
- No persistence between balls
- No states defined
- No transitions defined
- Debug logging disabled

## Notes

- State machines are managed within the context of a mode
- The state machine system automatically handles state transitions
- Debug logging can be enabled to track state machine operations
- State machines can be configured to work with specific game modes or features
- The state machine system prevents state conflicts
- State persistence is useful for maintaining state between balls
- Shows can be used to provide visual feedback for the current state
- Events can be used to trigger actions when states change
- Transitions can be triggered by multiple events
- Multiple source states can transition to the same target state 