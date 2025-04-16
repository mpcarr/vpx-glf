# DOF Player Configuration

The DOF player configuration allows you to set up and customize Direct Output Framework (DOF) events in your pinball machine. DOF players are used to trigger physical feedback effects like solenoids, motors, and other physical outputs.

## Configuration Options

### Basic Configuration
```vbscript
' Basic DOF player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .DofPlayer
        With .EventName("dof_event")
            .DOFEvent = 123  ' DOF event number
            .Action = "DOF_ON"  ' Action to perform
        End With
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced DOF player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .DofPlayer
        ' Configure debug
        .Debug = True
        
        ' Configure DOF events
        With .EventName("flipper_event")
            .DOFEvent = 10
            .Action = "DOF_PULSE"
        End With
        
        With .EventName("bumper_event")
            .DOFEvent = 20
            .Action = "DOF_ON"
        End With
        
        With .EventName("motor_event")
            .DOFEvent = 30
            .Action = "DOF_OFF"
        End With
    End With
End With
```

## Property Descriptions

### DOF Event Settings
- `DOFEvent`: The DOF event number to trigger (Default: Empty)
- `Action`: The action to perform on the DOF event (Default: Empty)
  - `DOF_OFF`: Turn off the output (value: 0)
  - `DOF_ON`: Turn on the output (value: 1)
  - `DOF_PULSE`: Pulse the output (value: 2)
- `Debug`: Boolean to enable debug logging for this DOF player (Default: False)

## Example Configurations

### Basic DOF Player Example
```vbscript
' Basic DOF player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .DofPlayer
        With .EventName("flipper_event")
            .DOFEvent = 10
            .Action = "DOF_PULSE"
        End With
    End With
End With
```

### Advanced DOF Player Example
```vbscript
' Advanced DOF player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .DofPlayer
        ' Configure debug
        .Debug = True
        
        ' Configure DOF events
        With .EventName("jackpot_event")
            .DOFEvent = 100
            .Action = "DOF_PULSE"
        End With
        
        With .EventName("special_event")
            .DOFEvent = 200
            .Action = "DOF_ON"
        End With
        
        With .EventName("reset_event")
            .DOFEvent = 300
            .Action = "DOF_OFF"
        End With
    End With
End With
```

## DOF Player System

The DOF player system manages physical output events with the following features:

- DOF events can be configured with different actions (ON, OFF, PULSE)
- DOF players are managed within the context of a mode
- Events are triggered when the corresponding event is fired

## Events

The DOF player system responds to events by triggering the configured DOF events with the specified actions.

## Default Behavior

By default, DOF players are configured with:
- No events defined
- Debug logging disabled

## Notes

- DOF players are managed within the context of a mode
- The DOF player system automatically handles event activation and deactivation
- Debug logging can be enabled to track DOF player operations
- DOF players can be configured to work with specific game modes or features
- The `DOFEvent` property specifies which DOF event to trigger
- The `Action` property determines what action to take on the DOF event
- DOF events are triggered when the corresponding event is fired
- The system uses the VPX DOF function to trigger physical outputs 