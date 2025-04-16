# Sequence Shot Configuration

The sequence shot configuration allows you to set up and customize sequential shot patterns in your pinball machine. Sequence shots track a series of events or switch hits that must occur in a specific order to complete a sequence.

## Configuration Options

### Basic Configuration
```vbscript
' Basic sequence shot configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SequenceShot("sequence_name")
        .EventSequence = Array("event1", "event2", "event3")
        .SequenceTimeout = 5000    ' Timeout in milliseconds
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced sequence shot configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SequenceShot("sequence_name")
        ' Configure event sequence
        .EventSequence = Array("event1", "event2", "event3")
        
        ' Configure switch sequence (alternative to event sequence)
        .SwitchSequence = Array("switch1", "switch2", "switch3")
        
        ' Configure sequence timeout
        .SequenceTimeout = 5000    ' Timeout in milliseconds
        
        ' Configure cancel events
        .CancelEvents = Array("cancel_event1", "cancel_event2")
        
        ' Configure cancel switches
        .CancelSwitches = Array("cancel_switch1", "cancel_switch2")
        
        ' Configure delay events
        .DelayEventList = Array("delay_event1", "delay_event2")
        
        ' Configure delay switches
        .DelaySwitchList = Array("delay_switch1", "delay_switch2")
        
        ' Debug settings
        .Debug = True      ' Enable debug logging for this sequence shot
    End With
End With
```

## Property Descriptions

### Sequence Settings
- `EventSequence`: Array of event names that must occur in sequence
- `SwitchSequence`: Array of switch names that must be hit in sequence
- `SequenceTimeout`: Time in milliseconds before the sequence times out (Default: 0)

### Cancel Settings
- `CancelEvents`: Array of events that will cancel the sequence
- `CancelSwitches`: Array of switches that will cancel the sequence

### Delay Settings
- `DelayEventList`: Array of events that will delay the sequence
- `DelaySwitchList`: Array of switches that will delay the sequence

### Debug Settings
- `Debug`: Boolean to enable debug logging for this sequence shot (Default: False)

## Example Configurations

### Basic Sequence Shot Example
```vbscript
' Basic sequence shot configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SequenceShot("ramp_sequence")
        .EventSequence = Array("left_ramp", "right_ramp", "center_ramp")
        .SequenceTimeout = 10000    ' 10 seconds to complete the sequence
    End With
End With
```

### Advanced Sequence Shot Example
```vbscript
' Advanced sequence shot configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SequenceShot("jackpot_sequence")
        ' Configure event sequence
        .EventSequence = Array("left_ramp", "right_ramp", "center_ramp")
        
        ' Configure cancel events
        .CancelEvents = Array("ball_lost", "multiball_end")
        
        ' Configure delay events
        .DelayEventList = Array("special_mode_start")
        
        ' Configure sequence timeout
        .SequenceTimeout = 15000    ' 15 seconds to complete the sequence
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Sequence System

The sequence shot system manages sequential patterns with the following features:

- Events or switches must occur in a specific order
- Sequences can be canceled by specific events or switches
- Sequences can be delayed by specific events or switches
- Sequences can have a timeout period
- Multiple sequences can be tracked simultaneously
- Sequences are cleared when the mode stops

## Events

The sequence shot system generates the following events:

- `sequence_name_hit`: Fired when a sequence is completed
- `sequence_name_timeout`: Fired when a sequence times out

## Default Behavior

By default, sequence shots are configured with:
- No events or switches in the sequence
- No timeout period
- No cancel events or switches
- No delay events or switches
- Debug logging disabled

## Notes

- Sequence shots are managed within the context of a mode
- The sequence system automatically handles event tracking and timing
- Debug logging can be enabled to track sequence shot operations
- Sequence shots can be configured to work with specific game modes or features
- The sequence system prevents events from overlapping or conflicting
- Timeouts can be used to create urgency in sequence completion
- Cancel events can be used to reset sequences when certain conditions are met
- Delay events can be used to pause sequences during special modes or events 