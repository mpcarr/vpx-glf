# Event Player Configuration

The event player configuration allows you to set up and customize event mapping and playback in your pinball machine. Event players are useful for creating complex event chains where one event can trigger multiple other events, with optional conditions.

## Configuration Options

### Basic Configuration
```vbscript
' Basic event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .EventPlayer
        .Add "trigger_event", Array("event1", "event2") ' When trigger_event occurs, play event1 and event2
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .EventPlayer
        ' Map multiple trigger events to their corresponding response events
        .Add "trigger_event1", Array("response_event1", "response_event2")
        .Add "trigger_event2", Array("response_event3", "response_event4")
        
        ' Debug settings
        .Debug = True ' Enable debug logging for this event player
    End With
End With
```

## Property Descriptions

### Event Mapping
- `Add(key, value)`: Maps a trigger event to an array of response events
  - `key`: The trigger event name
  - `value`: Array of response event names to be played when the trigger event occurs

### Debug Settings
- `Debug`: Boolean to enable debug logging for this event player (Default: False)

## Example Configurations

### Basic Event Player Example
```vbscript
' Basic event player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .EventPlayer
        .Add "ramp_complete", Array("ramp_award", "ramp_light")
    End With
End With
```

### Advanced Event Player Example
```vbscript
' Advanced event player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .EventPlayer
        ' Map multiple trigger events to their corresponding response events
        .Add "left_ramp_complete", Array("left_ramp_award", "left_ramp_light")
        .Add "right_ramp_complete", Array("right_ramp_award", "right_ramp_light")
        .Add "multiball_start", Array("multiball_light", "multiball_sound", "multiball_effect")
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Event Handling

The event player system works by:

1. Listening for trigger events specified in the configuration
2. When a trigger event occurs, it plays all the response events associated with that trigger
3. Events are played in the order they appear in the array
4. The event player automatically handles event registration and cleanup when the mode is activated or deactivated

## Events

The event player system doesn't generate its own events, but it listens for and responds to events defined in your configuration. When a trigger event occurs, the event player will automatically dispatch all the response events associated with that trigger.

## Default Behavior

By default, event players are configured with:
- No event mappings
- Debug logging disabled

## Notes

- Event players are managed within the context of a mode
- The event player system automatically handles event registration and cleanup
- Debug logging can be enabled to track event player operations
- Event players can be used to create complex event chains
- Event players can be configured to work with specific game modes or features
- The event player system is useful for creating reusable event patterns
- Event players can be used to simplify complex event handling logic 