# Mode Configuration

The mode configuration allows you to set up and customize game modes in your pinball machine. Modes are used to organize and manage different game states, features, and behaviors.

## Configuration Options

### Basic Configuration
```vbscript
' Basic mode configuration
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("event1", "event2")  ' Events that start the mode
    .StopEvents = Array("event3", "event4")   ' Events that stop the mode
End With
```

### Advanced Configuration
```vbscript
' Advanced mode configuration
With CreateGlfMode("mode_name", priority)
    ' Event settings
    .StartEvents = Array("event1", "event2")                    ' Events that start the mode
    .StopEvents = Array("event3", "event4")                     ' Events that stop the mode
    
    ' Mode settings
    .UseWaitQueue = True                                        ' Whether to wait for mode to stop before continuing
    
    ' Debug settings
    .Debug = True                                              ' Enable debug logging for this mode
End With
```

## Property Descriptions

### Basic Settings
- `StartEvents`: Array of event names that will start the mode (Default: Empty Array)
- `StopEvents`: Array of event names that will stop the mode (Default: Empty Array)
- `Priority`: Integer value determining the order of mode execution (Default: 0)

### Mode Behavior
- `UseWaitQueue`: Boolean indicating whether to wait for the mode to stop before continuing (Default: False)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this mode (Default: False)

## Available Components

A mode can include various components that are managed within its scope:

### Ball Management
- `BallSaves`: Ball save mechanisms
- `BallHolds`: Ball holding devices

### Game Elements
- `Counters`: Counting mechanisms
- `Timers`: Timing mechanisms
- `MultiballLocks`: Multiball lock mechanisms
- `Multiballs`: Multiball configurations
- `Shots`: Shot configurations
- `ShotGroups`: Groups of shots
- `SequenceShots`: Sequential shot patterns
- `ExtraBalls`: Extra ball configurations
- `ComboSwitches`: Combo switch configurations
- `TimedSwitches`: Timed switch configurations
- `Tilt`: Tilt mechanism
- `HighScore`: High score tracking

### Display and Feedback
- `LightPlayer`: Light control
- `ShowPlayer`: Show control
- `SegmentDisplayPlayer`: Segment display control
- `EventPlayer`: Event control
- `QueueEventPlayer`: Queued event control
- `QueueRelayPlayer`: Queue relay control
- `RandomEventPlayer`: Random event control
- `SoundPlayer`: Sound control
- `DOFPlayer`: Direct Output Framework control
- `SlidePlayer`: Slide control
- `VariablePlayer`: Variable control

### State Management
- `StateMachines`: State machine configurations
- `ShotProfiles`: Shot profile configurations

## Example Configurations

### Basic Mode Example
```vbscript
' Basic mode configuration
With CreateGlfMode("multiball", 10)
    .StartEvents = Array("multiball_start")
    .StopEvents = Array("multiball_end")
End With
```

### Advanced Mode Example
```vbscript
' Advanced mode configuration
With CreateGlfMode("special_mode", 20)
    ' Event settings
    .StartEvents = Array("special_mode_start")
    .StopEvents = Array("special_mode_end", "ball_end")
    .UseWaitQueue = True
    .Debug = True
    
    ' Event player for handling mode events
    With .EventPlayer()
        .Add "mode_special_mode_started", Array("stop_attract_mode", "enable_special_features")
        .Add "mode_special_mode_stopped", Array("disable_special_features")
    End With
    
    ' Sound player for mode-specific sounds
    With .SoundPlayer()
        With .EventName("mode_special_mode_started")
            .Key = "key_special_mode_music"
            .Sound = "special_mode_music"
        End With
        With .EventName("mode_special_mode_stopped")
            .Key = "key_special_mode_music"
            .Sound = "special_mode_music"
            .Action = "stop"
        End With
    End With
    
    ' Light player for mode-specific lighting
    With .LightPlayer()
        With .EventName("mode_special_mode_started")
            With .Lights("GI")
                .Color = "ff0000"  ' Red lighting
                .Fade = 200
            End With
        End With
        With .EventName("mode_special_mode_stopped")
            With .Lights("GI")
                .Color = "ffffff"  ' Normal lighting
                .Fade = 200
            End With
        End With
    End With
    
    ' Add a counter for tracking progress
    With .Counters("special_counter")
        .CountEvents = Array("special_target_hit")
        .CountCompleteValue = 5
        .EventsWhenComplete = Array("special_mode_complete")
    End With
    
    ' Add a timer for time-limited features
    With .Timers("special_timer")
        .TickInterval = 1000
        .StartValue = 30
        .EndValue = 0
        .StartRunning = True
        With .ControlEvents()
            .EventName = "mode_special_mode_started"
            .Action = "restart"
        End With
    End With
End With
```

## Events

The mode system generates several events that you can listen to:

- `{mode_name}_starting`: Triggered when the mode is about to start
- `{mode_name}_started`: Triggered when the mode has started
- `{mode_name}_stopping`: Triggered when the mode is about to stop
- `{mode_name}_stopped`: Triggered when the mode has stopped

Replace `{mode_name}` with your actual mode name (prefixed with "mode_" internally).

## Default Behavior

By default, modes are configured with:
- No start or stop events
- Priority of 0
- Wait queue disabled
- Debug logging disabled
- All components initialized but empty

## Notes

- Modes can be nested and have different priorities
- Components within a mode are automatically managed based on mode state
- Debug logging can be enabled to track mode operations
- The system prevents multiple mode operations from overlapping
- Modes can be used to organize complex game features and behaviors 