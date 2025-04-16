# Ball Save Configuration

The ball save configuration allows you to set up and customize ball saving mechanisms in your pinball machine. Ball saves are used to prevent players from losing balls too quickly and can be configured to work within specific game modes.

## Configuration Options

### Basic Configuration
```vbscript
' Basic ball save configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .BallSaves("ball_save_name")
        .ActiveTime = 6000                ' Time in milliseconds the ball save is active
        .EnableEvents = Array("event1")   ' Events that enable the ball save
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced ball save configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .BallSaves("ball_save_name")
        ' Timing settings
        .ActiveTime = 6000                ' Time in milliseconds the ball save is active
        .GracePeriod = 2000               ' Additional time after active time before disabling
        .HurryUpTime = 3000               ' Time before active time ends to enter hurry-up state
        
        ' Event settings
        .EnableEvents = Array("event1", "event2")           ' Events that enable the ball save
        .DisableEvents = Array("event3", "event4")          ' Events that disable the ball save
        .TimerStartEvents = Array("event5", "event6")       ' Events that start the ball save timer
        
        ' Ball save behavior
        .AutoLaunch = True                ' Whether to automatically launch the saved ball
        .BallsToSave = 2                  ' Number of balls to save before disabling
        
        ' Debug settings
        .Debug = True                     ' Enable debug logging for this ball save
    End With
End With
```

## Property Descriptions

### Timing Settings
- `ActiveTime`: Time in milliseconds the ball save is active (Default: Null)
- `GracePeriod`: Additional time after active time before disabling (Default: Null)
- `HurryUpTime`: Time before active time ends to enter hurry-up state (Default: Null)

### Event Control
- `EnableEvents`: Array of event names that will enable the ball save (Default: Empty Array)
- `DisableEvents`: Array of event names that will disable the ball save (Default: Empty Array)
- `TimerStartEvents`: Array of event names that will start the ball save timer (Default: Empty Array)

### Ball Save Behavior
- `AutoLaunch`: Boolean indicating whether to automatically launch the saved ball (Default: False)
- `BallsToSave`: Integer indicating the number of balls to save before disabling (Default: 1)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this ball save (Default: False)

## Example Configurations

### Basic Ball Save Example
```vbscript
' Basic ball save configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .BallSaves("new_ball")
        .ActiveTime = 6000
        .EnableEvents = Array("new_ball_active")
    End With
End With
```

### Advanced Ball Save Example
```vbscript
' Advanced ball save configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .BallSaves("multiball_save")
        ' Timing settings
        .ActiveTime = 10000
        .GracePeriod = 3000
        .HurryUpTime = 2000
        
        ' Event settings
        .EnableEvents = Array("multiball_started")
        .DisableEvents = Array("multiball_ended")
        .TimerStartEvents = Array("multiball_ball_launched")
        
        ' Ball save behavior
        .AutoLaunch = True
        .BallsToSave = 3
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Events

The ball save system generates several events that you can listen to:

- `{ball_save_name}_enabled`: Triggered when the ball save is enabled
- `{ball_save_name}_disabled`: Triggered when the ball save is disabled
- `{ball_save_name}_timer_start`: Triggered when the ball save timer starts
- `{ball_save_name}_grace_period`: Triggered when the ball save enters grace period
- `{ball_save_name}_hurry_up`: Triggered when the ball save enters hurry-up state
- `{ball_save_name}_saving_ball`: Triggered when a ball is being saved

Replace `{ball_save_name}` with your actual ball save name (prefixed with "ball_save_" internally).

## Default Behavior

By default, ball saves are configured with:
- No active time, grace period, or hurry-up time
- No enable, disable, or timer start events
- Auto launch disabled
- Balls to save set to 1
- Debug logging disabled

## Notes

- Ball saves are managed within the context of a mode
- The ball save system automatically handles ball draining and replacement
- Debug logging can be enabled to track ball save operations
- The system prevents multiple ball save operations from overlapping
- Ball saves can be configured to work with specific game modes or features
- The ball save timer can be started automatically or triggered by specific events
- Grace period and hurry-up states can be used to provide visual or audio cues to the player 