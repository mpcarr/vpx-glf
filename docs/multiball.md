# Multiball Configuration

The multiball configuration allows you to set up and customize multiball modes in your pinball machine. Multiball modes are special gameplay features where multiple balls are in play simultaneously, creating exciting and challenging gameplay.

## Configuration Options

### Basic Configuration
```vbscript
' Basic multiball configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Multiball("multiball_name")
        .BallCount = 3                ' Number of balls in multiball
        .StartEvents = Array("event1") ' Events that trigger multiball start
        .StopEvents = Array("event2")  ' Events that trigger multiball stop
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced multiball configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Multiball("multiball_name")
        ' Ball settings
        .BallCount = 3                ' Number of balls in multiball
        .BallCountType = "total"      ' "total" or "additional" (Default: "total")
        .BallLock = "lock_device"     ' Ball device to eject balls from
        
        ' Event settings
        .StartEvents = Array("event1", "event2") ' Events that trigger multiball start
        .StopEvents = Array("event3", "event4")  ' Events that trigger multiball stop
        .ResetEvents = Array("event5", "event6") ' Events that reset multiball
        .AddABallEvents = Array("event7", "event8") ' Events that add a ball to multiball
        
        ' Timing settings
        .ShootAgain = 10000           ' Time in milliseconds for shoot again (Default: 10000)
        .GracePeriod = 2000           ' Grace period in milliseconds after shoot again ends (Default: 0)
        .HurryUp = 5000               ' Time in milliseconds before hurry up warning (Default: 0)
        .AddABallShootAgain = 5000    ' Time in milliseconds for add a ball shoot again (Default: 5000)
        .AddABallGracePeriod = 1000   ' Grace period in milliseconds for add a ball (Default: 0)
        .AddABallHurryUpTime = 2000   ' Time in milliseconds before add a ball hurry up warning (Default: 0)
        
        ' Debug settings
        .Debug = True                 ' Enable debug logging for this multiball
    End With
End With
```

## Property Descriptions

### Ball Settings
- `BallCount`: Number of balls in multiball (Default: 0)
- `BallCountType`: Type of ball count - "total" or "additional" (Default: "total")
- `BallLock`: Ball device to eject balls from (Default: Empty)

### Event Control
- `StartEvents`: Array of event names that trigger multiball start (Default: Empty Array)
- `StopEvents`: Array of event names that trigger multiball stop (Default: Empty Array)
- `ResetEvents`: Array of event names that reset multiball (Default: Empty Array)
- `AddABallEvents`: Array of event names that add a ball to multiball (Default: Empty Array)
- `EnableEvents`: Array of event names that enable multiball (Default: Empty Array)
- `DisableEvents`: Array of event names that disable multiball (Default: Empty Array)

### Timing Settings
- `ShootAgain`: Time in milliseconds for shoot again (Default: 10000)
- `GracePeriod`: Grace period in milliseconds after shoot again ends (Default: 0)
- `HurryUp`: Time in milliseconds before hurry up warning (Default: 0)
- `AddABallShootAgain`: Time in milliseconds for add a ball shoot again (Default: 5000)
- `AddABallGracePeriod`: Grace period in milliseconds for add a ball (Default: 0)
- `AddABallHurryUpTime`: Time in milliseconds before add a ball hurry up warning (Default: 0)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this multiball (Default: False)

## Example Configurations

### Basic Multiball Example
```vbscript
' Basic multiball configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Multiball("basic_multiball")
        .BallCount = 3
        .StartEvents = Array("multiball_start")
        .StopEvents = Array("multiball_stop")
    End With
End With
```

### Advanced Multiball Example
```vbscript
' Advanced multiball configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .Multiball("jackpot_multiball")
        ' Ball settings
        .BallCount = 3
        .BallCountType = "total"
        .BallLock = "lock_device"
        
        ' Event settings
        .StartEvents = Array("jackpot_ready")
        .StopEvents = Array("jackpot_collected", "ball_ended")
        .ResetEvents = Array("multiball_reset")
        .AddABallEvents = Array("extra_ball_awarded")
        
        ' Timing settings
        .ShootAgain = 15000
        .GracePeriod = 3000
        .HurryUp = 5000
        .AddABallShootAgain = 8000
        .AddABallGracePeriod = 2000
        .AddABallHurryUpTime = 3000
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Multiball States

The multiball system manages several states:

- `enabled`: Multiball is enabled and ready to start
- `shoot_again_enabled`: Shoot again is active
- `grace_period_enabled`: Grace period is active
- `hurry_up_enabled`: Hurry up warning is active

## Events

The multiball system generates several events that you can listen to:

- `{multiball_name}_started`: Triggered when multiball starts
- `{multiball_name}_ended`: Triggered when multiball ends
- `{multiball_name}_reset_event`: Triggered when multiball is reset
- `{multiball_name}_shoot_again`: Triggered when a ball is saved during shoot again
- `{multiball_name}_shoot_again_ended`: Triggered when shoot again ends
- `{multiball_name}_grace_period`: Triggered when grace period starts
- `{multiball_name}_hurry_up`: Triggered when hurry up warning starts
- `{multiball_name}_ball_lost`: Triggered when a ball is lost during multiball

Replace `{multiball_name}` with your actual multiball name (prefixed with "multiball_" internally).

## Default Behavior

By default, multiballs are configured with:
- 0 balls
- "total" ball count type
- No ball lock device
- No start, stop, reset, or add a ball events
- 10000ms shoot again time
- 0ms grace period
- 0ms hurry up time
- 5000ms add a ball shoot again time
- 0ms add a ball grace period
- 0ms add a ball hurry up time
- Debug logging disabled

## Notes

- Multiballs are managed within the context of a mode
- The multiball system automatically handles ball tracking and shoot again
- Debug logging can be enabled to track multiball operations
- Multiballs can be configured to work with specific game modes or features
- The multiball system supports adding balls during gameplay
- Shoot again, grace period, and hurry up features can be customized
- Multiballs can be configured to eject balls from a ball lock device 