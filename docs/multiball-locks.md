# Multiball Locks Configuration

The multiball locks configuration allows you to set up and customize ball locking mechanisms for multiball modes in your pinball machine. Multiball locks are devices that capture and hold balls until they are released for multiball play.

## Configuration Options

### Basic Configuration
```vbscript
' Basic multiball lock configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .MultiballLock("lock_name")
        .LockDevice = "ball_device"   ' Ball device that acts as the lock
        .BallsToLock = 3              ' Number of balls that can be locked
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced multiball lock configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .MultiballLock("lock_name")
        ' Lock settings
        .LockDevice = "ball_device"   ' Ball device that acts as the lock
        .BallsToLock = 3              ' Number of balls that can be locked
        .BallsToReplace = 1           ' Number of balls to replace when locked (Default: -1, meaning all)
        
        ' Event settings
        .LockEvents = Array("event1", "event2") ' Events that trigger ball locking
        .ResetEvents = Array("event3", "event4") ' Events that reset the lock count
        
        ' Debug settings
        .Debug = True                 ' Enable debug logging for this lock
    End With
End With
```

## Property Descriptions

### Lock Settings
- `LockDevice`: Ball device that acts as the lock (Default: Empty)
- `BallsToLock`: Number of balls that can be locked (Default: 0)
- `BallsToReplace`: Number of balls to replace when locked (Default: -1, meaning all)

### Event Control
- `LockEvents`: Array of event names that trigger ball locking (Default: Empty Array)
- `ResetEvents`: Array of event names that reset the lock count (Default: Empty Array)
- `EnableEvents`: Array of event names that enable the lock (Default: Empty Array)
- `DisableEvents`: Array of event names that disable the lock (Default: Empty Array)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this lock (Default: False)

## Example Configurations

### Basic Multiball Lock Example
```vbscript
' Basic multiball lock configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .MultiballLock("basic_lock")
        .LockDevice = "lock_device"
        .BallsToLock = 3
    End With
End With
```

### Advanced Multiball Lock Example
```vbscript
' Advanced multiball lock configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .MultiballLock("jackpot_lock")
        ' Lock settings
        .LockDevice = "lock_device"
        .BallsToLock = 3
        .BallsToReplace = 1
        
        ' Event settings
        .LockEvents = Array("lock_ready")
        .ResetEvents = Array("multiball_reset")
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Lock Tracking

The multiball lock system tracks:
- The number of balls currently locked
- Whether the lock is enabled or disabled

## Events

The multiball lock system generates several events that you can listen to:

- `{lock_name}_locked_ball`: Triggered when a ball is locked, with the count of locked balls
- `{lock_name}_full`: Triggered when the lock is full, with the count of locked balls

Replace `{lock_name}` with your actual lock name (prefixed with "multiball_lock_" internally).

## Default Behavior

By default, multiball locks are configured with:
- No lock device
- 0 balls to lock
- -1 balls to replace (meaning all)
- No lock or reset events
- Debug logging disabled

## Notes

- Multiball locks are managed within the context of a mode
- The multiball lock system automatically tracks the number of balls locked
- Debug logging can be enabled to track lock operations
- Multiball locks can be configured to work with specific game modes or features
- The multiball lock system prevents exceeding the maximum number of balls that can be locked
- Multiball locks can be configured to replace a specific number of balls when locked
- The lock count is tracked per player using player state 