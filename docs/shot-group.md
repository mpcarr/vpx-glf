# Shot Group Configuration

The shot group configuration allows you to group multiple shots together and manage them as a single unit in your pinball machine. Shot groups are useful for creating patterns of shots that can be rotated, enabled, disabled, or reset together.

## Configuration Options

### Basic Configuration
```vbscript
' Basic shot group configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShotGroup("group_name")
        .Shots = Array("shot1", "shot2", "shot3")
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced shot group configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShotGroup("group_name")
        ' Configure shots in the group
        .Shots = Array("shot1", "shot2", "shot3")
        
        ' Configure enable events
        .EnableEvents = Array("enable_event1", "enable_event2")
        
        ' Configure disable events
        .DisableEvents = Array("disable_event1", "disable_event2")
        
        ' Configure rotation events
        .RotateEvents = Array("rotate_event1", "rotate_event2")
        
        ' Configure left rotation events
        .RotateLeftEvents = Array("rotate_left_event1", "rotate_left_event2")
        
        ' Configure right rotation events
        .RotateRightEvents = Array("rotate_right_event1", "rotate_right_event2")
        
        ' Configure enable rotation events
        .EnableRotationEvents = Array("enable_rotation_event1", "enable_rotation_event2")
        
        ' Configure disable rotation events
        .DisableRotationEvents = Array("disable_rotation_event1", "disable_rotation_event2")
        
        ' Configure restart events
        .RestartEvents = Array("restart_event1", "restart_event2")
        
        ' Configure reset events
        .ResetEvents = Array("reset_event1", "reset_event2")
        
        ' Debug settings
        .Debug = True      ' Enable debug logging for this shot group
    End With
End With
```

## Property Descriptions

### Shot Settings
- `Shots`: Array of shot names that belong to the group

### Event Control
- `EnableEvents`: Array of events that will enable the shot group
- `DisableEvents`: Array of events that will disable the shot group
- `RotateEvents`: Array of events that will rotate the shot group
- `RotateLeftEvents`: Array of events that will rotate the shot group left
- `RotateRightEvents`: Array of events that will rotate the shot group right
- `EnableRotationEvents`: Array of events that will enable rotation for the shot group
- `DisableRotationEvents`: Array of events that will disable rotation for the shot group
- `RestartEvents`: Array of events that will restart the shot group
- `ResetEvents`: Array of events that will reset the shot group

### Debug Settings
- `Debug`: Boolean to enable debug logging for this shot group (Default: False)

## Example Configurations

### Basic Shot Group Example
```vbscript
' Basic shot group configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShotGroup("ramp_group")
        .Shots = Array("left_ramp", "right_ramp", "center_ramp")
    End With
End With
```

### Advanced Shot Group Example
```vbscript
' Advanced shot group configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .ShotGroup("jackpot_group")
        ' Configure shots in the group
        .Shots = Array("left_ramp", "right_ramp", "center_ramp")
        
        ' Configure enable events
        .EnableEvents = Array("multiball_start")
        
        ' Configure disable events
        .DisableEvents = Array("multiball_end", "ball_lost")
        
        ' Configure rotation events
        .RotateEvents = Array("rotate_jackpot")
        
        ' Configure left rotation events
        .RotateLeftEvents = Array("rotate_jackpot_left")
        
        ' Configure right rotation events
        .RotateRightEvents = Array("rotate_jackpot_right")
        
        ' Configure enable rotation events
        .EnableRotationEvents = Array("enable_jackpot_rotation")
        
        ' Configure disable rotation events
        .DisableRotationEvents = Array("disable_jackpot_rotation")
        
        ' Configure restart events
        .RestartEvents = Array("restart_jackpot")
        
        ' Configure reset events
        .ResetEvents = Array("reset_jackpot")
        
        ' Debug settings
        .Debug = True
    End With
End With
```

## Shot Group System

The shot group system manages groups of shots with the following features:

- Shots can be enabled, disabled, reset, or restarted as a group
- Shot states can be rotated within the group
- Rotation can be enabled or disabled
- Shot groups can track when all shots in the group reach a common state
- Shot groups are managed within the context of a mode

## Events

The shot group system generates the following events:

- `group_name_hit`: Fired when any shot in the group is hit
- `group_name_complete`: Fired when all shots in the group reach a common state
- `group_name_state_complete`: Fired when all shots in the group reach a specific state

## Default Behavior

By default, shot groups are configured with:
- No shots in the group
- No enable or disable events
- No rotation events
- No restart or reset events
- Rotation enabled
- Debug logging disabled

## Notes

- Shot groups are managed within the context of a mode
- The shot group system automatically handles shot state tracking
- Debug logging can be enabled to track shot group operations
- Shot groups can be configured to work with specific game modes or features
- The shot group system prevents shots from conflicting with each other
- Rotation can be used to create variety in shot patterns
- Shot groups can be used to create complex shot patterns that must be completed in a specific order
- The common state feature can be used to track when all shots in a group reach a specific state 