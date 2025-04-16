# Drop Target Configuration

The drop target configuration allows you to set up and customize drop targets in your pinball machine. Drop targets are interactive targets that can be knocked down and reset, often used in banks or groups.

## Configuration Options

### Basic Configuration
```vbscript
' Basic drop target configuration
With CreateGlfDropTarget("droptarget_name")
    .Switch = "s_droptarget_switch"       ' Name of the switch that detects when the target is hit
    .Coil = "c_droptarget_coil"           ' Name of the coil to reset the target
    .ActionCallback = "DropTargetAction"  ' Name of the callback function for drop target actions
End With
```

### Advanced Configuration
```vbscript
' Advanced drop target configuration
With CreateGlfDropTarget("droptarget_name")
    ' Basic settings
    .Switch = "s_droptarget_switch"
    .Coil = "c_droptarget_coil"
    .ActionCallback = "DropTargetAction"
    
    ' Event settings
    .EnableEvents = Array("ball_started")                    ' Events that enable the drop target
    .DisableEvents = Array("ball_will_end", "service_mode")  ' Events that disable the drop target
    .ResetEvents = Array("event1", "event2")                 ' Events that reset the drop target
    
    ' Ball search settings
    .BallSearchEvents = Array("ball_search_started")         ' Events that trigger ball search
    .BallSearchDelay = 500                                   ' Delay before ball search reset
    
    ' Debug settings
    .Debug = True                                           ' Enable debug logging for this device
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: String name of the switch that detects when the target is hit (Default: Empty)
- `Coil`: String name of the coil to reset the target (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the drop target is hit or reset (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the drop target (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the drop target (Default: Array("ball_will_end", "service_mode_entered"))
- `ResetEvents`: Array of event names that will reset the drop target (Default: Array("machine_reset_phase_3", "ball_starting"))

### Ball Search Settings
- `BallSearchEvents`: Array of event names that will trigger ball search (Default: Array("ball_search_started"))
- `BallSearchDelay`: Integer specifying the delay in milliseconds before reset during ball search (Default: 500)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Example Configurations

### Basic Drop Target Example
```vbscript
' Basic drop target configuration
With CreateGlfDropTarget("target1")
    .Switch = "s_target1"
    .Coil = "c_target1"
    .ActionCallback = "TargetAction"
End With
```

### Advanced Drop Target Example
```vbscript
' Advanced drop target configuration
With CreateGlfDropTarget("bank1")
    .Switch = "s_bank1"
    .Coil = "c_bank1"
    .ActionCallback = "BankAction"
    .EnableEvents = Array("multiball_started")
    .DisableEvents = Array("multiball_ended")
    .ResetEvents = Array("bank_reset")
    .BallSearchEvents = Array("ball_search_started", "ball_search_ended")
    .BallSearchDelay = 750
    .Debug = True
End With
```

## Callback Examples

### Basic Drop Target Callback
```vbscript
' Basic drop target callback
Sub DropTargetAction(State)
    ' Code to handle drop target action
    If State = 1 Then
        ' Drop target is being hit
        PlaySound "target_sound"
    ElseIf State = 0 Then
        ' Drop target is being reset
        PlaySound "reset_sound"
    End If
End Sub
```

### Advanced Drop Target Callback
```vbscript
' Advanced drop target callback with additional logic
Sub DropTargetAction(State)
    If State = 1 Then
        ' Drop target is being hit
        PlaySound "target_sound"
        LightOn "l_target_hit"
        
        ' Add score
        AddScore 1000
        
        ' Check if all targets are down
        CheckAllTargetsDown
    ElseIf State = 0 Then
        ' Drop target is being reset
        PlaySound "reset_sound"
        LightOff "l_target_hit"
    End If
End Sub
```

## Events

The drop target system generates several events that you can listen to:

- `{droptarget_name}_hit`: Triggered when the drop target is hit
- `{droptarget_name}_resetting`: Triggered when the drop target is being reset
- `{droptarget_name}_reset`: Triggered when the drop target has been reset
- `{droptarget_name}_enabled`: Triggered when the drop target is enabled
- `{droptarget_name}_disabled`: Triggered when the drop target is disabled

Replace `{droptarget_name}` with your actual drop target name (prefixed with "droptarget_" internally).

## Default Behavior

By default, drop targets are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Reset on `machine_reset_phase_3` and `ball_starting` events
- Ball search triggered on `ball_search_started` event
- Ball search delay of 500 milliseconds
- Debug logging disabled

## Notes

- The drop target system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for hit, 0 for reset)
- Debug logging can be enabled to track drop target state changes
- Drop targets can be configured to reset during ball search with a configurable delay
- The system prevents multiple reset operations from overlapping
- Drop targets can be used individually or in banks for more complex gameplay 