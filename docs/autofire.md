# Auto-Fire Configuration

The auto-fire configuration allows you to set up and customize auto-firing devices in your pinball machine. Auto-fire devices are used to automatically fire coils or other devices when triggered by switches or events. Examples include Pop Bumpers, Slingshots

## Configuration Options

### Basic Configuration
```vbscript
' Basic auto-fire configuration
With CreateGlfAutoFireDevice("left_sling")
    .Switch = "s_LeftSlingshot"
    .ActionCallback = "LeftSlingshotAction"
    .DisabledCallback = "LeftSlingshotDisabled"
    .EnabledCallback = "LeftSlingshotEnabled"
    .DisableEvents = Array("kill_flippers")
    .EnableEvents = Array("ball_started","enable_flippers")
    .ExcludeFromBallSearch = False
End With
```

## Property Descriptions

### Basic Settings
- `Switch`: String name of the switch that triggers the auto-fire (Default: Empty)
- `ActionCallback`: String name of the callback function that executes when the auto-fire is triggered (Default: Empty)
- `EnabledCallback`: String name of the callback function that executes when the auto-fire is enabled (Default: Empty)
- `DisabledCallback`: String name of the callback function that executes when the auto-fire is disabled (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the auto-fire (Default: Array("ball_started"))
- `DisableEvents`: Array of event names that will disable the auto-fire (Default: Array("ball_will_end", "service_mode_entered"))

### Ball Search Settings
- `ExcludeFromBallSearch`: Boolean indicating whether to exclude this device from ball search operations (Default: False)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)

## Events

The auto-fire device system generates several events that you can listen to:

- `auto_fire_coil_{device_name}_activate`: Triggered when the auto-fire device is activated
- `auto_fire_coil_{device_name}_deactivate`: Triggered when the auto-fire device is deactivated

Replace `{device_name}` with your actual auto-fire device name.

## Default Behavior

By default, auto-fire devices are configured with:
- Enabled on `ball_started` event
- Disabled on `ball_will_end` and `service_mode_entered` events
- Debug logging disabled
