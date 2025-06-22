# Diverter Configuration

The diverter configuration allows you to set up and customize diverters in your pinball machine. Diverters are used to change the path of balls by activating mechanical devices.

## Configuration Options

```
With CreateGlfDiverter("lock_pin")
    .EnableEvents = Array("ball_started")
    .ActivateEvents = Array("raise_lock_pin")
    .DeactivateEvents = Array("drop_lock_pin")
    .ActivationTime = "3000"
    .ActionCallback = "DiverterLockPin"
End With
```

The *Action Callback* property needs implementing to actually move the diverter.

```
Sub DiverterLockPin(enabled)
	If enabled Then
		DiverterPin.isdropped=False
	Else
		DiverterPin.isdropped=True
	End If
End Sub
```

## Property Descriptions

### Basic Settings
- `ActionCallback`: String name of the callback function that executes when the diverter activates or deactivates (Default: Empty)

### Event Control
- `EnableEvents`: Array of event names that will enable the diverter (Default: Empty Array)
- `DisableEvents`: Array of event names that will disable the diverter (Default: Empty Array)
- `ActivateEvents`: Array of event names that will activate the diverter (Default: Empty Array)
- `DeactivateEvents`: Array of event names that will deactivate the diverter (Default: Empty Array)

### Timing Settings
- `ActivationTime`: Integer specifying how long to hold the diverter active in milliseconds (Default: 0)

### Switch Settings
- `ActivationSwitches`: Array of switch names that will activate the diverter when triggered (Default: Empty Array)

### Ball Search Settings
- `BallSearchHoldTime`: Integer specifying how long to hold the diverter during ball search in milliseconds (Default: 1000)
- `ExcludeFromBallSearch`: Boolean indicating whether to exclude this device from ball search operations (Default: False)

### Debug Settings
- `Debug`: Boolean to enable debug logging for this device (Default: False)


## Events

The diverter system generates several events that you can listen to:

- `diverter_{device_name}_activating`: Triggered when the diverter is being activated
- `diverter_{device_name}_deactivating`: Triggered when the diverter is being deactivated

Replace `{device_name}` with your actual diverter name.

## Default Behavior

By default, diverters are configured with:
- No enable/disable events
- No activation/deactivation events
- No activation switches
- Activation time of 0 milliseconds
- Ball search hold time of 1000 milliseconds
- Debug logging disabled
- Not excluded from ball search

## Notes

- The diverter system automatically manages enable/disable states based on configured events
- Action callbacks receive a parameter indicating the state (1 for activate, 0 for deactivate)
- Debug logging can be enabled to track diverter state changes
- Ball search functionality is available for troubleshooting
- Diverters can be configured to automatically deactivate after a specified time
- Activation switches provide direct control without requiring events
- The system prevents multiple activation/deactivation operations from overlapping 