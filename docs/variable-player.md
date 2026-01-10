# Variable Player Configuration

The variable player configuration allows you to set up and customize variable manipulation in your game. Variable players are used to modify player and machine variables based on events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic variable player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .VariablePlayer
        With .EventName("event_name")
            With .Variable("variable_name")
                .Action = "add"
                .Int = 100
            End With
        End With
    End With
End With
```

## Property Descriptions

### Variable Actions
- `Action`: Action to perform on the variable
  - `add`: Add a value to a player variable
  - `add_machine`: Add a value to a machine variable
  - `set`: Set a player variable to a specific value
  - `set_machine`: Set a machine variable to a specific value

### Variable Types
- `Int`: Integer value
- `Float`: Floating-point value
- `String`: String value

### Debug Settings
- `Debug`: Boolean to enable debug logging for this variable player (Default: False)

## Machine Variables

Machine variables are persistent across games and can be created using the `CreateMachineVar` function:

```vbscript
' Create a machine variable
Dim high_score
Set high_score = CreateMachineVar("high_score")
high_score.InitialValue = 0
high_score.Persist = True
high_score.ValueType = "int"
```

## Variable Player System

The variable player system manages variable manipulation with the following features:

- Variables can be modified based on events
- Both player and machine variables can be manipulated
- Variables can be added to or set to specific values
- Variables are managed within the context of a mode

## Events

The variable player system doesn't generate its own events, but it responds to events defined in your configuration.

## Default Behavior

By default, variable players are configured with:
- No events defined
- Debug logging disabled

## Notes

- Variable players are managed within the context of a mode
- The variable player system automatically handles variable manipulation
- Debug logging can be enabled to track variable player operations
- Variable players can be configured to work with specific game modes or features
- The variable player system prevents variable conflicts
- Player variables are reset when a new game starts
- Machine variables persist across games if the `Persist` property is set to `True`
- The `add` action is useful for incrementing scores or counters
- The `set` action is useful for resetting or initializing variables 