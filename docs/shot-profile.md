# Shot Profile Configuration

The shot profile configuration allows you to define the behavior and states of shots in your pinball machine. Shot profiles define how shots advance, whether they loop, and how they handle rotation.

## Configuration Options

### Basic Configuration
```vbscript
' Basic shot profile configuration
With CreateGlfShotProfile("profile_name")
    ' Configure basic settings
    ' Configure states
    With .States("state1")
        .Show = "show_name"
        .Loops = 1
        .Speed = 1
        .SyncMs = 0
    End With
    
    With .States("state2")
        .Show = "show_name2"
        .Loops = 1
        .Speed = 1
        .SyncMs = 0
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced shot profile configuration
With CreateGlfShotProfile("profile_name")
    ' Configure basic settings
    .AdvanceOnHit = True
    .Block = False
    .ProfileLoop = True
    .RotationPattern = Array("r", "l", "r")
    
    ' Configure states to rotate
    .StateNamesToRotate = Array("state1", "state2")
    
    ' Configure states not to rotate
    .StateNamesNotToRotate = Array("state3", "state4")
    
    ' Configure states
    With .States("state1")
        .Show = "show_name"
        .Loops = 1
        .Speed = 1
        .SyncMs = 0
        .Key = "key_state1"
        
        ' Configure show tokens
        With .Tokens
            .Add "token1", "value1"
            .Add "token2", "value2"
        End With
    End With
    
    With .States("state2")
        .Show = "show_name2"
        .Loops = 2
        .Speed = 1.5
        .SyncMs = 100
        .Key = "key_state2"
        
        ' Configure show tokens
        With .Tokens
            .Add "token1", "value3"
            .Add "token2", "value4"
        End With
    End With
End With
```

## Property Descriptions

### Basic Settings
- `AdvanceOnHit`: Boolean to determine if the shot advances on hit (Default: True)
- `Block`: Boolean to determine if the shot blocks events (Default: False)
- `ProfileLoop`: Boolean to determine if the shot loops back to the first state (Default: False)
- `RotationPattern`: Array of rotation directions ("r" for right, "l" for left) (Default: Array("r"))

### State Settings
- `States`: Dictionary of states for the shot profile
  - `Show`: Name of the show to play for this state
  - `Loops`: Number of times to loop the show (Default: 1)
  - `Speed`: Speed multiplier for the show (Default: 1)
  - `SyncMs`: Synchronization time in milliseconds (Default: 0)
  - `Key`: Unique key for the state, used for show identification
  - `Tokens`: Dictionary of tokens to pass to the show

### Rotation Settings
- `StateNamesToRotate`: Array of state names that should be included in rotation
- `StateNamesNotToRotate`: Array of state names that should not be included in rotation

## Example Configurations

### Basic Shot Profile Example
```vbscript
' Basic shot profile configuration
With CreateGlfShotProfile("ramp_profile")
    ' Configure basic settings
    .AdvanceOnHit = True
    .Block = False
    .ProfileLoop = False
    .RotationPattern = Array("r")
    
    ' Configure states
    With .States("lit")
        .Show = "ramp_lit"
        .Loops = 1
        .Speed = 1.0
        .SyncMs = 0
        .Key = "key_ramp_lit"
    End With
    
    With .States("unlit")
        .Show = "ramp_unlit"
        .Loops = 1
        .Speed = 1.0
        .SyncMs = 0
        .Key = "key_ramp_unlit"
    End With
End With
```

### Advanced Shot Profile Example
```vbscript
' Advanced shot profile configuration
With CreateGlfShotProfile("jackpot_profile")
    ' Configure basic settings
    .AdvanceOnHit = True
    .Block = False
    .ProfileLoop = True
    .RotationPattern = Array("r", "l", "r")
    
    ' Configure states to rotate
    .StateNamesToRotate = Array("lit", "super_lit")
    
    ' Configure states not to rotate
    .StateNamesNotToRotate = Array("unlit", "completed")
    
    ' Configure states
    With .States("unlit")
        .Show = "jackpot_unlit"
        .Loops = 1
        .Speed = 1
        .SyncMs = 0
        .Key = "key_jackpot_unlit"
    End With
    
    With .States("lit")
        .Show = "jackpot_lit"
        .Loops = 1
        .Speed = 1
        .SyncMs = 0
        .Key = "key_jackpot_lit"
        
        ' Configure show tokens
        With .Tokens
            .Add "color", "red"
        End With
    End With
    
    With .States("super_lit")
        .Show = "jackpot_super_lit"
        .Loops = 2
        .Speed = 1
        .SyncMs = 100
        .Key = "key_jackpot_super_lit"
        
        ' Configure show tokens
        With .Tokens
            .Add "color", "blue"
            .Add "intensity", "high"
        End With
    End With
    
    With .States("completed")
        .Show = "jackpot_completed"
        .Loops = 1
        .Speed = 1
        .SyncMs = 0
        .Key = "key_jackpot_completed"
    End With
End With
```

### Double Scoring Profile Example
```vbscript
' Double scoring profile configuration
With CreateGlfShotProfile("double_scoring")
    ' Configure basic settings
    .AdvanceOnHit = True
    .Block = False
    .ProfileLoop = False
    
    ' Configure states
    With .States("unlit")
        .Show = "off"
        .Key = "key_off_ds"
        With .Tokens
            .Add "lights", "LDS"
            .Add "color", DoubleScoringColor
        End With
    End With
    
    With .States("flashing")
        .Show = "flash_color_with_fade"
        .Key = "key_flashing_ds"
        .Speed = 2
        With .Tokens
            .Add "lights", "LDS"
            .Add "fade", 500
            .Add "color", DoubleScoringColor
        End With
    End With
    
    With .States("hurry")
        .Show = "flash_color"
        .Key = "key_hurry_ds"
        .Speed = 7
        With .Tokens
            .Add "lights", "LDS"
            .Add "color", DoubleScoringColor
        End With
    End With
End With
```

## Shot Profile System

The shot profile system manages shot behavior with the following features:

- Shots can advance on hit or remain in the same state
- Shots can loop back to the first state or stop at the last state
- Shots can block events or allow them to pass through
- Shots can have multiple states with different shows and behaviors
- Shots can be rotated according to a pattern
- Some states can be excluded from rotation
- Each state can have a unique key for identification
- States can pass tokens to shows for customization

## Events

The shot profile system doesn't generate its own events, but it defines the behavior of shots that use the profile.

## Default Behavior

By default, shot profiles are configured with:
- Advance on hit enabled
- Block disabled
- Loop disabled
- Right rotation pattern
- No states defined
- No states to rotate or not to rotate
- No keys defined for states
- No tokens defined for states

## Notes

- Shot profiles are used by shots to define their behavior
- The shot profile system automatically handles state transitions
- Shot profiles can be configured to work with specific game modes or features
- The shot profile system prevents states from conflicting with each other
- Rotation patterns can be used to create variety in shot behavior
- Show tokens can be used to customize shows for different states
- State synchronization can be used to coordinate multiple shots
- The loop feature can be used to create continuous shot patterns
- Tokens can be used to pass dynamic values to shows, such as colors or light groups 