# Random Event Player Configuration

The random event player configuration allows you to set up and customize random event generation in your pinball machine. Random event players are used to generate random events based on configured probabilities and conditions.

## Configuration Options

### Basic Configuration
```vbscript
' Basic random event player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    

    With .RandomEventPlayer()
        With .EventName("s_left_sling_active")
            .Add "flash_left_sling_red", 1
            .Add "flash_right_sling_blue", 1
        End With
        With .EventName("s_right_sling_active")
            .Add "flash_right_sling_red", 1
            .Add "flash_right_sling_blue", 1
        End With
    End With
End With
```
