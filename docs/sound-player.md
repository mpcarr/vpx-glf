# Sound Player Configuration

The sound player configuration allows you to set up and customize sound playback in your pinball machine. Sound players are used to play sounds in response to events.

## Configuration Options

### Basic Configuration
```vbscript
' Basic sound player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SoundPlayer
        With .EventName("event_name")
            .Sound = "sound_name"
            .Action = "play"
        End With
    End With
End With
```

### Advanced Configuration
```vbscript
' Advanced sound player configuration within a mode
With CreateGlfMode("mode_name", priority)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SoundPlayer
        ' Configure debug
        .Debug = True
        
        ' Configure sound events
        With .EventName("target_hit")
            .Sound = "target_sound"
            .Action = "play"
            .Volume = 80
            .Loops = 1
            .Key = "target_key"
        End With
        
        With .EventName("multiball_start")
            .Sound = "multiball_sound"
            .Action = "play"
            .Volume = 100
            .Loops = 0
            .Key = "multiball_key"
        End With
        
        With .EventName("multiball_end")
            .Sound = "multiball_end_sound"
            .Action = "stop"
            .Key = "multiball_key"
        End With
    End With
End With
```

## Property Descriptions

### Sound Settings
- `Sound`: The name of the sound to play
- `Action`: Action to perform on the sound ("play" or "stop") (Default: "play")
- `Volume`: Volume level for the sound (0-100) (Default: Empty)
- `Loops`: Number of times to loop the sound (0 for infinite) (Default: Empty)
- `Key`: Key to identify the sound for stopping (Default: Empty)
- `Debug`: Boolean to enable debug logging for this sound player (Default: False)

## Example Configurations

### Basic Sound Player Example
```vbscript
' Basic sound player configuration within a mode
With CreateGlfMode("base", 10)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SoundPlayer
        With .EventName("target_hit")
            .Sound = "target_sound"
            .Action = "play"
        End With
    End With
End With
```

### Advanced Sound Player Example
```vbscript
' Advanced sound player configuration within a mode
With CreateGlfMode("multiball", 20)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    
    With .SoundPlayer
        ' Configure debug
        .Debug = True
        
        ' Configure sound events
        With .EventName("jackpot_hit")
            .Sound = "jackpot_sound"
            .Action = "play"
            .Volume = 100
            .Loops = 1
            .Key = "jackpot_key"
        End With
        
        With .EventName("multiball_start")
            .Sound = "multiball_sound"
            .Action = "play"
            .Volume = 100
            .Loops = 0
            .Key = "multiball_key"
        End With
        
        With .EventName("multiball_end")
            .Sound = "multiball_end_sound"
            .Action = "stop"
            .Key = "multiball_key"
        End With
    End With
End With
```

## Sound Player System

The sound player system manages sound playback with the following features:

- Sounds can be played in response to events
- Sounds can be stopped in response to events
- Sounds can be configured with volume and loop settings
- Sounds can be identified by keys for stopping
- Sound players are managed within the context of a mode

## Events

The sound player system doesn't generate its own events, but it responds to events defined in your configuration.

## Default Behavior

By default, sound players are configured with:
- No events defined
- Debug logging disabled

## Notes

- Sound players are managed within the context of a mode
- The sound player system automatically handles sound playback
- Debug logging can be enabled to track sound player operations
- Sound players can be configured to work with specific game modes or features
- The sound player system prevents sound conflicts
- Sounds must be defined in the sound system before they can be played
- The `Action` property determines whether a sound is played or stopped
- The `Key` property is used to identify sounds for stopping
- The `Volume` property can be used to adjust the volume of sounds
- The `Loops` property can be used to make sounds loop
- When a mode is deactivated, all sounds played by that mode are stopped 