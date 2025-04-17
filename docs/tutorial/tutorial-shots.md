# Tutorial - Shots

Let's add shots to our pinball table. Shots are the primary interactive elements that players can hit to score points and trigger game events. In this tutorial, we'll learn how to configure shots, shot profiles, and shot groups.

## What are Shots?

In pinball, a "shot" is a switch (or combination of switches) that the player shoots for. Examples include:

- A standup target, drop target, or rollover lane
- A ramp, loop, or orbit
- A toy, subway, or VUK (Vertical Up Kicker)

Most shots have lights or LEDs associated with them which are on, off, flashing, and/or certain colors to reflect what "state" the shot is in.

## Shot Configuration

### Basic Shot Setup

First, let's create a basic shot configuration:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        ' Configure a basic shot
        With .Shot("left_ramp")
            .Switch = "sw_left_ramp"
            .Profile = "ramp_profile"
        End With
    End With
End Sub
```

This configuration:
1. Creates a shot called "left_ramp" in the "base" mode
2. Configures it to be triggered by the "sw_left_ramp" switch
3. Uses the "ramp_profile" shot profile to define its behavior

### Shot Profiles

Before we can use a shot profile, we need to define it. Shot profiles define how shots behave, including their states and visual effects:

```
Sub CreateShotProfiles()
    ' Create a basic ramp profile
    With CreateGlfShotProfile("ramp_profile")
        ' Configure basic settings
        .AdvanceOnHit = True
        .Block = False
        .ProfileLoop = False
        
        ' Configure states
        With .States("unlit")
            .Show = "off"
            .Key = "key_ramp_unlit"
        End With
        
        With .States("flashing")
            .Show = "flash"
            .Key = "key_ramp_flashing"
            With .Tokens
                .Add "color", "yellow"
            End With
        End With
        
        With .States("lit")
            .Show = "on"
            .Key = "key_ramp_lit"
            With .Tokens
                .Add "color", "red"
            End With
        End With
    End With
End Sub
```

This profile defines three states for the shot:
1. "unlit" - The shot is off
2. "flashing" - The shot is flashing yellow
3. "lit" - The shot is on with a red color (final state)

### Advanced Shot Configuration

Now let's create a more advanced shot configuration:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        ' Configure an advanced shot
        With .Shot("jackpot")
            ' Configure switches
            .Switches = Array("sw_left_ramp", "sw_right_ramp", "sw_center_ramp")
            
            ' Configure profile
            .Profile = "jackpot_profile"
            
            ' Configure enable events
            .EnableEvents = Array("multiball_start")
            
            ' Configure disable events
            .DisableEvents = Array("multiball_end", "ball_lost")
            
            ' Configure hit events
            .HitEvents = Array("add_score_1000")
            
            ' Configure show tokens
            With .Tokens
                .Add "color", "red"
            End With
            
            ' Configure start enabled
            .StartEnabled = True
            
            ' Configure persist
            .Persist = True
        End With
    End With
End Sub
```

This advanced configuration:
1. Creates a shot called "jackpot" that can be triggered by any of three switches
2. Uses the "jackpot_profile" shot profile
3. Enables the shot when "multiball_start" event occurs
4. Disables the shot when "multiball_end" or "ball_lost" events occur
5. Adds 1000 points when the shot is hit
6. Passes custom tokens to the show
7. Starts enabled
8. Persists its state between balls

### Shot Groups

Shot groups allow you to manage multiple shots together. For example, you might want to track when all shots in a group reach a certain state:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        ' Configure shot group
        With .ShotGroup("ramps")
            .Shots = Array("left_ramp", "right_ramp", "center_ramp")
            .CompleteEvents = Array("all_ramps_complete")
        End With
    End With
End Sub
```

This configuration:
1. Creates a shot group called "ramps" that includes three shots
2. Triggers the "all_ramps_complete" event when all shots in the group reach their final state

## Shot States and Events

Shots can be in different states, and the system generates events when shots change state:

- `shot_name_hit`: Fired when the shot is hit
- `shot_name_profile_hit`: Fired when the shot is hit with the profile name
- `shot_name_profile_state_hit`: Fired when the shot is hit with the profile name and state
- `shot_name_state_hit`: Fired when the shot is hit with the state name

You can use these events to trigger scores, achievements, shows, etc.

## Mode-Specific Shot Behavior

Shots can behave differently in different modes. For example, a ramp shot might score 1,000 points in the base mode, but score a jackpot in the multiball mode:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        ' Configure a shot in the base mode
        With .Shot("left_ramp")
            .Switch = "sw_left_ramp"
            .Profile = "ramp_profile"
            .HitEvents = Array("add_score_1000")
        End With
    End With
End Sub

Sub CreateMultiballMode()
    With CreateGlfMode("multiball", 2000)
        ' Configure the same shot in the multiball mode
        With .Shot("left_ramp_multiball")
            .Switch = "sw_left_ramp"
            .Profile = "jackpot_profile"
            .HitEvents = Array("add_score_10000", "jackpot_collected")
        End With
    End With
End Sub
```

## Sequence Shots

You can also configure shots that are only considered to be hit based on a series of switches that must be hit in the right order within a certain time frame:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        ' Configure a sequence shot
        With .SequenceShot("orbit_left")
            .Switches = Array("sw_orbit_left", "sw_orbit_center", "sw_orbit_right")
            .TimeWindow = 3000  ' 3 seconds
            .HitEvents = Array("add_score_5000")
        End With
        
        With .SequenceShot("orbit_right")
            .Switches = Array("sw_orbit_right", "sw_orbit_center", "sw_orbit_left")
            .TimeWindow = 3000  ' 3 seconds
            .HitEvents = Array("add_score_5000")
        End With
    End With
End Sub
```

This configuration:
1. Creates two sequence shots that use the same switches but in different orders
2. Sets a 3-second time window for the sequence to be completed
3. Adds 5,000 points when the sequence is completed

## Enabling the Modes

Finally, we need to enable these modes in our game by calling the appropriate subs in our `ConfigureGlfDevices` Sub:

```
Sub ConfigureGlfDevices
    CreateShotProfiles()  ' Create shot profiles
    CreateBaseMode()      ' Create base mode with shots
    CreateMultiballMode() ' Create multiball mode with shots
    
    'Other device config....
End Sub
```

## Conclusion

With these configurations, you now have a comprehensive shot system for your pinball game. You can:

1. Create basic and advanced shots
2. Define shot profiles with different states and visual effects
3. Group shots together for collective behavior
4. Configure mode-specific shot behavior
5. Create sequence shots for complex interactions
