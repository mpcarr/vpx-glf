# Diverter

A diverter in vpx is typically implemented using two walls or pimitives switching the colliable property depending on if the
diverter is active or not. This virutal device manages if the diverter should be active or not and fires the callback to move the
diverter.

## Examples 

*Example 1: A ramp diverter that enables when multiball locks are open and diverts the ball after the ramp entrace switch is hit*

```
Dim diverter_ramp
Set diverter_ramp = (new Diverter)("ramp")

With diverter_ramp
    .EnableEvents = Array("multiball_locks_open")
    .DisableEvents = Array("multiball_locks_closed")
    .ActivationTime = 2
    .ActivateEvents = Array("sw23_active")
    .ActionCallback = "RampDiverter"
End With
```

In the above example we have a diverter on a ramp, it is enabled when the **multibal_locks_open** event is emitted and disables when the **multiball_locks_closed** is emitted. Enabling the diverter does not mean it activates (moves), it means that it will start to listen for it's activate events. When the event **s23_active** (ramp entrance switch) is emitted, the diverter will call the action callback **RampDiverter**. After 2 seconds will will call the same callback again with a disabled param.

```
Sub RampDiverter(enabled)
    If enabled = 1 Then
        'Activate Diverter Logic
    Else
        'Deactivate Diverter Logic
    End If
End Sub
```

*Example 2: A diverter that opens during a mode and stays open until a disable event is emitted.*

```
Dim diverter_orbit : Set diverter_orbit = (new Diverter)("orbit")
With diverter_orbit
    .EnableEvents = Array("game_started")
    .DisableEvents = Array("game_ended)
    .ActivateEvents = Array("mode_hurry_started")
    .DisableEvents = Array("mode_hurry_ended")
    .ActionCallback = "OrbitDiverter"
End With
```

In the above example, the diverter is enabled at game start and disabled at game end, it listens for the **mode_hurry_started** event to activate and will stay activated unti the **mode_hurry_ended** event.

## Required Settings

### Name
```String```

The name for this diverter. Events emitted from the diverter will be in the format *name*_diverter

### Enable Events
```Array```

A list of events that will enable this diverter. When enabled, the diverter will listen for it's activation and deactivation events

### Disable Events
```Array```

A list of events that will disable the diverter. Didabled diverters won't react to any activation / deactivation events.

## Optional Settings

### Activate Events
```Array```

A list of events that will trigger the action callback with enabled = 1

### Deacctivate Events
```Array```

A list of events that will trigger the action callback with enabled = 0

### Action Callback
```String```

The sub or function that is called when the diverter is activated / deactivated.

### Activation Time
```Number```

The number of seconds after activation the diverter will stay open for

### Activation Switches
```Array```

TBC

### Debug
```Boolean```