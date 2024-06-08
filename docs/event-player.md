# Event Player

The event player listens for one event and emits one or more other events.

## Example

```
Dim event_player_super_pops
Set event_player_super_pops = (New EventPlayer)("super_pops", mode_super_pops)

event_player_super_pops.Add "mode_super_pops_started", _
Array(_
"play_super_pops_start_sound", _
"play_pops_light_show", _
"activate_top_gate _
")

event_player_super_pops.Add "mode_super_pops_ended",  _
Array(_
"play_super_pops_end_sound", _
"play_pops_ended_light_show", _
"deactivate_top_gate _
")
```

In the above example we created a event player called **event_player_super_pops** that belongs to the mode called **mode_super_pops**. The event player listens for the *mode_super_pops_started* event and emits three other events to play a sound, a light show and activate a diverter gate. It also listens for the *mode_super_pops_ended* event and then emits a end sound, end light show and deactivated the diverter.

## Required Setings

### Name
```String```

The name of this device. Events emitted from the device will be in the format **name_event_player**

### Mode
```String```

This is the mode the event player belongs to. 