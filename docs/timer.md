# Timer

Timers manage event durations, triggering multiball modes, bonus rounds, and target resets.

## Example

```
Dim timer_beasts_panther
Set timer_beasts_panther = (new ModeTimer)("beasts_panther", mode_beasts)

With timer_beasts_panther
    .StartEvents = Array("sw01_active")
    .StopEvents = Array("sw01_inactive")
    .Direction = "down"
    .StartValue = 10
    .EndValue = 0
End With
```

### Events

```timer_*name*_started```

```timer_*name*_stopped```

```timer_*name*_complete```

```timer_*name*_tick```
