# Dynamic runtime values

Some GLF properties accept expressions evaluated when the value is used. For these properties, pass the expression as a **VBScript string**. For example, `.Int = "100000 * current_player.scoring_multiplier"` defers evaluation until the scoring event; removing the quotes attempts to evaluate `current_player` as a VBScript object during configuration.

## Score with a multiplier

Initialize your custom variable during table configuration:

```vbscript
Glf_SetInitialPlayerVar "scoring_multiplier", 1
```

Inside an active mode, configure the scoring event:

```vbscript
With .VariablePlayer()
    With .EventName("jackpot_hit")
        With .Variable("score")
            .Action = "add"
            .Int = "100000 * current_player.scoring_multiplier"
        End With
    End With
End With
```

A multiplier of 2 awards 200,000 points when `jackpot_hit` occurs. Dark Chaos uses this pattern in `scripts/src/game/modes/score.vbs`; the example table constructs similar strings using `x & " * current_player.scoring_multiplier"`.

The exported YAML uses the expression directly:

```yaml
variable_player:
  jackpot_hit:
    score:
      action: add
      int: 100000 * current_player.scoring_multiplier
```

See [configuration](configuration.md) for how exported YAML relates to the VPX source.

## Runtime references

| Reference | Reads | Requirement |
| --- | --- | --- |
| `current_player.score` | Current player's score | An active player |
| `current_player.scoring_multiplier` | A custom player variable | Initialize it during configuration |
| `machine.my_variable` | A machine variable's value | Create the machine variable first |
| `device.timers.double_scoring.ticks` | A timer's current count | A configured timer named `double_scoring` |
| `device.timers.double_scoring.ticks_remaining` | A timer's remaining count | A configured timer named `double_scoring` |
| `kwargs.amount` | Data carried by the current event | The event must supply `amount` |

Device attributes are limited to those exposed by that device's `GetValue` implementation; arbitrary VBScript properties are not automatically available as placeholders. The parser also handles `players[...]`, `modes` and `game` references; check their implementations before relying on an attribute or index convention.

## Conditions on events

A condition inside braces filters an event subscription. This pattern from Dark Chaos changes a shot state near the end of a timer:

```vbscript
With .Shots("double_scoring")
    With .ControlEvents()
        .Events = Array("timer_double_scoring_tick{device.timers.double_scoring.ticks == 55}")
        .State = 2
    End With
End With
```

This excerpt belongs inside the mode containing the timer and requires a shot profile with state index 2. The condition is evaluated when the tick event arrives; it does not continuously watch the timer.

## Check whether a property supports expressions

In `scripts/src/commands/` or `scripts/src/devices/`, inspect the property's setter. A setter that calls `CreateGlfInput` supports the runtime input parser. A setter that stores a plain value may behave differently. Variable player `Int`, `Float` and `String` use `CreateGlfInput`; their values are evaluated in `Play`.

The implementation is in `scripts/src/_.main.vbs`: `Glf_ParseInput`, `Glf_ParseEventInput`, the `Glf_Replace...Attributes` helpers and `GlfInput.Value`. Event conditions and value expressions use different parsing paths, so do not assume every expression syntax works in both.
