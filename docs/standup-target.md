# Standup targets

Register a standup target so GLF can generate its hit/unhit wiring and emit switch events. This is a global device configured before `Glf_Init`, not a mode component.

## Basic target

The table must contain a target object named `s_ST1`:

```vbscript
With CreateGlfStanduptarget("target1")
    .Switch = "s_ST1"
End With
```

GLF generates the target's event subs. Do not also add it to `glf_switches` or duplicate its `_Hit`/`_UnHit` subs; the example table explicitly keeps standup targets out of those collections.

For a non-Roth target, generated hit/unhit handlers post `s_ST1_active` and `s_ST1_inactive` while the game is not tilted. Listen for the active event in a shot or event player; registration itself does not award points.

## Roth target integration

If the table already has working Roth standup physics and its associated `STHit` implementation:

```vbscript
With CreateGlfStanduptarget("target1")
    .Switch = "s_ST1"
    .UseRothStanduptarget = True
    .RothSTSwitchID = 1
End With
```

The ID must match the table's Roth configuration. GLF's generated hit handler calls `STHit` instead of directly posting the switch event. The physics integration must emit the GLF events at the appropriate target action; inspect the example table's `scripts/src/vpx/standup-targets.vbs` alongside `_configuration.vbs`.

| Property | Default | Purpose |
| --- | --- | --- |
| `Switch` | Empty | VPX target object name |
| `UseRothStanduptarget` | False | Route hits through Roth target physics |
| `RothSTSwitchID` | -1 | ID passed to `STHit` |
| `Debug` | False | Device logging flag |

Test one physical hit produces one rule hit, including after a tilt and on the next ball. Related: [shots](shot.md), [drop targets](drop-target.md). Source: `scripts/src/devices/standup-target.vbs` and handler generation in `scripts/src/_.main.vbs`.
