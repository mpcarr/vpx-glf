# Ball holds

A ball hold claims balls entering named ball devices and releases them on rule events. Use it to keep a ball in a scoop during an award or selection. Use a [multiball lock](multiball-locks.md) when you need lock qualification and replacement-ball behavior.

## Configure a scoop hold

First create a ball device named `scoop` with a working eject callback. Inside a mode that remains active for the hold's lifetime:

```vbscript
With .BallHolds("scoop_award")
    .BallsToHold = 1
    .HoldDevices = Array("scoop")
    .EnableEvents = Array("enable_scoop_award")
    .DisableEvents = Array("disable_scoop_award")
    .ReleaseAllEvents = Array("scoop_award_finished")
End With
```

Enable the hold before the ball enters. The hold listens to `balldevice_scoop_ball_enter`, claims available balls, and prevents the device's normal unclaimed-ball path from releasing them. Post `scoop_award_finished` after the award presentation. Mode deactivation disables the hold and releases held balls.

Dark Chaos uses this pattern in `scripts/src/game/modes/base.vbs` with `scoop_hold` and explicit enable, disable and release events.

## Settings

| Property | Meaning / default |
| --- | --- |
| `BallsToHold` | Capacity across the configured hold devices; default 0, so set it explicitly |
| `HoldDevices` | Array of registered ball-device names; default empty |
| `EnableEvents` | Enable the hold on these events; without any, mode activation enables it |
| `DisableEvents` | Disable and release all held balls |
| `ReleaseAllEvents` | Release everything; `tilt` is registered by default |
| `ReleaseOneEvents` | Release one held ball |
| `ReleaseOneIfFullEvents` | Release one only when the hold is full |
| `Debug` | Enable component logging |

Release-event setters add entries; the built-in tilt release remains when you add your own events. Avoid registering the same release key twice.

## Events and runtime values

`ball_hold_scoop_award_held_ball` carries `balls_held` for that claim and `total_balls_held` for the hold. `ball_hold_scoop_award_full` carries `balls`. Read the count with `device.ball_holds.scoop_award.balls_held` in supported runtime expressions.

Test release after success, cancellation, tilt and mode stop. Do not leave a presentation waiting indefinitely while it owns the ball. For multiple claims into the same device, test against the current implementation: its internal hold queue is keyed by device name, so repeated entries are not a general per-ball queue.

Source: `scripts/src/commands/ball-hold.vbs`. Related: [ball devices](ball-device.md), [ball search](ball-search.md), [events](events.md).
