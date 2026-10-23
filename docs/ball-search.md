# Ball search

Ball search tries device actions after a period without playfield activity. It relies on correct ball tracking, a configured plunger and working device callbacks. Set those up before enabling it.

## Enable during configuration

```vbscript
With EnableGlfBallSearch()
    .Timeout = 15000
    .SearchInterval = 150
End With
```

Call this once before `Glf_Init`. It creates the search controller and a `glf_ball_search` mode started by `reset_complete`. That mode detects a flipper cradle after three seconds, pauses search while the cradle is active and resets it on release.

| Setting | Default | Current behavior |
| --- | --- | --- |
| `Timeout` | 15000 ms | Delay used when resetting or retrying search |
| `SearchInterval` | 150 ms | Delay between device attempts |
| `BallSearchWaitAfterIteration` | 10000 ms | Exposed setting, but not used by the current search loop |
| `Debug` | False | Debug flag; verify logging against the framework revision |

The constructor initially schedules a 15-second timeout before these setters run. A custom `Timeout` applies to subsequent resets/retries, not retroactively to that first scheduled delay.

## When search runs

The controller checks that a game is running, balls are in play, at least one is not accounted for by ball holds, and the plunger is empty. A failed check schedules another attempt. Normal switch activity participates in resetting search through GLF's event handling.

Search visits autofire devices, ball devices, drop targets and diverters, calling each device's `BallSearch(phase)` method. It progresses through phases up to 3. The actual action is defined by each device. Flippers, magnets and standup targets are not additional groups in this loop.

The listed device types expose `ExcludeFromBallSearch`. Set it during device configuration for a mechanism that should not fire during recovery. Diverters also expose `BallSearchHoldTime` for their search action.

## Events and recovery

Listen for `ball_search_started` and `ball_search_stopped` if a feature needs to release a hold or end a presentation. Dark Chaos's victory-lap mode responds to search by releasing the moon ball and ending the feature. Keep recovery actions compatible with the device's current state.

Test a stationary playfield ball, an occupied plunger, a flipper cradle, a deliberately held scoop ball and every configured search device. Check the table's device arrays too: the current loop only advances to the next group after a device callback; an empty group can prevent traversal from reaching later groups. Do not assume a table with no autofire devices will search every other device.

Source: `scripts/src/ball-search.vbs` and each participating device's `BallSearch` implementation. The example table contains a commented-out setup block in `_configuration.vbs`; it is a reference, not evidence that search is enabled in that table.
