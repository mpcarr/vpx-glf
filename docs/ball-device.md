# Ball Device Configuration

Ball devices track and eject balls in scoops, physical locks, and staging areas. Configure occupancy switches and an eject callback for each device.

```vbscript
With CreateGlfBallDevice("scoop")
    .BallSwitches = Array("s_scoop1")
    .EjectCallback = "ScoopEjectCallback"
    .EjectTimeout = 2000
    .EjectRetryDelay = 500
End With
```

The callback receives the VPX ball object to kick. If `EjectEnableTime` is positive, it also receives `Null` when the enable period ends or the attempt completes or is cancelled. Ball search can also call it with `Null`; callbacks must handle that value before accessing the ball.

```vbscript
Sub ScoopEjectCallback(ball)
    If IsNull(ball) Then Exit Sub
    If s_scoop1.BallCntOver > 0 Then
        KickBall ball, 14.8, 70, 0, 0
    End If
End Sub
```

## Eject lifecycle and counting

`Eject` queues one ball, `EjectBalls n` queues a requested number, and `EjectAll` queues the physical balls not already reserved.

- Eject requests are queued and serialized. Each available physical ball is reserved once. Requests exceeding supply wait for a new arrival without polling an empty device indefinitely.
- An unexpected, unclaimed arrival requests release of one available ball, preferring the ball already at `EjectSwitchIndex`. For example, when one ball is locked at the exit and a second ball arrives behind it, the front ball ejects and one ball remains held. The actual selected ball is tracked by identity for departure, return/retry, and confirmation.
- The actual kick is the final handler of the asynchronous `_ball_eject_attempt` queue event. Higher-priority handlers can block it with GLF's existing `wait_for` mechanism. Reserve priority `-2147483647` for the internal final handler.
- A ball that does not leave retries the same request. A returning ball is recognised by VPX `ball.ID`, does not generate a second capture/claim event, and does not consume another incoming expectation.
- Timer callbacks carry attempt tokens. Old callbacks and confirmation listeners cannot complete a later attempt. Success completes one request once.
- With no `EjectTargets`, success requires a settled departure and an interval without a return. With `EjectTargets`, a matching physical ball must hit a configured confirmation switch. A missing confirm is reported as failure/missing, never success, and does not kick a different ball as a replacement.
- Counting reconciles unique physical balls across all configured occupancy switches. Duplicate hits/unhits, overlapping switches, and movements shorter than the settling window do not add/subtract extra balls. Both arrival and departure use settling delays. For a single-switch mechanical plunger, a matching Hit/UnHit before entrance settling proves a passage: it is received once at UnHit, consuming its incoming reservation, and remains counted until exit settling completes.
- Occupancy and confirmation listeners receive every physical switch event through the dispatcher's queued raw channel. The normal 50 ms hit filter still applies to ordinary event listeners. This prevents a rapid return hit from being discarded while its preceding exit was delivered. Use the updated event dispatcher together with the device class, or the rebuilt framework bundle.
- `EjectAll` reserves only remaining unreserved balls, including when an earlier ball has left and is awaiting confirmation. An empty call and `EjectBalls 0` are no-ops.
- `IncomingBalls` is consumed on arrival. The built-in ball-save, multiball, and multiball-lock readiness checks prevent ball saves/multiballs from serving into a plunger whose previous launch is still awaiting confirmation.
- A returned mechanical launch waits for another player launch. An explicit `Eject` can upgrade that pending request to a powered launch. Mechanical confirmation listeners start at UnHit, so a confirmation before exit settling is retained; a return invalidates it. A launch-button event while empty does not arm a future ball.
- An idle lost ball is tracked by identity and eventually emits missing-ball events; a different arriving ball is not mistaken for its return.
- `StopDevice` invalidates attempts and disables an enabled eject callback. Switch observations continue while stopped so `Reset` can recount. Reset clears queued eject requests; it does not create new requests or clear externally owned queue blockers.

## Timing and eject-position configuration

All durations below are integer milliseconds. Count delays and `EjectEnableTime` may be zero; timeout and retry durations must be greater than zero.

| Property | Default | Purpose |
|---|---:|---|
| `EntranceCountDelay` | 500 | Stable time after a switch activation. Applies to the default plunger too. |
| `ExitCountDelay` | 500 | Stable time after a switch deactivation. |
| `EjectTimeout` | 1000 | Departure timeout; also the subsequent confirmation/return-check interval. Must be greater than zero. |
| `EjectRetryDelay` | 500 | Delay between a failed attempt and its retry. |
| `MaxEjectAttempts` | 0 | Maximum attempts including the first; zero means unlimited. |
| `BallMissingTimeout` | 20000 | Additional grace after an explicit confirmation timeout before declaring the ball missing. |
| `IdleMissingBallTimeout` | 3000 | Time after a settled idle departure before declaring the ball missing. |
| `EjectSwitchIndex` | 0 | Zero-based entry in `BallSwitches` identifying the physical eject position. |
| `EjectEnableTime` | 0 | When positive, callback receives `Null` to disable after this duration, or earlier when its attempt is cancelled/completed. |

Example:

```vbscript
With CreateGlfBallDevice("moon_lock")
    .BallSwitches = Array("s_moon_lock")
    .EjectCallback = "MoonLockEject"
    .EjectTimeout = 1500
    .EntranceCountDelay = 500
    .ExitCountDelay = 500
    .EjectRetryDelay = 500
    .MaxEjectAttempts = 0
End With
```

For a multi-switch device, `BallSwitches` describes occupancy sensors, not eject-confirmation switches. The same physical ball can overlap sensors. `EjectSwitchIndex` identifies where the callback can actually eject a ball; the device waits for a reserved ball to reach that position.

`EjectTargets` retains GLF's existing switch-name meaning, for example `Array("s_lock_exit")` subscribes to `s_lock_exit_active`. It is **not** an MPF device-routing graph. Confirmation events must carry the VPX ball object, as generated GLF switch events already do.

Mechanical ejection is supported for a single occupancy switch, matching the single-ball plunger use case. Multi-switch powered devices use the same counting and queued-eject implementation without `MechanicalEject`.

## Properties and integration

| Property | Meaning |
|---|---|
| `Balls` | Number of unique physical balls, reconciled after settling. A fast mechanical passage is included from UnHit until its exit settles. |
| `HasBall` | Whether the device count is greater than zero, including balls away from the eject position and balls whose departure has not settled. |
| `IncomingBalls` | Outstanding arrivals, consumed when received. |
| `IsReadyToReceive` | Stable free switch capacity after incoming reservations, with no conflicting active eject or unconfirmed departure. Capacity is based on the number of configured occupancy switches. Use this for serving the plunger. |
| `PendingEjects` | Queue length, including the active request. |
| `State` | Current lifecycle state. |
| `EjectAttempts` | Attempts made for the current/most recent request. |

Any custom code that previously used `HasBall = False And IncomingBalls = 0` to serve the plunger should use `IsReadyToReceive = True`. `HasBall` is now an occupancy query, not an eject-readiness query.

Existing entry relay arguments remain numeric, and existing entered/exiting/ejecting/success events retain their `Null` payloads. Expected incoming balls still bypass GLF's claim relay for compatibility with the supplied implementation.

New events:

- `balldevice_<name>_ball_count_changed`: kwargs `balls`.
- `balldevice_<name>_ball_eject_failed`: kwargs `balls`, `retry`, `num_attempts`, `reason`.
- `balldevice_<name>_broken`: emitted when the configured attempt limit is reached. Call `Reset` and issue new requests after correcting the problem.
- `balldevice_<name>_ball_missing` and `balldevice_ball_missing`: kwargs `balls`, `name`, `ball_id`.

The queue attempt event now supplies `source`, `balls`, `num_attempts` (completed attempts before the next kick), and internal `glf_eject_token`. Queue handlers must return the kwargs object, as required by GLF.

## Balls already present at initialization

VPX's generic switch integration does not provide a reliable switch-to-ball-object query. For balls placed on switches before normal Hit events are enabled, seed the known ball objects once during initialization:

```vbscript
glf_ball_devices("moon_lock").SeedBall existingBallObject, 0
```

Seed each occupied sensor with its physical ball object and zero-based switch index. Seeding the same ball on overlapping sensors counts it once and does not emit capture events. Normal Hit/UnHit events must subsequently supply the actual ball object. This implementation does not infer initial occupancy from table geometry.

## Other configuration properties

| Property | Default | Meaning |
|---|---|---|
| `BallSwitches` | Empty array | Occupancy switch names. Hit/UnHit events must carry the VPX ball object. |
| `EjectCallback` | `Null` | Name of the callback that kicks the ball. |
| `EjectTargets` | Empty array | Confirmation switch names; see the distinction from occupancy sensors above. |
| `EjectEvents` | Empty array | Events that queue one eject. |
| `EjectAllEvents` | Empty array | Events that queue all currently unreserved balls. |
| `MechanicalEject` | `False` | Track a player-launched ball; requires a single occupancy switch. |
| `DefaultDevice` | `False` | Register this device as `glf_plunger`. |
| `PlayerControlledEjectEvent` | Empty string | Event that launches an available ball. An event while empty does not arm a future launch. |
| `AutoFireOnUnexpectedBall` | `True` | Automatically eject an unexpected, unclaimed arrival. |
| `ExcludeFromBallSearch` | `False` | Exclude the device from ball search. |
| `Debug` | `False` | Enable device logging. |

Add incoming reservations with `AddIncomingBalls`. `MaxEjectAttempts` includes the first attempt. After correcting a problem that caused `eject_broken`, call `Reset` and issue new requests.

## Existing events

Replace `<name>` with the name passed to `CreateGlfBallDevice`.

| Event | Payload / purpose |
|---|---|
| `balldevice_<name>_ball_enter` | Numeric claim relay for unexpected arrivals. |
| `balldevice_<name>_ball_entered` | `Null`; settled arrival. |
| `balldevice_<name>_ball_exiting` | `Null`; settled departure. |
| `balldevice_<name>_ball_eject_attempt` | Queue event described above. |
| `balldevice_<name>_ejecting_ball` | `Null`; an eject is being tracked. |
| `balldevice_<name>_ball_eject_success` | `Null`; one request completes successfully. |

## Regression tests

From the `scripts` directory on Windows:

```console
npm run test:ball-device
```

Alternatively, run `cscript //nologo tests\ball-device-tests.vbs`. The runner loads the current source device, delay/event helpers, frame dispatcher, and multiball-lock code. It mocks VPX ball objects, the clock, mode configuration, player state, and the trough boundary. It defines 68 scenarios, including fast manual launches and a first lock followed by a replacement serve and auto-launch, and exits nonzero for assertion failures or script errors. The hit helpers also update the shared debounce clock, as the generated VPX switch handlers do. The Windows GitHub Actions workflow runs this suite for relevant pull-request changes. Follow it with VPX checks of the actual lock return path, plunger, multi-ball movement, and confirmation-switch placement.

## Scope

This implements local eject lifecycle and counting. It does not add MPF's device-routing graph, per-source incoming-transfer objects, jam-switch/reorder logic, pulse-strength escalation, or global playfield missing-ball accounting. Missing-ball events provide a hook for the latter; they do not themselves change GLF's balls-in-play count.
