# Events and event ordering

GLF rules communicate by posting named events. Physical switch events typically use `<switch_name>_active` and `<switch_name>_inactive`; devices and modes emit their own names. An event name is not a function call. A mode's players respond only while their listeners are active.

## Map events inside a mode

```vbscript
With .EventPlayer()
    .Add "s_ramp_active", Array("ramp_award", "ramp_feedback")
    .Add "s_ramp_active{current_player.scoring_multiplier == 2}", Array("double_ramp_feedback")
End With
```

Replace the switch name and initialize `scoring_multiplier` before use. The conditional entry is an additional listener, not an `else` branch: both mappings run when the condition is true. See [Event Player](event-player.md) and [dynamic values](dynamic_values.md).

## Priorities and conditions

Higher numeric listener priorities run first. Mode players normally add the mode priority to an event-specific suffix: `ramp_award.2` listens for `ramp_award` with a priority offset of 2. A suffix is not part of the dispatched event name. A condition follows in braces, for example `ramp_award.2{current_player.score > 100000}`.

Use explicit priorities when ordering matters; do not rely on the order of separate player sections. Prefer a separate follow-up event when one operation must happen after another. Avoid circular mappings such as A posting B and B posting A.

## Three dispatch paths

| API | Behavior | Typical use |
| --- | --- | --- |
| `DispatchPinEvent name, data` | Adds an ordinary event to the dispatch queue processed by the game timer | Switch notifications, scoring, feedback |
| `DispatchRelayPinEvent(name, value)` | Calls listeners synchronously, passing each return value to the next | Claiming drained or entering balls |
| `DispatchQueuePinEvent name, data` | Queues an event whose handlers can pause the remaining handlers with `wait_for` | Bonus before ball end, a presentation before continuing |

Do not assume an ordinary event's effects have happened immediately after `DispatchPinEvent` returns. Relay callbacks must return the value for the next listener. Queued callbacks must return the event dictionary, using `Set` for the object result.

## Custom listeners

Register a callback during configuration:

```vbscript
AddPinEventListener "feature_award", "my_feature_award", "OnFeatureAward", 100, Null
```

The callback receives `args(0)` (registration arguments), `args(1)` (event data), and, for pin-event dispatch, `args(2)` (event name).

```vbscript
Function OnFeatureAward(args)
    Dim data
    Set data = args(1)
    SetPlayerState "score", GetPlayerState("score") + data("amount")
    Set OnFeatureAward = data
End Function
```

Post it only with an active player and the expected dictionary:

```vbscript
Dim awardData
Set awardData = GlfKwargs()
awardData.Add "amount", 10000
DispatchPinEvent "feature_award", awardData
```

The same event can be consumed by a variable player with `.Int = "kwargs.amount"`. Choose one scoring handler to avoid awarding twice. `kwargs` is meaningful only when the event supplies that data.

Remove a manually registered listener with its exact event and key:

```vbscript
RemovePinEventListener "feature_award", "my_feature_award"
```

## Relay and queued handlers

A minimal scalar relay callback preserves the incoming value:

```vbscript
Function PreserveBallCount(args)
    PreserveBallCount = args(1)
End Function
```

Ball-management handlers modify that count when they claim a ball. Returning nothing can break downstream accounting; use GLF's ball saves, holds and locks rather than adding another drain handler for a normal rule.

For queued work, prefer [Queue Relay Player](queue-relay-player.md) or a mode with `UseWaitQueue`, as shown in [game lifecycle](guides/game-lifecycle.md). Make every wait reachable through a completion or cancellation path.

## Debug an event

Check the emitted name, active mode, condition inputs, listener priority, and whether an earlier handler is waiting. Enable logging as described in [Debugging](guides/debugging.md). Core implementation: `scripts/src/pin-event-handlers.vbs`, the dispatch loop in `scripts/src/_.main.vbs`, and `scripts/src/events.vbs` for event parsing objects.
