# Queue relay player

The queue relay player pauses the remaining handlers of a queued event, posts another event to begin work, and resumes when a completion event arrives. It is useful for finishing an award presentation before ball end.

## Delay ball end for a presentation

Inside a mode that is already active when `ball_ending` arrives and has priority above the core end handler's 10:

```vbscript
With .QueueRelayPlayer()
    With .EventName("ball_ending")
        .Post = "start_award_presentation"
        .WaitFor = "award_presentation_finished"
    End With
End With
```

`ball_ending` is dispatched as a queued event by GLF. The player posts `start_award_presentation` and returns a `wait_for` entry naming `award_presentation_finished`. When that completion event is dispatched, the remaining ball-ending handlers continue.

This excerpt requires your presentation/timer logic. Every success, skip and cancellation path must post `award_presentation_finished`. Keep the controlling mode alive until it can signal completion; removing its event listeners alone does not release the paused queue.

Dark Chaos's `scripts/src/game/modes/victory-lap.vbs` uses the same pattern to wait for `end_victory_lap` while music finishes. Its `final-wave-wizard.vbs` contains another example.

## Settings and limits

| Property | Purpose |
| --- | --- |
| `EventName` | Queued event to intercept; supports the event parser's condition and priority suffix |
| `Post` | Ordinary event posted to begin the asynchronous work |
| `WaitFor` | Event that releases the remaining queued handlers |
| `Debug` | Component debug flag |

Supply both `Post` and `WaitFor`; their defaults are empty. The posted event uses `Null` data, so this player does not automatically forward the original event arguments to the presentation.

Use this with `DispatchQueuePinEvent`, not as a way to make an ordinary notification block. A mode's `UseWaitQueue` offers a related pattern: starting that mode from a queued event waits for the mode to stop. See [game lifecycle](guides/game-lifecycle.md).

Source: `scripts/src/commands/queue-relay-player.vbs`, `pin-event-handlers.vbs` and the active dispatch loop in `_.main.vbs`. [Queue Event Player](queue-event-player.md) posts a queued event; queue relay player intercepts one already being processed.
