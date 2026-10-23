# Game and ball lifecycle

Use lifecycle events to decide how long a mode lives and when its state resets. The implementation is in `scripts/src/game.vbs`; event constants are defined in `scripts/src/_.main.vbs`.

## From start button to first ball

| Event | What it means for your rules |
| --- | --- |
| `request_to_start_game` | A start request is checked, including whether the trough contains the expected balls. Do not use it to award points. |
| `player_added` | A player is being added; `kwargs.num` is one-based. The event is posted as the player is created. |
| `game_start` | The game flag and first player have been set. Start game-long modes here. |
| `game_started` | Posted after a short startup delay. Player-state refresh and ball-release handlers run from here. |
| `balldevice_trough_ball_eject_attempt` | A queued ball-release event. A release carrying `new_ball` starts a new ball. |
| `ball_started` | The new ball is counted in play and is being served. Start per-ball rules and reset per-ball variables here. |

`ball_started` is a serve-time event, not proof that the player has plunged or hit the playfield. Use a suitable playfield switch or your ball-save timer-start event for features that should begin only after launch.

## From drain to next player

An ordinary drain passes through the `ball_drain` relay. Saves and multiball handlers can consume the drain before the core game handler ends the ball. Do not end a ball from a raw drain switch hit.

When the last unclaimed ball drains, the core posts `ball_will_end`, then queues `ball_ending`. The low-priority end-of-ball handler posts `ball_ended` after higher-priority queued work is released. This is where bonus fits.

After `ball_ended`, an extra ball keeps the same player and ball number, consuming one `extra_balls` award. Otherwise the outgoing player's ball number increases and the current player advances. If the next player still has balls remaining, `next_player` leads to another ball release. Even an extra-ball serve uses this continuation path; the event name alone does not prove the player changed.

At the end of the game, the core posts `game_will_end` and queues `game_ending`. It copies final scores to machine variables, clears player state, then posts `game_ended`. Read final scores from the machine variables after this point, not `current_player`.

## Choose a mode lifetime

```vbscript
With CreateGlfMode("scoring", 2000)
    .StartEvents = Array("game_start")
    .StopEvents = Array("game_ended")
End With

With CreateGlfMode("base", 100)
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_will_end", "tilt")
End With
```

These are setup excerpts; add components inside each mode. A feature that belongs to base can stop on `mode_base_stopping`. The example table's `modes/score.vbs` uses a game-long mode, while `modes/multiball.vbs` starts each ball and stops with base.

## Wait for bonus

The example table's `scripts/src/game/modes/bonus.vbs` uses:

```vbscript
With CreateGlfMode("bonus", 150)
    .StartEvents = Array("ball_ending{game.tilted == False}")
    .StopEvents = Array("bonus_finished")
    .UseWaitQueue = True
End With
```

Add the actual bonus calculation and presentation before using this excerpt. `UseWaitQueue` asks the queued start event to wait for `mode_bonus_stopped`. Every completion and skip path must eventually post `bonus_finished`; otherwise the next ball cannot proceed. Use a priority above the core `ball_ending` handler's 10 so the wait is installed before that handler runs.

Tilt has its own drain/settle handling in `scripts/src/commands/tilt.vbs`. Stop scoring features on tilt and use the condition above to skip bonus. Do not manually advance players to bypass a stalled queue; find the missing completion event.

## Game settings and verification

Use the table's `Balls Per Game` option to choose `3 Balls` or `5 Balls`. `Glf_Init` creates the game settings object with three balls, then calls `Glf_Options`, which applies that option. A separate settings call before initialization is overwritten; use the existing option for the supported choices. This is the number of turns per player, independent of `tnob`, the physical ball count.

Test a normal drain, saved drain, multiball drain, tilt, extra ball, two-player change and final game end. Confirm that progress resets at the intended boundary. See [player state](../player-state.md), [queue relay player](../queue-relay-player.md) and [debugging](debugging.md).
