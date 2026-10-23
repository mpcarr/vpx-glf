# Player variables and persistence

Choose a variable's lifetime before wiring it to events. Player variables belong to one player in the current game. Machine variables belong to the table and can optionally be saved between sessions.

| State | Store it in | Reset it when |
| --- | --- | --- |
| Score or qualification carried across balls | Player variable | A new player is created for a new game |
| Hits for this ball or a playfield multiplier | Player variable | Explicitly on `ball_started` |
| Physical occupancy shared by all players | Machine variable, usually `Persist = False` | Reconcile with the actual device state |
| A saved table statistic | Machine variable with `Persist = True` | Only when your rules require it |

## Initialize custom player variables

Before `Glf_Init`, register defaults once:

```vbscript
Glf_SetInitialPlayerVar "ramp_hits", 0
Glf_SetInitialPlayerVar "scoring_multiplier", 1
```

Each new player receives these defaults. GLF already initializes `score`, `initials`, `ball` and `extra_balls`; do not register those again as custom initial variables. Defaults do not automatically reset on the next ball.

Inside a mode active for the whole game, reset per-ball state and award points:

```vbscript
With .VariablePlayer()
    With .EventName("ball_started")
        With .Variable("ramp_hits")
            .Action = "set"
            .Int = 0
        End With
        With .Variable("scoring_multiplier")
            .Action = "set"
            .Int = 1
        End With
    End With
    With .EventName("ramp_award")
        With .Variable("ramp_hits")
            .Action = "add"
            .Int = 1
        End With
        With .Variable("score")
            .Action = "add"
            .Int = "10000 * current_player.scoring_multiplier"
        End With
    End With
End With
```

If the containing mode itself starts on `ball_started`, use its `mode_<name>_started` event for initialization so its players are already listening. See [dynamic values](dynamic_values.md) for expression syntax.

## Read or write from VBScript

```vbscript
SetPlayerState "ramp_hits", 0
Dim hits
hits = GetPlayerState("ramp_hits")
```

Use these only while a player exists. `GetPlayerState` returns `False` for a missing key, so a false or zero result does not prove a variable was initialized.

`GetPlayerStateForPlayer(0, "score")` reads player 1; indices 0–3 represent players 1–4. `SetPlayerStateByPlayer "ramp_hits", 0, 1` writes player 2. Confirm that player exists first. The `players[0].score` runtime expression uses this same zero-based convention. BCP player numbers and the `player_added` event use one-based numbers instead.

## Machine variables

Create custom machine variables during configuration:

```vbscript
With CreateMachineVar("feature_completions")
    .InitialValue = 0
    .ValueType = "int"
    .Persist = True
End With
```

Then update through a variable player:

```vbscript
With .VariablePlayer()
    With .EventName("feature_complete")
        With .Variable("feature_completions")
            .Action = "add_machine"
            .Int = 1
        End With
    End With
End With
```

Use `set_machine` to replace a value. Read it in expressions as `machine.feature_completions`. Defaults are `InitialValue = 0`, `ValueType = "int"` and `Persist = True`; explicitly set persistence for temporary state.

`Glf_Init` reads saved values from `<cGameName>_glf.ini`, including the `MachineVars` and `HighScores` sections. Saved values can replace configured initial values. `Glf_Exit` writes persistent machine variables, so keep the table exit hook and do not treat a runtime update as an immediate disk save. The current INI reader assigns text values without enforcing `ValueType`; test arithmetic and formatting after a reload as well as in a fresh session.

## Observe player-state changes

The listener signature includes a player index:

```vbscript
AddPlayerStateEventListener "score", "my_score_display", 0, "OnScoreChanged", 1000, Null

Sub OnScoreChanged(args)
    Dim change
    change = args(1)
    ' change(0): variable name
    ' change(1): new value
    ' change(2): previous value
End Sub
```

Use the same key to remove it:

```vbscript
RemovePlayerStateEventListener "score", "my_score_display"
```

These are player-state callbacks, not pin-event callbacks. Updates can be delayed/coalesced, and a player turn refresh also emits state. Do not use a display listener as a reliable count of every switch hit. For gameplay, award and count from named rule events.

The current filtering also admits a listener whose configured index equals the current player's index; it is not a strict arbitrary-player subscription in every case. Test multi-player displays, or use GLF's segment display facilities.

Source: `scripts/src/player-state.vbs`, `game.vbs`, `commands/variable-player.vbs` and the machine-variable read/write functions in `_.main.vbs`. Working examples: both games' `_configuration.vbs` and `modes/score.vbs`.
