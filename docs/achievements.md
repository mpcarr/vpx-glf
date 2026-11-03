# Achievement tracking and implementation status

The current `scripts/src/commands/achievements.vbs` is unfinished. Its `Complete()` method contains only a TODO, and its listener registration names `AchievementsEventHandler` while the function in that file is named `AchievementsHandler`. The mode API does not expose an `Achievements` collection. Do not treat this file as a working configurable achievement device.

## Track a completed feature with player state

For a simple once-per-player award, initialize a custom variable during configuration:

```vbscript
Glf_SetInitialPlayerVar "challenge_achievement", 0
```

Inside a mode that is active when the completion event arrives:

```vbscript
With .VariablePlayer()
    With .EventName("challenge_complete{current_player.challenge_achievement == 0}")
        With .Variable("score")
            .Action = "add"
            .Int = 50000
        End With
        With .Variable("challenge_achievement")
            .Action = "set"
            .Int = 1
        End With
    End With
End With
```

The condition is evaluated before that entry's updates. On subsequent completion events, the flag prevents another award. The flag belongs to each player and resets for a new game. This example awards an additional 50,000 if used alongside the complete-mode tutorial's existing completion award; remove one award if you want only one.

Use [shots and profiles](shot.md) for visible progress or a [state machine](state-machine.md) for multi-stage qualification. Use a persistent [machine variable](player-state.md) only if the achievement should be shared and retained between sessions.

This is a supported composition of existing components, not a completed implementation of `GlfAchievements`. Test multiple players and repeated completion events before using it in a rule set.
