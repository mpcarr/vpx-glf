---
title: Using dynamic runtime values in config files
---

# Using dynamic runtime values in config files


Some config command properties can accept a dynamic values which are evaluated live when the game is running rather than
being hard-coded.

Dynamic values can come from several sources, including player
variables, machine variables.

For example, you might want to have a shot called "jackpot" that
scores a multiplier which is the number of shots made times 100k points.

Without dynamic values, your variable_player (scoring) section would be
static, like this:

``` vbs
With .VariablePlayer()
  With .EventName("shot_jackpot_hit")
    With .Variable("score")
      .Action = "add"
      .Int = 100000
    End With
  End With
End With
```

But let's say you have a player variable called "troll_hits" which
holds the number of trolls hit that you want to multiply by 100,000 when
the shot is made. You can use the "current_player" dynamic value in
your variable_player config like this:

``` vbs
With .VariablePlayer()
  With .EventName("shot_jackpot_hit")
    With .Variable("score")
      .Action = "add"
      .Int = 100000 * current_player.troll_hits
    End With
  End With
End With
```

## Types of dynamic values

You can use the following types of placeholders.

### Current Player Variables

You can access a player variable `X` of the current player using
`current_player.X`. For instance, `current_player.my_player_var` will
access `my_player_var` of the current player. This placeholder is only
available when a game is active.

Common player variables are:

* `current_player.score` - Score of the current player
* `current_player.ball` - Current ball

*TODO: Add more examples of available dynamic values*
