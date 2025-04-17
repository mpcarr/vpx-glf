# Tutorial - Scoring

Let's add scoring to our pinball table. Scoring is a fundamental part of any pinball game, allowing players to track their progress and compete for high scores.

## Configuration

To configure scoring in your game, we need to set up a variable player that will handle score tracking. The variable player allows you to define variables that can be modified by events, such as adding points when certain actions occur.

### Score Configuration

Create the following configuration to set up a scoring system:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        ' Add scoring configuration
        With VariablePlayer("base")
            
            ' Define events that will add points to the score
            With .EventName("add_score_1000")
                With .Variable("score")
                    .Action = "add"
                    .Int = 1000
                End With
            End With
            
            ' Add more scoring events for different actions
            With .EventName("add_score_500")
                With .Variable("score")
                    .Action = "add"
                    .Int = 500
                End With
            End With
            
            With .EventName("add_score_100")
                With .Variable("score")
                    .Action = "add"
                    .Int = 100
                End With
            End With
        End With
        
    End With
End Sub
```

## Advanced Scoring Features

### Bonus Scoring

You can implement bonus scoring that increases based on player performance:

```
Sub CreateBaseMode()
    With CreateGlfMode("base", 1000)
        With VariablePlayer("base")
            ' Define a bonus multiplier variable
            With .Variable("bonus_multiplier")
                .InitialValue = 1
            End With
            
            ' Event to increase the bonus multiplier
            With .EventName("increase_bonus")
                With .Variable("bonus_multiplier")
                    .Action = "add"
                    .Int = 1
                End With
            End With
            
            ' Event to add bonus points (multiplied by the current multiplier)
            With .EventName("add_bonus_points")
                With .Variable("score")
                    .Action = "add"
                    .Expression = "1000 * current_player.bonus_multiplier"
                End With
            End With
        End With
    End With
End Sub
```