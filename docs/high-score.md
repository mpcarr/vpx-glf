# High Score Configuration

The high score configuration allows you to set up and customize high score tracking in your game. High scores are used to track and display player scores across multiple games.

## Configuration Options

Within your ConfigureGlfDevices function, you need to enable highscores. You can create different categories and defaults. Initial values are set first time the game runs. After that, values are read from the machines ini file.

```vbscript
    With EnableGlfHighScores()
        With .Categories()
            .Add "score", Array("GRAND CHAMPION", "HIGH SCORE 1", "HIGH SCORE 2", "HIGH SCORE 3") 
        End With
        With .Defaults("score")
            .Add "DAN", 9000000
            .Add "MPC", 7000000
            .Add "AVE", 5000000
            .Add "DIG", 3000000
        End With
        .EnterInitialsTimeout = 65000
    End With
```
