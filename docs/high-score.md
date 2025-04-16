# High Score Configuration

The high score configuration allows you to set up and customize high score tracking in your pinball machine. High scores are used to track and display player achievements across multiple games.

## Configuration Options

### Basic Configuration
```vbscript
' Basic high score configuration
Dim high_score_mode : Set high_score_mode = CreateGlfMode("glf_high_scores", 80)
high_score_mode.StartEvents = Array("game_ending")
high_score_mode.StopEvents = Array("high_score_complete")
high_score_mode.UseWaitQueue = True
Dim high_score : Set high_score = (new GlfHighScore)(high_score_mode)
high_score.Debug = True
Set high_score_mode.HighScore = high_score
Set glf_highscore = high_score
```

### Advanced Configuration
```vbscript
' Advanced high score configuration
Dim high_score_mode : Set high_score_mode = CreateGlfMode("glf_high_scores", 80)
high_score_mode.StartEvents = Array("game_ending")
high_score_mode.StopEvents = Array("high_score_complete")
high_score_mode.UseWaitQueue = True
Dim high_score : Set high_score = (new GlfHighScore)(high_score_mode)
high_score.Debug = True

' Configure high score categories
With high_score.Categories
    With .Category("score")
        .Add "GRAND CHAMPION", 1000000
        .Add "HIGH SCORE", 500000
        .Add "SCORE", 100000
    End With
    With .Category("time")
        .Add "BEST TIME", 120
        .Add "TIME", 180
    End With
End With

' Configure default values
With high_score.Defaults("score")
    .Add "GRAND CHAMPION", 1000000
    .Add "HIGH SCORE", 500000
    .Add "SCORE", 100000
End With

' Configure player variables
With high_score.Vars("score")
    .Add "value", "score"
End With

' Configure display settings
high_score.AwardSlideDisplayTime = 4000
high_score.EnterInitialsTimeout = 20000

' Configure reset events
high_score.ResetHighScoreEvents = Array("reset_high_scores")

Set high_score_mode.HighScore = high_score
Set glf_highscore = high_score
```

## Property Descriptions

### High Score Settings
- `Categories`: Dictionary of high score categories with their labels and positions
- `Defaults`: Dictionary of default values for each category
- `Vars`: Dictionary of player variables to track for each category
- `AwardSlideDisplayTime`: Time in milliseconds to display award slides (Default: 4000)
- `EnterInitialsTimeout`: Time in milliseconds to wait for initials input (Default: 20000)
- `ResetHighScoreEvents`: Array of events that reset high scores (Default: Empty)
- `Debug`: Boolean to enable debug logging for this high score system (Default: False)

## Example Configurations

### Basic High Score Example
```vbscript
' Basic high score configuration
Function EnableGlfHighScores()
    Dim high_score_mode : Set high_score_mode = CreateGlfMode("glf_high_scores", 80)
    high_score_mode.StartEvents = Array("game_ending")
    high_score_mode.StopEvents = Array("high_score_complete")
    high_score_mode.UseWaitQueue = True
    Dim high_score : Set high_score = (new GlfHighScore)(high_score_mode)
    high_score.Debug = True
    Set high_score_mode.HighScore = high_score
    Set glf_highscore = high_score
    Set EnableGlfHighScores = glf_highscore
End Function
```

### Advanced High Score Example
```vbscript
' Advanced high score configuration
Function EnableGlfHighScores()
    Dim high_score_mode : Set high_score_mode = CreateGlfMode("glf_high_scores", 80)
    high_score_mode.StartEvents = Array("game_ending")
    high_score_mode.StopEvents = Array("high_score_complete")
    high_score_mode.UseWaitQueue = True
    Dim high_score : Set high_score = (new GlfHighScore)(high_score_mode)
    high_score.Debug = True

    ' Configure high score categories
    With high_score.Categories
        With .Category("score")
            .Add "GRAND CHAMPION", 1000000
            .Add "HIGH SCORE", 500000
            .Add "SCORE", 100000
        End With
        With .Category("time")
            .Add "BEST TIME", 120
            .Add "TIME", 180
        End With
    End With

    ' Configure default values
    With high_score.Defaults("score")
        .Add "GRAND CHAMPION", 1000000
        .Add "HIGH SCORE", 500000
        .Add "SCORE", 100000
    End With

    ' Configure player variables
    With high_score.Vars("score")
        .Add "value", "score"
    End With

    ' Configure display settings
    high_score.AwardSlideDisplayTime = 4000
    high_score.EnterInitialsTimeout = 20000

    ' Configure reset events
    high_score.ResetHighScoreEvents = Array("reset_high_scores")

    Set high_score_mode.HighScore = high_score
    Set glf_highscore = high_score
    Set EnableGlfHighScores = glf_highscore
End Function
```

## High Score System

The high score system manages player achievements with the following features:

- High scores are tracked across multiple games
- Scores are stored in an INI file for persistence
- Players can enter their initials for high scores
- Award slides are displayed for high scores
- High scores can be reset with specific events

## Events

The high score system generates and responds to the following events:
- `game_ending`: Starts the high score mode
- `high_score_complete`: Ends the high score mode
- `high_score_enter_initials`: Triggers the initials input process
- `text_input_high_score_complete`: Indicates initials input is complete
- `high_score_award_display`: Triggers the award display
- `award_display_complete`: Indicates the award display is complete
- `category_award_display`: Category-specific award display event
- `award_award_display`: Award-specific display event

## Default Behavior

By default, high scores are configured with:
- No categories defined
- Award slide display time of 4000ms
- Initials input timeout of 20000ms
- Debug logging disabled

## Notes

- High scores are managed within the context of a mode
- The high score system automatically handles score comparison and storage
- Debug logging can be enabled to track high score operations
- High scores can be configured to work with specific game modes or features
- The `Categories` property defines the structure of high scores
- The `Defaults` property sets default values for high scores
- The `Vars` property links player variables to high score categories
- High scores are stored in an INI file for persistence
- The system automatically handles score comparison and storage
- Players can enter their initials for high scores
- Award slides are displayed for high scores
- High scores can be reset with specific events 