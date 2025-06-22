Function GlfGameSettings()
    Set glf_game = (new GlfGame)()
	Set GlfGameSettings = glf_game
End Function

Class GlfGame

    Private m_balls_per_game

    Public Property Get BallsPerGame()
        BallsPerGame = m_balls_per_game.Value()
    End Property
    Public Property Let BallsPerGame(input)
        Set m_balls_per_game = CreateGlfInput(input)
    End Property

	Public default Function init()

        Set m_balls_per_game = CreateGlfInput(3)
       
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml : yaml = ""
        ToYaml = yaml
    End Function

End Class