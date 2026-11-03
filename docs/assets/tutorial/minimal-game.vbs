' GLF tutorial: completed core game (chapters 1-6).
' Requires the VPX objects and collections listed in tutorial/index.md.
' Put vpx-glf.vbs in the VPX Scripts folder. Do not paste a second GLF copy.
Const cGameName = "GLFFirstGame"
Const BallSize = 50
Const BallMass = 1
Const tnob = 1
Const lob = 0
Dim gBot
Dim tablewidth : tablewidth = Table1.Width
Dim tableheight : tableheight = Table1.Height

ExecuteGlobal GetTextFile("controller.vbs")
ExecuteGlobal GetTextFile("core.vbs")
ExecuteGlobal GetTextFile("vpx-glf.vbs")

Sub Table1_Init()
    ConfigureGlfDevices
    Glf_Init Table1
End Sub

Sub Table1_Exit()
    Glf_Exit
End Sub

Sub Table1_KeyDown(ByVal keycode)
    Glf_KeyDown keycode
    If keycode = PlungerKey Then Plunger.Pullback
End Sub

Sub Table1_KeyUp(ByVal keycode)
    Glf_KeyUp keycode
    If keycode = PlungerKey Then Plunger.Fire
End Sub

Sub Table1_OptionEvent(ByVal eventId)
    Glf_Options eventId
End Sub

Sub ConfigureGlfDevices()
    ConfigureTutorialPlunger
    ConfigureTutorialFlippers
    CreateTutorialScoring
    CreateTutorialLights
End Sub

Sub ConfigureTutorialPlunger()
    With CreateGlfBallDevice("plunger")
        .BallSwitches = Array("s_plunger")
        .DefaultDevice = True
        .MechanicalEject = True
        .EjectTimeout = 2000
        .EjectCallback = "TutorialPlungerEject"
    End With
End Sub

Sub TutorialPlungerEject(ball)
    If IsObject(ball) Then
        ball.VelX = 0
        ball.VelY = -45
    End If
End Sub

Sub ConfigureTutorialFlippers()
    With CreateGlfFlipper("left")
        .Switch = "s_left_flipper"
        .ActionCallback = "TutorialLeftFlipper"
    End With
    With CreateGlfFlipper("right")
        .Switch = "s_right_flipper"
        .ActionCallback = "TutorialRightFlipper"
    End With
End Sub

Sub TutorialLeftFlipper(enabled)
    If enabled Then
        LeftFlipper.RotateToEnd
    Else
        LeftFlipper.RotateToStart
    End If
End Sub

Sub TutorialRightFlipper(enabled)
    If enabled Then
        RightFlipper.RotateToEnd
    Else
        RightFlipper.RotateToStart
    End If
End Sub

Sub CreateTutorialScoring()
    With CreateGlfMode("tutorial_scoring", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .VariablePlayer()
            With .EventName("s_score_active")
                With .Variable("score")
                    .Action = "add"
                    .Int = 1000
                End With
            End With
        End With
    End With
End Sub

Sub TutorialDisplayTimer_Timer()
    If glf_gameStarted And Not IsNull(glf_currentPlayer) Then
        ScoreText.Text = "PLAYER " & CStr(GetPlayerState("number")) & _
            "   BALL " & CStr(GetPlayerState("ball")) & _
            "   SCORE " & CStr(GetPlayerState("score"))
    Else
        ScoreText.Text = "PRESS START"
    End If
End Sub

Sub CreateTutorialLights()
    With CreateGlfMode("tutorial_lights", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_will_end", "tilt")
        With .LightPlayer()
            With .EventName("mode_tutorial_lights_started")
                With .Lights("GI")
                    .Color = "ffffff"
                End With
            End With
        End With
    End With
End Sub
