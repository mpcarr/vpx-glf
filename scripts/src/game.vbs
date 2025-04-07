

'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub Glf_AddPlayer()
    Dim kwargs : Set kwargs = GlfKwargs()
    With kwargs
        .Add "num", -1
    End With
    Select Case UBound(glf_playerState.Keys())
        Case -1:
            kwargs("num") = 1
            DispatchPinEvent "player_added", kwargs
            glf_playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            SetPlayerStateByPlayer GLF_SCORE, 0, 0
            SetPlayerStateByPlayer "number", 1, 0
            Glf_BcpAddPlayer 1
            glf_currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 2
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 2", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 1
                SetPlayerStateByPlayer "number", 2, 1
                Glf_BcpAddPlayer 2
            End If
        Case 1:
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 3
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 3", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 2
                SetPlayerStateByPlayer "number", 3, 2
                Glf_BcpAddPlayer 3
            End If     
        Case 2:   
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 4
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 4", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 3
                SetPlayerStateByPlayer "number", 4, 3
                Glf_BcpAddPlayer 4
            End If  
            glf_canAddPlayers = False
    End Select
End Sub

Function Glf_InitNewPlayer()
    Dim state : Set state = CreateObject("Scripting.Dictionary")
    state.Add GLF_SCORE, 0
    Glf_MonitorPlayerStateUpdate GLF_SCORE, 0
    state.Add GLF_INITIALS, ""
    Glf_MonitorPlayerStateUpdate GLF_INITIALS, ""
    state.Add GLF_CURRENT_BALL, 1
    Glf_MonitorPlayerStateUpdate GLF_CURRENT_BALL, 1
    state.Add "extra_balls", 0
    Glf_MonitorPlayerStateUpdate "extra_balls", 0
    Dim i
    Dim init_var_keys : init_var_keys = glf_initialVars.Keys()
    Dim init_var_items : init_var_items = glf_initialVars.Items()
    For i=0 To UBound(init_var_keys)
        state.Add init_var_keys(i), init_var_items(i)
        Glf_MonitorPlayerStateUpdate init_var_keys(i), init_var_items(i)
    Next
    Set Glf_InitNewPlayer = state
End Function


'****************************
' Setup Player
' Event Listeners:  
    AddPinEventListener GLF_GAME_STARTED,   "start_game_setup",   "Glf_SetupPlayer", 1000, Null
    AddPinEventListener GLF_NEXT_PLAYER,    "next_player_setup",  "Glf_SetupPlayer", 1000, Null
'
'*****************************
Function Glf_SetupPlayer(args)
    EmitAllglf_playerEvents()
End Function

'****************************
' StartGame
'
'*****************************

AddPinEventListener "request_to_start_game", "request_to_start_game_ball_controller", "Glf_BallController", 30, Null
Function Glf_BallController(args)
    Dim balls_in_trough : balls_in_trough = 0
    If glf_troughSize = 1 Then
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 2 Then
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 3 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 4 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 5 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 6 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 7 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough7.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 8 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough7.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If Drain.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
  
    If glf_troughSize <> balls_in_trough Then
        Glf_BallController = False
        Exit Function
    End If

    Glf_BallController = True
End Function

AddPinEventListener "request_to_start_game", "request_to_start_game_result", "Glf_StartGame", 20, Null

Function Glf_StartGame(args)
    If args(1) = True And glf_gameStarted = False Then
        Glf_AddPlayer()
        glf_gameStarted = True
        DispatchPinEvent GLF_GAME_START, Null
        If useBcp Then
            bcpController.Send "player_turn_start?player_num=int:1"
            bcpController.Send "ball_start?player_num=int:1&ball=int:1"
            bcpController.SendPlayerVariable "number", 1, 0
        End If
        SetDelay GLF_GAME_STARTED, "Glf_DispatchGameStarted", Null, 50
    End If
End Function

Sub Glf_EndBall()

    glf_BIP = 0
    DispatchPinEvent GLF_BALL_WILL_END, Null
    DispatchQueuePinEvent GLF_BALL_ENDING, Null
    Dim device
    For Each device in glf_flippers.Items()
        device.Disable()
    Next
    For Each device in glf_autofiredevices.Items()
        device.Disable()
    Next

End Sub

Public Function Glf_DispatchGameStarted(args)
    DispatchPinEvent GLF_GAME_STARTED, Null
End Function


'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener GLF_GAME_STARTED, "start_game_release_ball",   "Glf_ReleaseBall", 20, True
AddPinEventListener GLF_NEXT_PLAYER, "next_player_release_ball",   "Glf_ReleaseBall", 20, True
'
'*****************************
Function Glf_ReleaseBall(args)
    Dim kwargs
    Set kwargs = GlfKwargs()
    If Not IsNull(args) Then
        If args(0) = True Then
            kwargs.Add "new_ball", True
        End If
    End If
    DispatchQueuePinEvent "balldevice_trough_ball_eject_attempt", kwargs
End Function


'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener "balldevice_trough_ball_eject_attempt", "trough_eject",  "Glf_TroughReleaseBall", 20, Null
'
'*****************************
Function Glf_TroughReleaseBall(args)

    If Not IsNull(args) Then
        If IsObject(args(1)) Then
            If args(1)("new_ball") = True Then
                glf_BIP = glf_BIP + 1
                DispatchPinEvent GLF_BALL_STARTED, Null
                If useBcp Then
                    bcpController.SendPlayerVariable GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL), GetPlayerState(GLF_CURRENT_BALL)-1
                    bcpController.SendPlayerVariable GLF_SCORE, GetPlayerState(GLF_SCORE), GetPlayerState(GLF_SCORE)
                End If
            End If
        End If
    End If
    Glf_WriteDebugLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    glf_plunger.AddIncomingBalls = 1
    swTrough1.kick 90, 10
    DispatchPinEvent "trough_eject", Null
    Glf_WriteDebugLog "Release Ball", "Just Kicked"
    If Not IsNull(args) Then
		If IsObject(args(1)) Then
			Set Glf_TroughReleaseBall = args(1)
		Else
			Glf_TroughReleaseBall = args(1)
		End If
	Else
		Glf_TroughReleaseBall = Null
	End If
End Function

'****************************
' Ball Drain
' Event Listeners:      
    AddPinEventListener GLF_BALL_DRAIN, "ball_drain", "Glf_Drain", 20, Null
'
'*****************************
Function Glf_Drain(args)
    
    If Not glf_gameTilted Then
        Dim ballsToSave : ballsToSave = args(1) 
        Glf_WriteDebugLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
        Glf_WriteDebugLog "end_of_ball, balls in play", CStr(glf_BIP)
        If ballsToSave <= 0 Then
            Exit Function
        End If

        glf_BIP = glf_BIP - 1
        glf_debugLog.WriteToLog "Trough", "Ball Drained: BIP: " & glf_BIP

        If glf_BIP > 0 Then
            Exit Function
        End If
        
        DispatchPinEvent GLF_BALL_WILL_END, Null
        DispatchQueuePinEvent GLF_BALL_ENDING, Null
    End If
    
End Function

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener GLF_BALL_ENDING, "ball_will_end", "Glf_BallWillEnd", 10, Null
'
'*****************************
Function Glf_BallWillEnd(args)
    DispatchPinEvent GLF_BALL_ENDED, Null
    If Not IsNull(args) Then
		If IsObject(args(1)) Then
			Set Glf_BallWillEnd = args(1)
		Else
			Glf_BallWillEnd = args(1)
		End If
	Else
		Glf_BallWillEnd = Null
	End If
End Function

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener GLF_BALL_ENDED, "end_of_ball", "Glf_EndOfBall", 20, Null
'
'*****************************
Function Glf_EndOfBall(args)

    If GetPlayerState("extra_balls") > 0 Then
        'self.debug_log("Awarded extra ball to Player %s. Shoot Again", self.player.index + 1)
        'self.player.extra_balls -= 1
        SetPlayerState "extra_balls", GetPlayerState("extra_balls") - 1
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000
        Exit Function
    End If
    
    SetPlayerState GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            If UBound(glf_playerState.Keys()) > 0 Then
                glf_currentPlayer = "PLAYER 2"
            End If
        Case "PLAYER 2":
            If UBound(glf_playerState.Keys()) > 1 Then
                glf_currentPlayer = "PLAYER 3"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 3":
            If UBound(glf_playerState.Keys()) > 2 Then
                glf_currentPlayer = "PLAYER 4"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 4":
            glf_currentPlayer = "PLAYER 1"
    End Select
    
    If useBcp Then
        bcpController.SendPlayerVariable "number", Getglf_currentPlayerNumber(), previousPlayerNumber
    End If
    If GetPlayerState(GLF_CURRENT_BALL) > glf_ballsPerGame Then
        Dim device
        For Each device in glf_ball_devices.Items()
            If device.HasBall() Then
                device.EjectAll()
            End If
        Next
        DispatchPinEvent "game_will_end", Null
        DispatchQueuePinEvent "game_ending", Null
    Else
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000 
    End If
    


End Function

'****************************
' Game Over
' Event Listeners:      
AddPinEventListener "game_ending", "glf_game_over", "Glf_EndGame", 20, Null
'
'*****************************
Function Glf_EndGame(args)
    If GetPlayerStateForPlayer("0", "score") = False Then
        glf_machine_vars("player1_score").Value = 0
    Else
        glf_machine_vars("player1_score").Value = GetPlayerStateForPlayer("0", "score")
    End If
    If GetPlayerStateForPlayer("1", "score") = False Then
        glf_machine_vars("player2_score").Value = 0
    Else
        glf_machine_vars("player2_score").Value = GetPlayerStateForPlayer("1", "score")
    End If
    If GetPlayerStateForPlayer("2", "score") = False Then
        glf_machine_vars("player3_score").Value = 0
    Else
        glf_machine_vars("player3_score").Value = GetPlayerStateForPlayer("2", "score")
    End If
    If GetPlayerStateForPlayer("3", "score") = False Then
        glf_machine_vars("player4_score").Value = 0
    Else
        glf_machine_vars("player4_score").Value = GetPlayerStateForPlayer("3", "score")
    End If
    glf_gameStarted = False
    glf_currentPlayer = Null
    glf_playerState.RemoveAll()

    DispatchPinEvent "game_ended", Null
End Function

Public Function EndOfBallNextPlayer(args)
    DispatchPinEvent GLF_NEXT_PLAYER, Null
End Function
