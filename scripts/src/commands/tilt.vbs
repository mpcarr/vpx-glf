Class GlfTilt

    Private m_name
    Private m_priority
    Private m_base_device
    Private m_reset_warnings_events
    Private m_tilt_events
    Private m_tilt_warning_events
    Private m_tilt_slam_tilt_events
    Private m_settle_time
    Private m_warnings_to_tilt
    Private m_multiple_hit_window
    Private m_tilt_warning_switch
    Private m_tilt_switch
    Private m_slam_tilt_switch
    Private m_last_tilt_warning_switch 
    Private m_last_warning
    Private m_balls_to_collect
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "tilt_settle_ms_remaining":
                GetValue = TiltSettleMsRemaining()
            Case "tilt_warnings_remaining":
                GetValue = TiltWarningsRemaining()
        End Select
    End Property

    Public Property Let ResetWarningEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_warnings_events.Add newEvent.Raw, newEvent
        Next
    End Property
    'Public Property Let TiltEvents(value)
    '    Dim x
    '    For x=0 to UBound(value)
    '        Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
    '        m_tilt_events.Add newEvent.Raw, newEvent
    '    Next
    'End Property
    'Public Property Let TiltWarningEvents(value)
    '    Dim x
    '    For x=0 to UBound(value)
    '        Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
    '        m_tilt_warning_events.Add newEvent.Raw, newEvent
    '    Next
    'End Property
    'Public Property Let SlamTiltEvents(value)
    '    Dim x
    '    For x=0 to UBound(value)
    '        Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
    '        m_tilt_slam_tilt_events.Add newEvent.Raw, newEvent
    '    Next
    'End Property
    Public Property Let SettleTime(value): Set m_settle_time = CreateGlfInput(value): End Property
    Public Property Let WarningsToTilt(value): Set m_warnings_to_tilt = CreateGlfInput(value): End Property
    Public Property Let MultipleHitWindow(value): Set m_multiple_hit_window = CreateGlfInput(value): End Property
    'Public Property Let TiltWarningSwitch(value): m_tilt_warning_switch = value: End Property
    'Public Property Let TiltSwitch(value): m_tilt_switch = value: End Property
    'Public Property Let SlamTiltSwitch(value): m_slam_tilt_switch = value: End Property

    Private Property Get TiltSettleMsRemaining()
        TiltSettleMsRemaining = 0
        If m_last_tilt_warning_switch > 0 Then
            Dim delta
            delta = m_settle_time.Value - (gametime - m_last_tilt_warning_switch)
            If delta > 0 Then
                TiltSettleMsRemaining = delta
            End If
        End If
    End Property

    Private Property Get TiltWarningsRemaining() 
        TiltWarningsRemaining = 0

        If glf_gameStarted Then
            TiltWarningsRemaining = m_warnings_to_tilt.Value() - GetPlayerState("tilt_warnings")
        End If   
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(mode)
        m_name = "tilt_" & mode.name
        m_priority = mode.Priority
        Set m_reset_warnings_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_warning_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_slam_tilt_events = CreateObject("Scripting.Dictionary")
        Set m_settle_time = CreateGlfInput(0)
        Set m_warnings_to_tilt = CreateGlfInput(0)
        Set m_multiple_hit_window = CreateGlfInput(0)
        m_tilt_switch = Empty
        m_tilt_warning_switch = Empty
        m_slam_tilt_switch = Empty
        m_last_tilt_warning_switch = 0
        m_last_warning = 0
        m_balls_to_collect = 0
        Set m_base_device = (new GlfBaseModeDevice)(mode, "tilt", Me)
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_reset_warnings_events.Keys()
            AddPinEventListener m_reset_warnings_events(evt).EventName, m_name & "_" & evt & "_reset_warnings", "TiltHandler", m_priority+m_reset_warnings_events(evt).Priority, Array("reset_warnings", Me, m_reset_warnings_events(evt))
        Next
        For Each evt in m_tilt_events.Keys()
            AddPinEventListener m_tilt_events(evt).EventName, m_name & "_" & evt & "_tilt", "TiltHandler", m_priority+m_tilt_events(evt).Priority, Array("tilt", Me, m_tilt_events(evt))
        Next
        For Each evt in m_tilt_warning_events.Keys()
            AddPinEventListener m_tilt_warning_events(evt).EventName, m_name & "_" & evt & "_tilt_warning", "TiltHandler", m_priority+m_tilt_warning_events(evt).Priority, Array("tilt_warning", Me, m_tilt_warning_events(evt))
        Next
        For Each evt in m_tilt_slam_tilt_events.Keys()
            AddPinEventListener m_tilt_slam_tilt_events(evt).EventName, m_name & "_" & evt & "_slam_tilt", "TiltHandler", m_priority+m_tilt_slam_tilt_events(evt).Priority, Array("slam_tilt", Me, m_tilt_slam_tilt_events(evt))
        Next
        
        AddPinEventListener  "s_tilt_warning_active", m_name & "_tilt_warning_switch_active", "TiltHandler", m_priority, Array("_tilt_warning_switch_active", Me)
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_reset_warnings_events.Keys()
            RemovePinEventListener m_reset_warnings_events(evt).EventName, m_name & "_" & evt & "_reset_warnings"
        Next
        For Each evt in m_tilt_events.Keys()
            RemovePinEventListener m_tilt_events(evt).EventName, m_name & "_" & evt & "_tilt"
        Next
        For Each evt in m_tilt_warning_events.Keys()
            RemovePinEventListener m_tilt_warning_events(evt).EventName, m_name & "_" & evt & "_tilt_warning"
        Next
        For Each evt in m_tilt_slam_tilt_events.Keys()
            RemovePinEventListener m_tilt_slam_tilt_events(evt).EventName, m_name & "_" & evt & "_slam__tilt"
        Next

        RemovePinEventListener "s_tilt_warning_active", m_name & "_tilt_warning_switch_active"

    End Sub

    Public Sub TiltWarning()
        'Process a tilt warning.
        'If the number of warnings is than the number to cause a tilt, a tilt will be
        'processed.

        m_last_tilt_warning_switch = gametime
        If glf_gameStarted = False Or glf_gameTilted = True Then
            Exit Sub
        End If
        Log "Tilt Warning"
        m_last_warning = gametime
        SetPlayerState "tilt_warnings", GetPlayerState("tilt_warnings")+1
        Dim warnings : warnings = GetPlayerState("tilt_warnings")
        Dim warnings_to_tilt : warnings_to_tilt = m_warnings_to_tilt.Value()
        If warnings>=warnings_to_tilt Then
            Tilt()
        Else
            Dim kwargs
            Set kwargs = GlfKwargs()
            With kwargs
                .Add "warnings", warnings
                .Add "warnings_remaining", warnings_to_tilt - warnings
            End With
            DispatchPinEvent "tilt_warning", kwargs
            DispatchPinEvent "tilt_warning_" & warnings, Null
        End If
    End Sub

    Public Sub ResetWarnings()
        'Reset the tilt warnings for the current player.
        If glf_gamestarted = False or glf_gameEnding = True Then
            Exit Sub
        End If
        SetPlayerState "tilt_warnings", 0
    End Sub

    Public Sub Tilt()
        'Cause the ball to tilt.
        'This will post an event called *tilt*, set the game mode's ``tilted``
        'attribute to *True*, disable the flippers and autofire devices, end the
        'current ball, and wait for all the balls to drain.
        If glf_gameStarted = False or glf_gameTilted=True or glf_gameEnding = True Then
            Exit Sub
        End If
        glf_gametilted = True
        m_balls_to_collect = glf_BIP
        Log "Processing Tilt. Balls to collect: " & m_balls_to_collect
        DispatchPinEvent "tilt", Null
        AddPinEventListener GLF_BALL_ENDING, m_name & "_ball_ending", "TiltHandler", 20, Array("tilt_ball_ending", Me)
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "TiltHandler", 999999, Array("tilt_ball_drain", Me)
        Glf_EndBall()
    End Sub

    Public Sub TiltedBallDrain(unclaimed_balls)
        Log "Tilted ball drain, unclaimed balls: " & unclaimed_balls
        m_balls_to_collect = m_balls_to_collect - unclaimed_balls
        Log "Tilted ball drain. Balls to collect: " & m_balls_to_collect
        If m_balls_to_collect <= 0 Then
            TiltDone()
        End If
    End Sub

    Public Sub HandleTiltWarningSwitch()
        Log "Handling Tilt Warning Switch"
        If m_last_warning = 0 Or (m_last_warning + m_multiple_hit_window.Value()) <= gametime Then
            TiltWarning()
        End If
    End Sub

    Public Sub BallEndingTilted()
        If m_balls_to_collect<=0 Then
            TiltDone()
        End If
    End Sub

    Public Sub TiltDone()
        If TiltSettleMsRemaining() > 0 Then
            SetDelay "delay_tilt_clear", "TiltHandler", Array(Array("tilt_done", Me), Null), TiltSettleMsRemaining()
            Exit Sub
        End If
        Log "Tilt Done"
        RemovePinEventListener GLF_BALL_ENDING, m_name & "_ball_ending"
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        glf_gameTilted = False
        DispatchPinEvent m_name & "_clear", Null
    End Sub
    
    Public Sub SlamTilt()
        'Process a slam tilt.
        'This method posts the *slam_tilt* event and (if a game is active) sets
        'the game mode's ``slam_tilted`` attribute to *True*.
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TiltHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim tilt : Set tilt = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set TiltHandler = kwargs
                Else
                    TiltHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "_tilt_warning_switch_active":
            tilt.HandleTiltWarningSwitch
        Case "tilt_ball_ending"
            kwargs.Add "wait_for", tilt.Name & "_clear"
            tilt.BallEndingTilted
        Case "tilt_ball_drain"
            tilt.TiltedBallDrain kwargs
            kwargs = kwargs -1
        Case "tilt_done"
            tilt.TiltDone
        Case "reset_warnings"
            tilt.ResetWarnings
    End Select

    If IsObject(args(1)) Then
        Set TiltHandler = kwargs
    Else
        TiltHandler = kwargs
    End If
End Function
