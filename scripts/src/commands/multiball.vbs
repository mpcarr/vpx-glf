Class GlfMultiballs

    Private m_name
    Private m_configname
    Private m_mode
    Private m_priority
    Private m_base_device
    Private m_ball_count
    Private m_ball_locks
    Private m_add_a_ball_events
    Private m_add_a_ball_grace_period
    Private m_add_a_ball_hurry_up_time
    Private m_add_a_ball_shoot_again
    Private m_ball_count_type
    Private m_disable_events
    Private m_enable_events
    Private m_grace_period
    Private m_grace_period_enabled
    Private m_hurry_up
    Private m_hurry_up_enabled
    Private m_replace_balls_in_play
    Private m_reset_events
    Private m_shoot_again
    Private m_source_playfield
    Private m_start_events
    Private m_stop_events
    Private m_balls_added_live
    Private m_balls_live_target
    Private m_enabled
    Private m_shoot_again_enabled
    Private m_queued_balls
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
        End Select
    End Property

    Public Property Let BallCount(value): Set m_ball_count = CreateGlfInput(value): End Property
    Public Property Let AddABallEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_add_a_ball_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let AddABallGracePeriod(value): Set m_add_a_ball_grace_period = CreateGlfInput(value): End Property
    Public Property Let AddABallHurryUpTime(value): Set m_add_a_ball_hurry_up_time = CreateGlfInput(value): End Property
    Public Property Let AddABallShootAgain(value): Set m_add_a_ball_shoot_again = CreateGlfInput(value): End Property
    Public Property Let BallCountType(value): m_ball_count_type = value: End Property
    Public Property Let BallLocks(value): m_ball_locks = value: End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let GracePeriod(value): Set m_grace_period = CreateGlfInput(value): End Property
    Public Property Let HurryUp(value): Set m_hurry_up = CreateGlfInput(value): End Property
    Public Property Let ReplaceBallsInPlay(value): m_replace_balls_in_play = value: End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ShootAgain(value): Set m_shoot_again = CreateGlfInput(value): End Property
    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Raw, newEvent
        Next
    End Property
        
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(name, mode)
        m_name = "multiball_" & name
        m_configname = name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_ball_count = CreateGlfInput(0)
        Set m_add_a_ball_events = CreateObject("Scripting.Dictionary")
        Set m_add_a_ball_grace_period = CreateGlfInput(0)
        Set m_add_a_ball_hurry_up_time = CreateGlfInput(0)
        Set m_add_a_ball_shoot_again = CreateGlfInput(5000)
        m_ball_count_type = "total"
        m_ball_locks = Array()
        Set m_grace_period = CreateGlfInput(0)
        Set m_hurry_up = CreateGlfInput(0)
        m_replace_balls_in_play = False
        Set m_shoot_again = CreateGlfInput(10000)
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")
        m_replace_balls_in_play = False
        m_balls_added_live = 0
        m_balls_live_target = 0
        m_queued_balls = 0
        m_enabled = False
        m_shoot_again_enabled = False
        m_grace_period_enabled = False
        m_hurry_up_enabled = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "multiball", Me)
        glf_multiballs.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        If UBound(m_base_device.EnableEvents.Keys()) = -1 Then
            Enable()
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling " & m_name
        m_enabled = True
        Dim evt
        For Each evt in m_start_events.Keys
            AddPinEventListener m_start_events(evt).EventName, m_name & "_" & evt & "_start", "MultiballsHandler", m_priority+m_start_events(evt).Priority, Array("start", Me, m_start_events(evt))
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset", "MultiballsHandler", m_priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_add_a_ball_events.Keys
            AddPinEventListener m_add_a_ball_events(evt).EventName, m_name & "_" & evt & "_add_a_ball", "MultiballsHandler", m_priority, Array("add_a_ball", Me, m_add_a_ball_events(evt))
        Next
        For Each evt in m_stop_events.Keys
            AddPinEventListener m_stop_events(evt).EventName, m_name & "_" & evt & "_stop", "MultiballsHandler", m_priority+m_stop_events(evt).Priority, Array("stop", Me, m_stop_events(evt))
        Next
    End Sub
    
    Public Sub Disable()
        Log "Disabling " & m_name
        m_enabled = False
        m_balls_added_live = 0
        m_balls_live_target = 0
        m_shoot_again_enabled = False
        StopMultiball()
        Dim evt
        For Each evt in m_start_events.Keys
            RemovePinEventListener m_start_events(evt).EventName, m_name & "_" & evt & "_start"
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_add_a_ball_events.Keys
            RemovePinEventListener m_add_a_ball_events(evt).EventName, m_name & "_" & evt & "_add_a_ball"
        Next
        For Each evt in m_stop_events.Keys
            RemovePinEventListener m_stop_events(evt).EventName, m_name & "_" & evt & "_stop"
        Next
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        'RemoveDelay m_name & "_queued_release"
    End Sub
    
    Private Sub HandleBallsInPlayAndBallsLive()
        'Dim balls_to_replace
        'If m_replace_balls_in_play = True Then
        '    balls_to_replace = glf_BIP
        'Else
        '    balls_to_replace = 0
        'End If
        'Log("Going to add an additional " & balls_to_replace & " balls for replace_balls_in_play")
        m_balls_added_live = 0 
        Dim ball_count_value : ball_count_value = m_ball_count.Value
        If m_ball_count_type = "total" Then
            Log "glf_BIP: " & glf_BIP
            If ball_count_value > glf_BIP Then
                m_balls_added_live = ball_count_value - glf_BIP
                'glf_BIP = m_ball_count
            End If
            m_balls_live_target = ball_count_value
        Else
            m_balls_added_live = ball_count_value
            'glf_BIP = glf_BIP + m_balls_added_live
            m_balls_live_target = glf_BIP + m_balls_added_live
        End If

    End Sub

    Public Function BallsDrained(balls)
        If m_shoot_again_enabled Then
            balls = BallDrainShootAgain(balls)
        Else
            BallDrainCountBalls(balls)
        End If
        BallsDrained = balls
    End Function

    Public Sub Start()
        ' Start multiball.
        If not m_enabled Then
            Exit Sub
        End If

        If m_balls_live_target > 0 Then
            Log("Cannot start MB because " & m_balls_live_target & " are still in play")
            Exit Sub
        End If

        m_shoot_again_enabled = True

        HandleBallsInPlayAndBallsLive()
        Log("Starting multiball with " & m_balls_live_target & " balls (added " & m_balls_added_live & ")")
        'msgbox("Starting multiball with " & m_balls_live_target & " balls (added " & m_balls_added_live & ")")    
        Dim balls_added : balls_added = 0

        'eject balls from locks
        Dim ball_lock
        For Each ball_lock in m_ball_locks
            Dim available_balls : available_balls = glf_ball_devices(ball_lock).Balls()
            If available_balls > 0 Then
                glf_ball_devices(ball_lock).EjectAll()
            End If
            balls_added = balls_added + available_balls
        Next

        glf_BIP = m_balls_live_target

        'request remaining balls
        m_queued_balls = (m_balls_added_live - balls_added)
        If m_queued_balls > 0 Then
            SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me),Null), 1000
        End If

        If m_shoot_again.Value = 0 Then
            'No shoot again. Just stop multiball right away
            StopMultiball()
        else
            'Enable shoot again
            TimerStart()
        End If
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "MultiballsHandler", m_priority, Array("drain", Me)

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls", m_balls_live_target
        End With
        DispatchPinEvent m_name & "_started", kwargs
    End Sub

    Sub TimerStart()
        DispatchPinEvent "ball_save_" & m_configname & "_timer_start", Null 'desc: The multiball ball save called (name) has just start its countdown timer.
        StartShootAgain m_shoot_again.Value, m_grace_period.Value, m_hurry_up.Value
    End Sub

    Sub StartShootAgain(shoot_again_ms, grace_period_ms, hurry_up_time_ms)
        'Set callbacks for shoot again, grace period, and hurry up, if values above 0 are provided.
        'This is started for both beginning multiball ball save and add a ball ball save
        If shoot_again_ms > 0 Then
            Log("Starting ball save timer: " & shoot_again_ms)
            SetDelay m_name&"_disable_shoot_again", "MultiballsHandler" , Array(Array("stop", Me),Null), shoot_again_ms+grace_period_ms
        End If
        If grace_period_ms > 0 Then
            m_grace_period_enabled = True
            SetDelay m_name&"_grace_period", "MultiballsHandler" , Array(Array("grace_period", Me),Null), shoot_again_ms
        End If
        If hurry_up_time_ms > 0 Then
            m_hurry_up_enabled = True
            SetDelay m_name&"_hurry_up", "MultiballsHandler" , Array(Array("hurry_up", Me),Null), shoot_again_ms - hurry_up_time_ms
        End If
    End Sub

    Sub RunHurryUp()
        Log("Starting Hurry Up")
        m_hurry_up_enabled = False
        DispatchPinEvent m_name & "_hurry_up", Null
    End Sub

    Sub RunGracePeriod()
        Log("Starting Grace Period")
        m_grace_period_enabled = False
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Function BallDrainShootAgain(balls):
        Dim balls_to_save, kwargs

        If balls = 0 Then
            BallDrainShootAgain = balls
            Exit Function
        End If

        balls_to_save = m_balls_live_target - balls

        Log "Balls to save: " & balls_to_save & ". Balls live target: " & m_balls_live_target & ". Balls in Play: " & glf_BIP & ". Balls Drained: " & balls
        
        If balls_to_save <= 0 Then
            BallDrainShootAgain = balls
        End If

        If balls_to_save > balls Then
            balls_to_save = balls
        End If

        Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls", balls_to_save
        End With
        DispatchPinEvent m_name & "_shoot_again", kwargs
        
        Log("Ball drained during MB. Requesting a new one")
        m_queued_balls = m_queued_balls + 1
        SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me, m_queued_balls),Null), 1000

        BallDrainShootAgain = balls - balls_to_save
    End Function

    Function BallDrainCountBalls(balls):
        DispatchPinEvent m_name & "_ball_lost", Null
        If not glf_gameStarted or (glf_BIP - balls) = 1 Then
            m_balls_added_live = 0
            m_balls_live_target = 0
            DispatchPinEvent m_name & "_ended", Null
            RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
            Log("Ball drained. MB ended.")
        End If
        BallDrainCountBalls = balls
    End Function

    Public Sub Reset()
        Log "Resetting multiball: " & m_name
        DispatchPinEvent m_name & "_reset_event", Null

        Disable()
        m_shoot_again_enabled = False
        m_balls_added_live = 0
        m_balls_live_target = 0
    End Sub

    Public Sub AddABall()
        If m_balls_live_target > 0 Then
            Log "Adding a ball to multiball: " & m_name
            m_balls_live_target = m_balls_live_target + 1
            m_balls_added_live = m_balls_added_live + 1
            m_queued_balls = m_queued_balls + 1
            glf_BIP = glf_BIP + 1
            SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me, m_queued_balls),Null), 1000
        End If
    End Sub

    Public Sub AddAballTimerStart()
        'Start the timer for add a ball ball save.
        'This is started when multiball add a ball is triggered if configured,
        'and the default timer is not still running.
        If m_shoot_again_enabled = True Then
            Exit Sub
        End If

        m_shoot_again_enabled = True

        Dim shoot_again_ms : shoot_again_ms = m_add_a_ball_shoot_again.Value()
        if shoot_again_ms = 0 Then
            'No shoot again. Just stop multiball right away
            StopMultiball()
            Exit Sub
        End If

        DispatchPinEvent "ball_save_" & m_configname & "_add_a_ball_timer_start", Null

        Dim grace_period_ms : grace_period_ms = m_add_a_ball_grace_period.Value()
        Dim hurry_up_time_ms : hurry_up_time_ms = m_add_a_ball_hurry_up_time.Value()
        StartShootAgain shoot_again_ms, grace_period_ms, hurry_up_time_ms
    End Sub

    Public Sub StopMultiball()
        '"""Stop shoot again."""
        Log("Stopping shoot again of multiball")
        m_shoot_again_enabled = False

        '# disable shoot again
        RemoveDelay m_name&"_disable_shoot_again"

        If m_grace_period_enabled Then
            RemoveDelay m_name&"_grace_period"
            RunGracePeriod()
        End If
        If m_hurry_up_enabled Then
            RemoveDelay m_name&"_hurry_up"
            RunHurryUp()
        End If
        Log "Stop Shoot Again, Queued Balls: " & QueuedBalls()
        'RemoveDelay m_name & "_queued_release"

        DispatchPinEvent m_name & "_shoot_again_ended", Null
    End Sub

    Public Function QueuedBalls()
        QueuedBalls = m_queued_balls
    End Function

    Public Function ReleaseQueuedBalls()
        m_queued_balls = m_queued_balls - 1
        Log "Queued Balls: " & m_queued_balls
        ReleaseQueuedBalls = m_queued_balls
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function MultiballsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set MultiballsHandler = kwargs
                Else
                    MultiballsHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "start"
            multiball.Start
        Case "reset"
            multiball.Reset
        Case "add_a_ball"
            multiball.AddABall
        Case "stop"
            multiball.StopMultiball
        Case "grace_period"
            multiball.RunGracePeriod
        Case "hurry_up"
            multiball.RunHurryUp
        Case "drain"
            kwargs = multiball.BallsDrained(kwargs)
        Case "queue_release"
            If multiball.QueuedBalls() > 0 Then
                If glf_plunger.HasBall = False And ballInReleasePostion = True And glf_plunger.IncomingBalls = 0 Then
                    Glf_ReleaseBall(Null)
                    SetDelay multiball.Name&"_auto_launch", "MultiballsHandler" , Array(Array("auto_launch", multiball),Null), 500
                    If multiball.ReleaseQueuedBalls() > 0 Then
                        SetDelay multiball.Name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", multiball), Null), 1000    
                    End If
                Else
                    SetDelay multiball.Name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", multiball), Null), 1000
                End If
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay multiball.Name&"_auto_launch", "MultiballsHandler" , Array(Array("auto_launch", multiball), Null), 500
            End If
    End Select

    If IsObject(args(1)) Then
        Set MultiballsHandler = kwargs
    Else
        MultiballsHandler = kwargs
    End If
End Function
