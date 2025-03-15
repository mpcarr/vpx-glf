Function EnableGlfBallSearch()
	Dim ball_search : Set ball_search = (new GlfBallSearch)()
    With CreateGlfMode("glf_ball_search", 100)
        .StartEvents = Array("reset_complete")

        With .TimedSwitches("flipper_cradle")
            .Switches = Array("s_left_flipper", "s_right_flipper")
            .Time = 3000
            .EventsWhenActive = Array("flipper_cradle")
            .EventsWhenReleased = Array("flipper_release")
            .Debug = True
        End With
    End With
	Set EnableGlfBallSearch = ball_search
End Function

Class GlfBallSearch

    Private m_debug
    Private m_timeout
    Private m_search_interval
    Private m_ball_search_wait_after_iteration
    Private m_phase
    Private m_devices
    Private m_current_device_type

    Public Property Get GetValue(value)
        'Select Case value
            'Case   
        'End Select
        GetValue = True
    End Property

    Public Property Let Timeout(value): Set m_timeout = CreateGlfInput(value): End Property
    Public Property Get Timeout(): Timeout = m_timeout.Value(): End Property

    Public Property Let SearchInterval(value): Set m_search_interval = CreateGlfInput(value): End Property
    Public Property Get SearchInterval(): SearchInterval = m_search_interval.Value(): End Property

    Public Property Let BallSearchWaitAfterIteration(value): Set m_ball_search_wait_after_iteration = CreateGlfInput(value): End Property
    Public Property Get BallSearchWaitAfterIteration(): BallSearchWaitAfterIteration = m_ball_search_wait_after_iteration.Value(): End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init()
        Set m_timeout = CreateGlfInput(15000)
        Set m_search_interval = CreateGlfInput(150)
        Set m_ball_search_wait_after_iteration = CreateGlfInput(10000)
        m_phase = 0
        m_devices = Array()
        m_current_device_type = Empty
        Set glf_ballsearch = Me
        SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
        AddPinEventListener "flipper_cradle", "ball_search_flipper_cradle", "BallSearchHandler", 30, Array("stop", Me)
        AddPinEventListener "flipper_release", "ball_search_flipper_cradle", "BallSearchHandler", 30, Array("reset", Me)
        Set Init = Me
    End Function

    Public Sub Start(phase)
        Dim ball_hold, held_balls
        held_balls = 0
        For Each ball_hold in glf_ball_holds.Items()
            held_balls = held_balls + ball_hold.GetValue("balls_held")
        Next
        If glf_gameStarted = True And glf_BIP > 0 And (glf_BIP-held_balls)>0 And glf_plunger.HasBall() = False Then
            m_phase = phase
            glf_last_switch_hit_time = 0
            'Fire all auto fire devices, slings, pops.
            m_devices = glf_autofiredevices.Items()
            m_current_device_type = "autofire"
            If UBound(m_devices) > -1 Then
                m_devices(0).BallSearch(m_phase)
                SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
            End If
        Else
            SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
        End If
    End Sub

    Public Sub NextDevice(device_index)
        If UBound(m_devices) > device_index Then
            m_devices(device_index+1).BallSearch(m_phase)
            SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, device_index+1), Null), m_search_interval.Value
        Else
            If m_current_device_type = "autofire" Then
                m_devices = glf_ball_devices.Items()
                m_current_device_type = "balldevices"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            ElseIf m_current_device_type = "balldevices" Then
                m_devices = glf_droptargets.Items()
                m_current_device_type = "droptargets"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            ElseIf m_current_device_type = "droptargets" Then
                m_devices = glf_diverters.Items()
                m_current_device_type = "diverters"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            Else
                m_current_device_type = Empty
                If m_phase < 3 Then
                    Start m_phase+1
                Else
                    m_phase = 0
                    SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
                End If
            End If
        End If
    End Sub

    Public Sub Reset()
        RemoveDelay "ball_search_next_device"
        m_phase = 0
        SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
    End Sub

    Public Sub StopBallSearch()
        RemoveDelay "ball_search_next_device"
        m_phase = 0
        RemoveDelay "ball_search"
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function BallSearchHandler(args)
    Dim ownProps, kwargs
    ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ball_search : Set ball_search = ownProps(1)
    Select Case evt
        Case "start"
            ball_search.Start 1
        Case "reset":
            ball_search.Reset
        Case "stop":
            ball_search.StopBallSearch
        Case "next_device"
            ball_search.NextDevice ownProps(2)
    End Select
End Function
