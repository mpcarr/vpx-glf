Function CreateGlfBallDevice(name)
	Dim device : Set device = (new GlfBallDevice)(name)
	Set CreateGlfBallDevice = device
End Function

Class GlfBallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeout
    Private m_balls
    Private m_balls_in_device
    Private m_eject_angle
    Private m_eject_pitch
    Private m_eject_strength
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_balls_to_eject
    Private m_ejecting_all
    Private m_ejecting
    Private m_mechanical_eject
    Private m_eject_targets
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Let DefaultDevice(value)
        m_default_device = value
        If m_default_device = True Then
            Set glf_plunger = Me
        End If
    End Property
	Public Property Get HasBall()
        HasBall = (Not IsNull(m_balls(0)) And m_ejecting = False)
    End Property
    
    Public Property Get Balls(): Balls = m_balls_in_device : End Property

    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    
    Public Property Let EjectAngle(value) : m_eject_angle = glf_PI * value / 180 : End Property
    Public Property Let EjectPitch(value) : m_eject_pitch = glf_PI * value / 180 : End Property
    Public Property Let EjectStrength(value) : m_eject_strength = value : End Property
    
    Public Property Let EjectTimeout(value) : m_eject_timeout = value * 1000 : End Property
    Public Property Let EjectAllEvents(value)
        m_eject_all_events = value
        Dim evt
        For Each evt in m_eject_all_events
            AddPinEventListener evt, m_name & "_eject_all", "BallDeviceEventHandler", 1000, Array("ball_eject_all", Me)
        Next
    End Property
    Public Property Let EjectTargets(value)
        m_eject_targets = value
        Dim evt
        For Each evt in m_eject_targets
            AddPinEventListener evt & "_active", m_name & "_eject_target", "BallDeviceEventHandler", 1000, Array("eject_timeout", Me)
        Next
    End Property
    Public Property Let PlayerControlledEjectEvent(value)
        m_player_controlled_eject_event = value
        AddPinEventListener m_player_controlled_eject_event, m_name & "_eject_attempt", "BallDeviceEventHandler", 1000, Array("ball_eject", Me)
    End Property
    Public Property Let BallSwitches(value)
        m_ball_switches = value
        m_balls = Array(Ubound(m_ball_switches))
        Dim x
        For x=0 to UBound(m_ball_switches)
            m_balls(x) = Null
            AddPinEventListener m_ball_switches(x)&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_entering", Me, x)
            AddPinEventListener m_ball_switches(x)&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, x)
        Next
    End Property
    Public Property Let MechanicalEject(value) : m_mechanical_eject = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property
        
	Public default Function init(name)
        m_name = "balldevice_" & name
        m_ball_switches = Array()
        m_eject_all_events = Array()
        m_eject_targets = Array()
        m_balls = Array()
        m_debug = False
        m_default_device = False
        m_eject_pitch = 0
        m_eject_angle = 0
        m_eject_strength = 0
        m_ejecting = False
        m_eject_callback = Null
        m_ejecting_all = False
        m_balls_to_eject = 0
        m_balls_in_device = 0
        m_mechanical_eject = False
        m_eject_timeout = 1000
        glf_ball_devices.Add name, Me
	    Set Init = Me
	End Function

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_eject_timeout"
        'SoundSaucerLockAtBall ball
        Set m_balls(switch) = ball
        m_balls_in_device = m_balls_in_device + 1
        Log "Ball Entered" 
        Dim unclaimed_balls
        unclaimed_balls = DispatchRelayPinEvent(m_name & "_ball_enter", 1)
        Log "Unclaimed Balls: " & unclaimed_balls
        If (m_default_device = False Or m_ejecting = True) And unclaimed_balls > 0 And Not IsNull(m_balls(0)) Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        m_balls(switch) = Null
        m_balls_in_device = m_balls_in_device - 1
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechanical_eject = True And m_eject_timeout > 0 Then
            SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), ball), m_eject_timeout
        End If
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        m_ejecting = False
        RemoveDelay m_name & "_eject_timeout"
        DispatchPinEvent m_name & "_ball_eject_success", Null
        Log "Ball successfully exited"
        If m_ejecting_all = True Then
            If m_balls_to_eject = 0 Then
                m_ejecting_all = False
                Exit Sub
            End If
            If Not IsNull(m_balls(0)) Then
                m_balls_to_eject = m_balls_to_eject - 1
                Eject()
            Else
                SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 600
            End If
        End If
    End Sub

    Public Sub Eject
        Log "Ejecting."
        SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_balls(0)), m_eject_timeout
        m_ejecting = True
        If m_eject_strength > 0 Then
            If Not IsNull(m_balls(0)) Then
                m_balls(0).VelX = m_eject_strength * Cos(m_eject_pitch) * Sin(m_eject_angle)
                m_balls(0).VelY = m_eject_strength * Cos(m_eject_pitch) * Cos(m_eject_angle) * (-1)
                m_balls(0).VelZ = m_eject_strength * Sin(m_eject_pitch)
                Log "VelX: " &  m_balls(0).VelX & ", VelY: " &  m_balls(0).VelY & ", VelZ: " &  m_balls(0).VelZ
            End If
        End If

        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(m_balls(0))
        End If
    End Sub

    Public Sub EjectBalls(balls)
        Log "Ejecting "&balls&" Balls."
        m_ejecting_all = True
        m_balls_to_eject = balls - 1
        Eject()
    End Sub

    Public Sub EjectAll
        Log "Ejecting All." 
        m_ejecting_all = True
        m_balls_to_eject = m_balls_in_device - 1 
        Eject()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallDeviceEventHandler(args)
    Dim ownProps, ball : ownProps = args(0) : Set ball = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballDevice : Set ballDevice = ownProps(1)
    Dim switch
    Select Case evt
        Case "ball_entering"
            switch = ownProps(2)
            SetDelay ballDevice.Name & "_" & switch & "_ball_enter", "BallDeviceEventHandler", Array(Array("ball_enter", ballDevice, switch), ball), 500
        Case "ball_enter"
            switch = ownProps(2)
            ballDevice.BallEnter ball, switch
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_eject_all"
            ballDevice.EjectAll
        Case "ball_exiting"
            switch = ownProps(2)
            If RemoveDelay(ballDevice.Name & "_" & switch & "_ball_enter") = False Then
                ballDevice.BallExiting ball, switch
            End If
        Case "eject_timeout"
            ballDevice.BallExitSuccess ball
    End Select
End Function