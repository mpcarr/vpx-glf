
Function CreateGlfAutoFireDevice(name)
	Dim flipper : Set flipper = (new GlfAutoFireDevice)(name)
	Set CreateGlfAutoFireDevice = flipper
End Function

Class GlfAutoFireDevice

    Private m_name
    Private m_enable_events
    Private m_disable_events
    Private m_enabled
    Private m_switch
    Private m_action_cb
    Private m_disabled_cb
    Private m_enabled_cb
    Private m_exclude_from_ball_search
    Private m_debug

    Public Property Let Switch(value)
        m_switch = value
    End Property
    Public Property Let ExcludeFromBallSearch(value) : m_exclude_from_ball_search = value : End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let DisabledCallback(value) : m_disabled_cb = value : End Property
    Public Property Let EnabledCallback(value) : m_enabled_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "AutoFireDeviceEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "AutoFireDeviceEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "auto_fire_coil_" & name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_enabled = False
        m_action_cb = Empty
        m_disabled_cb = Empty
        m_enabled_cb = Empty
        m_switch = Empty
        m_debug = False
        m_exclude_from_ball_search = False
        glf_autofiredevices.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        If Not IsEmpty(m_enabled_cb) Then
            GetRef(m_enabled_cb)()
        End If
        If Not IsEmpty(m_switch) Then
            AddPinEventListener m_switch & "_active", m_name & "_active", "AutoFireDeviceEventHandler", 1000, Array("activate", Me)
            AddPinEventListener m_switch & "_inactive", m_name & "_inactive", "AutoFireDeviceEventHandler", 1000, Array("deactivate", Me)
        End If
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        If Not IsEmpty(m_disabled_cb) Then
            GetRef(m_disabled_cb)()
        End If
        Deactivate(Null)
        RemovePinEventListener m_switch & "_active", m_name & "_active"
        RemovePinEventListener m_switch & "_inactive", m_name & "_inactive"
    End Sub

    Public Sub Activate(active_ball)
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(1, active_ball))
        End If
        DispatchPinEvent m_name & "_activate", Null
    End Sub

    Public Sub Deactivate(active_ball)
        Log "Deactivating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(0, active_ball))
        End If
        DispatchPinEvent m_name & "_deactivate", Null
    End Sub

    Public Sub BallSearch(phase)
        If m_exclude_from_ball_search = True Then
            Exit Sub
        End If
        Log "Ball Search, phase " & phase
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(1, Null))
        End If
        SetDelay m_name & "ball_search_deactivate", "AutoFireDeviceEventHandler", Array(Array("deactivate", Me), Null), 150
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function AutoFireDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim flipper : Set flipper = ownProps(1)
    Select Case evt
        Case "enable"
            flipper.Enable
        Case "disable"
            flipper.Disable
        Case "activate"
            flipper.Activate kwargs
        Case "deactivate"
            flipper.Deactivate kwargs
    End Select
    If IsObject(args(1)) Then
        Set AutoFireDeviceEventHandler = kwargs
    Else
        AutoFireDeviceEventHandler = kwargs
    End If
End Function