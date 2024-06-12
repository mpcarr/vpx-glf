
Class Multiball

    Private m_name
    Private m_priority
    Private m_mode
    Private m_multiball_locks
    Private m_enable_events
    Private m_disable_events
    Private m_start_events
    Private m_ball_save
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let StartEvents(value) : m_start_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, multiball_locks, mode)
        m_name = "multiball_" & name
        m_mode = mode.Name
        m_start_events = Array()
        m_multiball_locks = multiball_locks
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "MultiballHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "MultiballHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MultiballHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_starting", "MultiballHandler", m_priority, Array("start", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_start_events
            RemovePinEventListener evt, m_name & "_starting"
        Next
    End Sub

    Public Sub StartMultiball()
        glf_BIP = glf_BIP + GetPlayerState(m_multiball_locks & "_balls_locked")
        DispatchPinEvent m_name & "_started", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MultiballHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    Select Case evt
        Case "activate"
            multiball.Activate
        Case "deactivate"
            multiball.Deactivate
        Case "enable"
            multiball.Enable
        Case "disable"
            multiball.Disable
        Case "start"
            multiball.StartMultiball
    End Select
    MultiballHandler = kwargs
End Function