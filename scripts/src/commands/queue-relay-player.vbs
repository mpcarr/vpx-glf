

Class GlfQueueRelayPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "queue_relay_player" : End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "queue_relay_player", Me)
        Set Init = Me
	End Function

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get EventName(name)
        If m_events.Exists(name) Then
            Set EventName = m_eventValues(name)
        Else
            Dim new_event : Set new_event = (new GlfEvent)(name)
            m_events.Add new_event.Raw, new_event
            Dim new_event_value : Set new_event_value = (new GlfQueueRelayEvent)()
            m_eventValues.Add new_event.Raw, new_event_value
            Set EventName = new_event_value
        End If
    End Property

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_queue_relay_player_play", "QueueRelayPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_queue_relay_player_play"
        Next
    End Sub

    Public Function FireEvent(evt)
        FireEvent=Empty
        If m_events(evt).Evaluate() Then
            'post a new event, and wait for the release
            DispatchPinEvent m_eventValues(evt).Post, Null
            FireEvent = m_eventValues(evt).WaitFor
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_queue_relay_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & m_events(key).Raw & ": " & vbCrLf
                For Each evt in m_eventValues(key)
                    yaml = yaml & "    - " & evt & vbCrLf
                Next
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function QueueRelayPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            Dim wait_for : wait_for = eventPlayer.FireEvent(ownProps(2))
            If Not IsEmpty(wait_for) Then
                kwargs.Add "wait_for", wait_for
            End If
    End Select
    If IsObject(args(1)) Then
        Set QueueRelayPlayerEventHandler = kwargs
    Else
        QueueRelayPlayerEventHandler = kwargs
    End If
End Function

Class GlfQueueRelayEvent

	Private m_wait_for, m_post
  
    Public Property Get WaitFor() : WaitFor = m_wait_for : End Property
    Public Property Let WaitFor(input) : m_wait_for = input : End Property
    Public Property Get Post() : Post = m_post : End Property
    Public Property Let Post(input) : m_post = input : End Property
        
	Public default Function init()
        m_wait_for = Empty
        m_post = Empty
	    Set Init = Me
	End Function

End Class