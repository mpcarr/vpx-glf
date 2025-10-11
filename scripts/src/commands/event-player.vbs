

Class GlfEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "event_player" : End Property

    Public Property Get Events() : Set Events = m_events : End Property
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
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Raw, newEvent
        
        Dim evtValue, evtValues(), i
        Redim evtValues(UBound(value))
        i=0
        For Each evtValue in value
            Dim newEventValue : Set newEventValue = (new GlfEventDispatch)(evtValue)
            Set evtValues(i) = newEventValue
            i=i+1
        Next
        m_eventValues.Add newEvent.Raw, evtValues  
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            Log "Adding Event Listener for: " & m_events(evt).EventName
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_event_player_play", "EventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        Log "Dispatching Event: " & evt
        If Not IsNull(m_events(evt).Condition) Then
            'msgbox m_events(evt).Condition
            If GetRef(m_events(evt).Condition)(Null) = False Then
                Exit Sub
            End If
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            Log "Dispatching Event: " & evtValue.EventName
            Glf_BcpSendEvent evtValue.EventName
            DispatchPinEvent evtValue.EventName, evtValue.Kwargs
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_event_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt,key
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & Replace(Replace(m_events(key).Raw, "&&", "and"), "||", "or") & ": " & vbCrLf
                For Each evt in m_eventValues(key)
                    yaml = yaml & "    - " & Replace(Replace(evt.Raw, "&&", "and"), "||", "or") & vbCrLf
                Next
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function EventPlayerEventHandler(args)
    
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
            eventPlayer.FireEvent ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set EventPlayerEventHandler = kwargs
    Else
        EventPlayerEventHandler = kwargs
    End If
End Function