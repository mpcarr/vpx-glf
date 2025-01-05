

Class GlfRandomEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "random_event_player" : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get EventName(value)
        
        Dim newEvent : Set newEvent = (new GlfEvent)(value)
        m_events.Add newEvent.Raw, newEvent
        Dim newRandomEvent : Set newRandomEvent = (new GlfRandomEvent)(value, m_mode, UBound(m_events.Keys))
        m_eventValues.Add newEvent.Raw, newRandomEvent
        
        Set EventName = newRandomEvent
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "random_event_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_event_player_play", "RandomEventPlayerEventHandler", m_priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        If m_events(evt).Evaluate() Then
            Dim event_to_fire
            event_to_fire = m_eventValues(evt).GetNextRandomEvent()
            If Not IsEmpty(event_to_fire) Then
                Log "Dispatching Event: " & event_to_fire
                DispatchPinEvent event_to_fire, Null
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_random_event_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml : yaml = ""
        ToYaml = yaml
    End Function

End Class

Function RandomEventPlayerEventHandler(args)
    
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
        Set RandomEventPlayerEventHandler = kwargs
    Else
        RandomEventPlayerEventHandler = kwargs
    End If
End Function