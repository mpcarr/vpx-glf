

Class GlfEventPlayer

    Private m_priority
    Private m_mode
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "event_player" : End Property
    Public Property Get Events() : Set Events = m_events : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Name, newEvent
        m_eventValues.Add newEvent.Name, value
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_event_player_play", "EventPlayerEventHandler", m_priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        If Not IsNull(m_events(evt).Condition) Then
            If GetRef(m_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            DispatchPinEvent evtValue, Null
        Next
    End Sub

    Public Function ToYaml
        Dim yaml
        Dim evt
        For Each evt In m_events.Keys()
            yaml = yaml & "  evt: " & Join(m_events(evt), ",") & vbCrLf
        Next
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