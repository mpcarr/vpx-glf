

Class GlfEventPlayer

    Private m_priority
    Private m_mode
    private m_base_device
    Private m_events

    Public Property Get Events() : Set Events = m_events : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_event_player_play", "EventPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_event_player_play"
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
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            dim evtToFire
            For Each evtToFire in ownProps(2)
                DispatchPinEvent evtToFire, Null
            Next
    End Select
    EventPlayerEventHandler = kwargs
End Function