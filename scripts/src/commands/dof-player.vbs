

Class GlfDofPlayer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "dof_player" : End Property


    Public Property Get EventDOF() : EventDOF = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_dof : Set new_dof = (new GlfDofPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_dof
        Set EventName = new_dof
        
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "dof_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "dof_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_dof_player_play", "DofPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_dof_player_play"
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            Log "Firing DOF Event: " & m_eventValues(evt).DOFEvent & " State: " & m_eventValues(evt).Action
            DOF m_eventValues(evt).DOFEvent, m_eventValues(evt).Action  
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_dof_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function DofPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim dof_player : Set dof_player = ownProps(1)
    Select Case evt
        Case "activate"
            dof_player.Activate
        Case "deactivate"
            dof_player.Deactivate
        Case "play"
            dof_player.Play(ownProps(2))
    End Select
    If IsObject(args(1)) Then
        Set DofPlayerEventHandler = kwargs
    Else
        DofPlayerEventHandler = kwargs
    End If
End Function

Class GlfDofPlayerItem
	Private m_dof_event, m_action
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        Select Case input
            Case "DOF_OFF"
                m_action = 0
            Case "DOF_ON"
                m_action = 1
            Case "DOF_PULSE"
                m_action = 2
        End Select
    End Property

    Public Property Get DOFEvent(): DOFEvent = m_dof_event: End Property
    Public Property Let DOFEvent(input): m_dof_event = CInt(input): End Property

	Public default Function init()
        m_action = Empty
        m_dof_event = Empty
        Set Init = Me
	End Function

End Class
