Class GlfStateMachine
    Private m_name
    Private m_player_var_name
    Private m_mode
    Private m_debug
    Private m_priority
    Private m_states
    Private m_transitions
    private m_base_device
 
    Private m_state
    Private m_persist_state
    Private m_starting_state
 
    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
        
    Public Property Get GetValue(value)
        Select Case value
            Case "state":
                GetValue = State()
        End Select
    End Property

    Public Property Get State()
        If m_persist_state = True Then
            Dim s : s = GetPlayerState(m_player_var_name)
            If s=False Then
                State = Null
            Else
                State = s
            End If
        Else
            State = m_state
        End If
    End Property

    Public Property Let State(value)
        If m_persist_state = True Then
            SetPlayerState m_player_var_name, value
            m_state = value
        Else
            m_state = value
        End If
    End Property
    
    Public Property Get States(name)
        If m_states.Exists(name) Then
            Set States = m_states(name)
        Else
            Dim new_state : Set new_state = (new GlfStateMachineState)(name)
            m_states.Add name, new_state
            Set States = new_state
        End If
    End Property
    Public Property Get StateItems(): StateItems = m_states.Items(): End Property
 
    Public Property Get Transitions()
        Dim count : count = UBound(m_transitions.Keys)
        Dim new_transition : Set new_transition = (new GlfStateMachineTranistion)()
        m_transitions.Add CStr(count), new_transition
        Set Transitions = new_transition
    End Property
 
    Public Property Get PersistState(): PersistState = m_persist_state : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property

    Public Property Get StartingState(): StartingState = m_starting_state : End Property
    Public Property Let StartingState(value) : m_starting_state = value : End Property
 
    Public default Function init(name, mode)
        m_name = name
        m_player_var_name = "state_machine_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        m_persist_state = False
        m_starting_state = "start"
        m_state = Null
 
        Set m_states = CreateObject("Scripting.Dictionary")
        Set m_transitions = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "state_machine", Me)
        glf_state_machines.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        If IsNull(State()) Then
            StartState m_starting_state
        Else
            AddHandlersForCurrentState()
            RunShowForCurrentState()
        End If

    End Sub

    Public Sub Disable()
        RemoveHandlers()
        StopShowForCurrentState()
        m_state = Null
    End Sub

    Public Sub StartState(start_state)
        Log("Starting state " & start_state)
        If Not m_states.Exists(start_state) Then
            Log("Invalid state " & start_state)
            Exit Sub
        End If

        Dim state_config : Set state_config = m_states(start_state)

        State() = start_state
        If UBound(state_config.EventsWhenStarted().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStarted().Items()
                If evt.Evaluate() = True Then
                    Glf_BcpSendEvent evt.EventName
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        AddHandlersForCurrentState()
        RunShowForCurrentState()
    End Sub

    Public Sub StopCurrentState()
        Log "Stopping state " & State()
        RemoveHandlers()
        If IsNull(state) Then
            Exit Sub
        End If
        Dim state_config : Set state_config = m_states(state)

        If UBound(state_config.EventsWhenStopped().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStopped().Items()
                If evt.Evaluate() = True Then
                    Glf_BcpSendEvent evt.EventName
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        StopShowForCurrentState()

        State() = Null
    End Sub

    Public Sub RunShowForCurrentState()
        If IsNull(state) Then
            Exit Sub
        End If
        Log state
        Dim state_config : Set state_config = m_states(state)
        If Not IsNull(state_config.ShowWhenActive().Show) Then
            Dim show : Set show = state_config.ShowWhenActive
            Log "Starting show %s" & m_name & "_" & show.Key
            Dim new_running_show
            Set new_running_show = (new GlfRunningShow)(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key, show.Key, show, m_priority, Null, state_config.InternalCacheId, m_mode)
        End If
    End Sub

    Public Sub StopShowForCurrentState()
        If IsNull(state) Then
            Exit Sub
        End If
        Dim state_config : Set state_config = m_states(state)
        If Not IsNull(state_config.ShowWhenActive().Show) Then
            Dim show : Set show = state_config.ShowWhenActive
            Log "Stopping show %s" & m_name & "_" & show.Key
            If glf_running_shows.Exists(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key) Then 
                glf_running_shows(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key).StopRunningShow()
            End If
        End If
    End Sub

    Public Sub AddHandlersForCurrentState()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            If transition.Source.Exists(State()) Then
                For Each evt in transition.Events.Items()
                    AddPinEventListener evt.EventName, m_name & "_" & transition.Target & "_" & evt.EventName & "_transition", "StateMachineTransitionHandler", m_priority+evt.Priority, Array("transition", Me, evt, transition)
                Next
            End If
        Next
    End Sub

    Public Sub RemoveHandlers()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            For Each evt in transition.Events.Items()
                RemovePinEventListener evt.EventName, m_name & "_" & transition.Target & "_" & evt.EventName & "_transition"
            Next
        Next
    End Sub

    Public Sub MakeTransition(transition)

        Log "Transitioning from " & State() & " to " & transition.Target
        StopCurrentState()
        If UBound(transition.EventsWhenTransitioning().Keys()) > -1 Then
            Dim evt
            For Each evt in transition.EventsWhenTransitioning().Items()
                If evt.Evaluate() = True Then
                    Glf_BcpSendEvent evt.EventName
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If
        
        StartState transition.Target

    End Sub

    Public Function ToYaml
        Dim yaml,x,key,y,cEvt
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml & "    starting_state: " & m_starting_state & vbCrLf
        If m_persist_state Then
            yaml = yaml & "    persist_state: " & m_persist_state & vbCrLf
        End If
        If UBound(m_states.Keys) > -1 Then
            yaml = yaml & "    states: " & vbCrLf
            x=0
            For Each key in m_states.keys
                yaml = yaml & "      " & m_states(key).Name & ": " & vbCrLf
                If Not IsEmpty(m_states(key).Label) Then
                    yaml = yaml & "        label: " & m_states(key).Label & vbCrLf
                End If
                
                If UBound(m_states(key).EventsWhenStarted.Keys) > -1 Then
                    yaml = yaml & "        events_when_started: "
                    y=0
                    For Each cEvt in m_states(key).EventsWhenStarted.Items
                        yaml = yaml & cEvt.Raw
                        If y <> UBound(m_states(key).EventsWhenStarted.Keys) Then
                            yaml = yaml & ", "
                        End If
                        y = y + 1
                    Next
                    yaml = yaml & vbCrLf
                End If

                If UBound(m_states(key).EventsWhenStopped.Keys) > -1 Then
                    yaml = yaml & "        events_when_stopped: "
                    y=0
                    For Each cEvt in m_states(key).EventsWhenStopped.Items
                        yaml = yaml & cEvt.Raw
                        If y <> UBound(m_states(key).EventsWhenStopped.Keys) Then
                            yaml = yaml & ", "
                        End If
                        y = y + 1
                    Next
                    yaml = yaml & vbCrLf
                End If
                
            Next
            yaml = yaml & "    transitions: " & vbCrLf
            For Each key in m_transitions.keys
                yaml = yaml & "      - source: "
                y=0
                For Each cEvt in m_transitions(key).Source.Keys
                    yaml = yaml & cEvt
                    If y <> UBound(m_transitions(key).Source.Keys) Then
                        yaml = yaml & ", "
                    End If
                    y = y + 1
                Next
                yaml = yaml & vbCrLf
                yaml = yaml & "        target: " & m_transitions(key).Target & vbCrLf
                yaml = yaml & "        events: " & vbCrLf
                
                If UBound(m_transitions(key).Events().Keys) > -1 Then
                    For Each cEvt in m_transitions(key).Events().keys()
                        yaml = yaml & "          - " & Replace(Replace(cEvt, "&&", "and"), "||", "or") & vbCrLf
                    Next
                End If
            Next
        End If

        ToYaml = yaml
    End Function

    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class
 
Class GlfStateMachineState
	Private m_name, m_label, m_show_when_active, m_events_when_started, m_events_when_stopped, m_internal_cache_id
 

    Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property

	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property
 
    Public Property Get Label(): Label = m_label: End Property
    Public Property Let Label(input): m_label = input: End Property
 
    Public Property Get ShowWhenActive()
        Set ShowWhenActive = m_show_when_active
    End Property

    Public Property Get EventsWhenStarted(): Set EventsWhenStarted = m_events_when_started: End Property
    Public Property Let EventsWhenStarted(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_started.Add newEvent.Raw, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenStopped(): Set EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(input)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_stopped.Add newEvent.Raw, newEvent
        Next
    End Property
 
	Public default Function init(name)
        m_name = name
        m_label = Empty
        m_internal_cache_id = -1
        Set m_show_when_active = (new GlfShowPlayerItem)()
        Set m_events_when_started = CreateObject("Scripting.Dictionary")
        Set m_events_when_stopped = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class
 
Class GlfStateMachineTranistion
	Private m_name, m_sources, m_target, m_events, m_events_when_transitioning

    Public Property Get Source(): Set Source = m_sources: End Property
    Public Property Let Source(value)
        Dim x
        For x=0 to UBound(value)
            m_sources.Add value(x), True
        Next    
    End Property
 
    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input): m_target = input: End Property  

    Public Property Get Events(): Set Events = m_events: End Property
    Public Property Let Events(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events.Add newEvent.Raw, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenTransitioning(): Set EventsWhenTransitioning = m_events_when_transitioning: End Property
    Public Property Let EventsWhenTransitioning(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_transitioning.Add newEvent.Raw, newEvent
        Next    
    End Property
 
	Public default Function init()
        Set m_sources = CreateObject("Scripting.Dictionary")
        m_target = Empty
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_events_when_transitioning = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class


Public Function StateMachineTransitionHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim state_machine : Set state_machine = ownProps(1)
    Select Case evt
        Case "transition"
            Dim glf_event : Set glf_event = ownProps(2)
            If glf_event.Evaluate() = True Then
                state_machine.MakeTransition ownProps(3)
            Else
                If glf_debug_level = "Debug" Then
                    glf_debugLog.WriteToLog "State machine transition",  "failed condition: " & glf_event.Raw
                End If
            End If
    End Select
    If IsObject(args(1)) Then
        Set StateMachineTransitionHandler = kwargs
    Else
        StateMachineTransitionHandler = kwargs
    End If
End Function