Class GlfStateMachine
    Private m_name
    Private m_player_var_name
    Private m_mode
    Private m_priority
    Private m_states
    Private m_transitions
    private m_base_device
 
    Private m_state
    Private m_persist_state
    Private m_starting_state
    Private m_show
 
    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "state":
                GetValue = m_state
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
 
    Public Property Get Transitions(name)
        If m_transitions.Exists(name) Then
            Set Transitions = m_transitions(name)
        Else
            Dim new_transition : Set new_transition = (new GlfStateMachineTranistion)(name)
            m_transitions.Add name, new_transition
            Set Transitions = new_transition
        End If
    End Property
 
    Public Property Get PersistState(): PersistState = m_persist_state : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property
 
    Public default Function init(name, mode)
        m_name = name
        m_player_var_name = "state_machine_" & name
        m_mode = mode.Name
        m_priority = mode.Priority

        m_persist_state = False
        m_starting_state = "start"
        m_state = Null
 
        Set m_states = CreateObject("Scripting.Dictionary")
        Set m_transitions = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "state_machine", Me)
 
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
        m_state = Null
        If Not IsEmpty(m_show) Then
            'm_show.Stop()
            m_show = Empty
        End If
    End Sub

    Public Sub StartState(start_state)
        m_base_device.Log("Starting state " & start_state)
        If Not m_states.Exists(start_state) Then
            m_base_device.Log("Invalid state " & start_state)
            Exit Sub
        End If

        Dim state_config : Set state_config = m_states(start_state)

        State() = start_state
        If UBound(state_config.EventsWhenStarted().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStarted().Items()
                If Not IsNull(evt.Condition) Then
                    If GetRef(evt.Condition)() = True Then
                        DispatchPinEvent evt.EventName, Null
                    End If
                Else
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        AddHandlersForCurrentState()
        RunShowForCurrentState()
    End Sub

    Public Sub StopCurrentState()
        m_base_device.Log "Stopping state " & State()
        RemoveHandlers()
        Dim state_config : Set state_config = m_states(state)

        If UBound(state_config.EventsWhenStopped().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStopped().Items()
                If Not IsNull(evt.Condition) Then
                    If GetRef(evt.Condition)() = True Then
                        DispatchPinEvent evt.EventName, Null
                    End If
                Else
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        If Not IsEmpty(m_show) Then
            m_base_device.Log "Stopping show " & m_show
            m_show.Stop()
            m_show = Empty
        End If

        State() = Null
    End Sub

    Public Sub RunShowForCurrentState()
        If IsNull(m_show) Then
            Exit Sub
        End If
        Dim state_config : Set state_config = m_states(state)
        If Not IsEmpty(state_config.ShowWhenActive) Then
            m_base_device.Log "Starting show %s" & state_config.ShowWhenActive
            'm_show = self.machine.show_controller.play_show_with_config(state_config['show_when_active'],
        End If
    End Sub

    Public Sub AddHandlersForCurrentState()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            If transition.Source.Exists(State()) Then
                For Each evt in transition.Events.Items()
                    AddPinEventListener evt.EventName, m_name & "_" & transition.Name & "_" & evt.EventName & "_transition", "StateMachineTransitionHandler", m_priority, Array("transition", Me, evt, transition)
                Next
            End If
        Next
    End Sub

    Public Sub RemoveHandlers()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            For Each evt in transition.Events.Items()
                RemovePinEventListener evt.EventName, m_name & "_" & transition.Name & "_" & evt.EventName & "_transition"
            Next
        Next
    End Sub

    Public Sub MakeTransition(transition)

        m_base_device.Log "Transitioning from " & State() & " to " & transition.Target
        StopCurrentState()
        If UBound(transition.EventsWhenTransitioning().Keys()) > -1 Then
            Dim evt
            For Each evt in transition.EventsWhenTransitioning().Items()
                If Not IsNull(evt.Condition) Then
                    If GetRef(evt.Condition)() = True Then
                        DispatchPinEvent evt.EventName, Null
                    End If
                Else
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If
        
        StartState transition.Target

    End Sub

End Class
 
Class GlfStateMachineState
	Private m_name, m_label, m_show_when_active, m_show_tokens, m_events_when_started, m_events_when_stopped
 
	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property
 
    Public Property Get Label(): Label = m_label: End Property
    Public Property Let Label(input): m_label = input: End Property
 
    Public Property Get ShowWhenActive(): ShowWhenActive = m_show_when_active: End Property
    Public Property Let ShowWhenActive(input): m_show_when_active = input: End Property
 
    Public Property Get ShowTokens(): Set ShowTokens = m_show_tokens: End Property
 
    Public Property Get EventsWhenStarted(): Set EventsWhenStarted = m_events_when_started: End Property
    Public Property Let EventsWhenStarted(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_started.Add newEvent.Name, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenStopped(): Set EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(input)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_stopped.Add newEvent.Name, newEvent
        Next
    End Property
 
	Public default Function init(name)
        m_name = name
        m_label = Empty
        m_show_when_active = Empty
        Set m_show_tokens = CreateObject("Scripting.Dictionary")
        Set m_events_when_started = CreateObject("Scripting.Dictionary")
        Set m_events_when_stopped = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class
 
Class GlfStateMachineTranistion
	Private m_name, m_sources, m_target, m_events, m_events_when_transitioning
 
	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property    

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
            m_events.Add newEvent.Name, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenTransitioning(): Set EventsWhenTransitioning = m_events_when_transitioning: End Property
    Public Property Let EventsWhenTransitioning(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_transitioning.Add newEvent.Name, newEvent
        Next    
    End Property
 
	Public default Function init(name)
        m_name = name
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
            If Not IsNull(glf_event.Condition) Then
                If GetRef(glf_event.Condition)() = True Then
                    state_machine.MakeTransition ownProps(3)
                End If
            Else
                state_machine.MakeTransition ownProps(3)
            End If
    End Select
    If IsObject(args(1)) Then
        Set StateMachineTransitionHandler = kwargs
    Else
        StateMachineTransitionHandler = kwargs
    End If
End Function