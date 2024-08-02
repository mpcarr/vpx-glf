

Class GlfTimer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device

    Private m_control_events
    Private m_running
    Private m_ticks
    Private m_ticks_remaining
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_tick_interval
    Private m_starting_tick_interval
    Private m_max_value
    Private restart_on_complete
    Private m_start_running

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get ControlEvents(name)
        If m_control_events.Exists(name) Then
            Set ControlEvents = m_control_events(name)
        Else
            Dim newEvent : Set newEvent = (new GlfTimerControlEvent)()
            m_control_events.Add name, newEvent
            Set ControlEvents = newEvent
        End If
    End Property
    Public Property Get StartValue() : StartValue = m_start_value : End Property
    Public Property Get EndValue() : EndValue = m_end_value : End Property
    Public Property Get Direction() : Direction = m_direction : End Property

    Public Property Let StartValue(value) : m_start_value = value : End Property
    Public Property Let EndValue(value) : m_end_value = value : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let MaxValue(value) : m_max_value = value : End Property
    Public Property Let RestartOnComplete(value) : restart_on_complete = value : End Property
    Public Property Let StartRunning(value) : m_start_running = value : End Property
    Public Property Let TickInterval(value)
        m_tick_interval = value * 1000
        m_starting_tick_interval = value
    End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_direction = "up"
        m_ticks = 0
        m_ticks_remaining = 0
        m_tick_interval = 1000
        m_starting_tick_interval = 1
        restart_on_complete = False
        start_running = False

        Set m_control_events = CreateObject("Scripting.Dictionary")
        m_running = False

        Set m_base_device = (new GlfBaseModeDevice)(mode, "timer", Me)

        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_control_events.Keys
            AddPinEventListener m_control_events(evt).EventName, m_name & "_action", "TimerEventHandler", m_priority, Array("action", Me, m_control_events(evt))
        Next
        m_ticks = m_start_value
        m_ticks_remaining = m_ticks
        If m_start_running Then
            StartTimer()
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_control_events.Keys
            RemovePinEventListener m_control_events(evt).EventName, m_name & "_action"
        Next
        RemoveDelay m_name & "_tick"
        m_running = False
    End Sub

    Public Sub Action(controlEvent)

        dim value : value = controlEvent.Value
        Select Case controlEvent.Action
            Case "add"
                Add GetRef(value(0))()
            Case "subtract"
                Subtract GetRef(value(0))()
            Case "jump"
                Jump GetRef(value(0))()
            Case "start"
                StartTimer()
            Case "stop"
                StopTimer()
            Case "reset"
                Reset()
            Case "restart"
                Restart()
            Case "pause"
                Pause GetRef(value(0))()
            Case "set_tick_interval"
                SetTickInterval GetRef(value(0))()
            Case "change_tick_interval"
                ChangeTickInterval GetRef(value(0))()
            Case "reset_tick_interval"
                SetTickInterval m_starting_tick_interval
        End Select

    End Sub

    Private Sub StartTimer()
        If m_running Then
            Exit Sub
        End If

        m_base_device.Log "Starting Timer"
        m_running = True
        RemoveDelay m_name & "_unpause"
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_started", kwargs
        PostTickEvents()
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval
    End Sub

    Private Sub StopTimer()
        m_base_device.Log "Stopping Timer"
        m_running = False
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_stopped", kwargs
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub Pause(pause_ms)
        m_base_device.Log "Pausing Timer for "&pause_ms&" ms"
        m_running = False
        RemoveDelay m_name & "_tick"
        
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_paused", kwargs

        If pause_ms > 0 Then
            Dim startControlEvent : Set startControlEvent = (new GlfTimerControlEvent)()
            startControlEvent.Action = "start"
            SetDelay m_name & "_unpause", "TimerEventHandler", Array(Array("action", Me, startControlEvent), Null), pause_ms
        End If
    End Sub 

    Public Sub Tick()
        m_base_device.Log "Timer Tick"
        If Not m_running Then
            m_base_device.Log "Timer is not running. Will remove."
            Exit Sub
        End If

        Dim newValue
        If m_direction = "down" Then
            newValue = m_ticks - 1
        Else
            newValue = m_ticks + 1
        End If
        
        m_base_device.Log "ticking: old value: "& m_ticks & ", new Value: " & newValue & ", target: "& m_end_value
        m_ticks = newValue
        If Not PostTickEvents() Then
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval    
        End If
    End Sub

    Private Function CheckForDone

        ' Checks to see if this timer is done. Automatically called anytime the
        ' timer's value changes.
        m_base_device.Log "Checking to see if timer is done. Ticks: "&m_ticks&", End Value: "&m_end_value&", Direction: "& m_direction

        if m_direction = "up" And Not IsEmpty(m_end_value) And m_ticks >= m_end_value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If m_direction = "down" And m_ticks <= m_end_value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If Not IsEmpty(m_end_value) Then 
            m_ticks_remaining = abs(m_end_value - m_ticks)
        End If
        m_base_device.Log "Timer is not done"

        CheckForDone = False

    End Function

    Private Sub TimerComplete

        m_base_device.Log "Timer Complete"

        StopTimer()
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        
        If m_restart_on_complete Then
            m_base_device.Log "Restart on complete: True"
            Restart()
        End If
    End Sub

    Private Sub Restart
        Reset()
        If Not m_running Then
            StartTimer()
        Else
            PostTickEvents()
        End If
    End Sub

    Private Sub Reset
        m_base_device.Log "Resetting timer. New value: "& m_start_value
        Jump m_start_value
    End Sub

    Private Sub Jump(timer_value)
        m_ticks = (timer_value/1000)

        If m_max_value and m_ticks > m_max_value Then
            m_ticks = m_max_value
        End If

        CheckForDone()
    End Sub

    Public Sub ChangeTickInterval(change)
        m_tick_interval = m_tick_interval * (change/1000)
    End Sub

    Public Sub SetTickInterval(timer_value)
        m_tick_interval = timer_value
    End Sub
        
    Private Function PostTickEvents()
        PostTickEvents = True
        If Not CheckForDone() Then
            PostTickEvents = False
            Dim kwargs : Set kwargs = GlfKwargs()
            With kwargs
                .Add "ticks", m_ticks
                .Add "ticks_remaining", m_ticks_remaining
            End With
            DispatchPinEvent "m_name" & "_tick", kwargs
            m_base_device.Log "Ticks: "&m_ticks&", Remaining: " & m_ticks_remaining
        End If
    End Function

    Public Sub Add(timer_value) 
        Dim new_value

        new_value = m_ticks + (timer_value/1000)

        If Not IsEmpty(m_max_value) And new_value > m_max_value Then
            new_value = m_max_value
        End If
        m_ticks = new_value
        timer_value = new_value - timer_value

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_added", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_added", kwargs
        CheckForDone()
    End Sub

    Public Sub Subtract(timer_value)
        m_ticks = m_ticks - (timer_value/1000)
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_subtracted", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_subtracted", kwargs
        
        CheckForDone()
    End Sub
End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    
    Select Case evt
        Case "action"
            Dim controlEvent : Set controlEvent = ownProps(2)
            timer.Action controlEvent
        Case "tick"
            timer.Tick 
    End Select
    TimerEventHandler = kwargs
End Function

Class GlfTimerControlEvent
	Private m_event, m_action, m_value
  
	Public Property Get EventName(): EventName = m_event: End Property
    Public Property Let EventName(input): m_event = input: End Property

    Public Property Get Action(): Action = m_action : End Property
    Public Property Let Action(input): m_action = input : End Property

    Public Property Get Value()
        Value = m_value
    End Property
    Public Property Let Value(input)
        m_value = Glf_ParseInput(input, True)
    End Property

	Public default Function init()
        m_event = Empty
        m_action = Empty
        m_value = Empty
	    Set Init = Me
	End Function

End Class