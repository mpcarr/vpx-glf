

Class GlfTimer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_start_events
    Private m_stop_events
    Private m_debug

    Private m_value

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get StartValue() : StartValue = m_start_value : End Property
    Public Property Get EndValue() : EndValue = m_end_value : End Property
    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Get StartEvents() : StartEvents = m_start_events : End Property
    Public Property Get StopEvents() : StopEvents = m_stop_events : End Property
    
    Public Property Let StartValue(value) : m_start_value = value : End Property
    Public Property Let EndValue(value) : m_end_value = value : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "TimerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "TimerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_start_events.Keys
            AddPinEventListener m_start_events(evt).EventName, m_name & "_start", "TimerEventHandler", m_priority, Array("start", Me, evt)
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events.Keys
                AddPinEventListener m_stop_events(evt).EventName, m_name & "_stop", "TimerEventHandler", m_priority, Array("stop", Me, evt)
            Next
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_start_events.Keys
            RemovePinEventListener m_start_events(evt).EventName, m_name & "_start"
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events.Keys
                RemovePinEventListener m_stop_events(evt).EventName, m_name & "_stop"
            Next
        End If
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub StartTimer(evt)
        If Not IsNull(m_start_events(evt).Condition) Then
            If GetRef(m_start_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        Log "Started"
        DispatchPinEvent m_name & "_started", Null
        m_value = m_start_value
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
    End Sub

    Public Sub StopTimer(evt)
        If Not IsNull(m_stop_events(evt).Condition) Then
            If GetRef(m_stop_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        Log "Stopped"
        DispatchPinEvent m_name & "_stopped", Null
        RemoveDelay m_name & "_tick"
        m_value = m_start_value
    End Sub

    Public Sub Tick()
        Dim newValue
        If m_direction = "down" Then
            newValue = m_value - 1
        Else
            newValue = m_value + 1
        End If
        Log "ticking: old value: "& m_value & ", new Value: " & newValue & ", target: "& m_end_value
        m_value = newValue
        If m_value = m_end_value Then
            DispatchPinEvent m_name & "_complete", Null
        Else
            DispatchPinEvent m_name & "_tick", Null
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    Select Case evt
        Case "activate"
            timer.Activate
        Case "deactivate"
            timer.Deactivate
        Case "start"
            timer.StartTimer ownProps(2)
        Case "stop"
            timer.StopTimer ownProps(2)
        Case "tick"
            timer.Tick 
    End Select
    TimerEventHandler = kwargs
End Function