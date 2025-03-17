Class GlfTimedSwitches

    Private m_name
    Private m_priority
    Private m_base_device
    Private m_time
    Private m_switches
    Private m_events_when_active
    Private m_events_when_released
    Private m_active_switches
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        'Select Case value
            'Case   
        'End Select
        GetValue = True
    End Property

    Public Property Let Time(value): Set m_time = CreateGlfInput(value): End Property
    Public Property Get Time(): Time = m_time.Value(): End Property
    
    Public Property Let Switches(value): m_switches = value: End Property
    
    Public Property Let EventsWhenActive(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_active.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenReleased(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_released.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(name, mode)
        m_name = "timed_switch_" & name
        m_priority = mode.Priority
        Set m_events_when_active = CreateObject("Scripting.Dictionary")
        Set m_events_when_released = CreateObject("Scripting.Dictionary")
        Set m_time = CreateGlfInput(0)
        m_switches = Array()
        Set m_active_switches = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "timed_switch", Me)
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim switch
        For Each switch in m_switches
            AddPinEventListener switch & "_active", m_name & "_active", "TimedSwitchHandler", m_priority, Array("active", Me, switch)
            AddPinEventListener switch & "_inactive", m_name & "_inactive", "TimedSwitchHandler", m_priority, Array("inactive", Me, switch)
        Next
    End Sub

    Public Sub Deactivate()
        Dim switch
        For Each switch in m_switches
            RemovePinEventListener switch & "_active", m_name & "_active"
            RemovePinEventListener switch & "_inactive", m_name & "_inactive"
        Next
    End Sub

    Public Sub SwitchActive(switch)
        If UBound(m_active_switches.Keys()) = -1 Then
            Dim evt
            For Each evt in m_events_when_active.Keys()
                Log "Switch Active: " & switch & ". Event: " & m_events_when_active(evt).EventName
                DispatchPinEvent m_events_when_active(evt).EventName, Null
            Next
        End If
        If Not m_active_switches.Exists(switch) Then
            m_active_switches.Add switch, True
        End If
    End Sub

    Public Sub SwitchInactive(switch)
        RemoveDelay m_name & "_" & switch & "_active"
        If m_active_switches.Exists(switch) Then
            m_active_switches.Remove switch
            If UBound(m_active_switches.Keys()) = -1 Then
                Dim evt
                For Each evt in m_events_when_released.Keys()
                    Log "Switch Release: " & switch & ". Event: " & m_events_when_released(evt).EventName
                    DispatchPinEvent m_events_when_released(evt).EventName, Null
                Next
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TimedSwitchHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim timed_switch : Set timed_switch = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set TimedSwitchHandler = kwargs
                Else
                    TimedSwitchHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "active"
            SetDelay timed_switch.Name & "_" & ownProps(2) & "_active" , "TimedSwitchHandler" , Array(Array("passed_time", timed_switch, ownProps(2)),Null), timed_switch.Time
        Case "passed_time"
            timed_switch.SwitchActive ownProps(2)
        Case "inactive"
            timed_switch.SwitchInactive ownProps(2)
    End Select

    If IsObject(args(1)) Then
        Set TimedSwitchHandler = kwargs
    Else
        TimedSwitchHandler = kwargs
    End If
End Function
