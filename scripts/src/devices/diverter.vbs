Function CreateGlfDiverter(name)
	Dim diverter : Set diverter = (new GlfDiverter)(name)
	Set CreateGlfDiverter = diverter
End Function

Class GlfDiverter

    Private m_name
    Private m_activate_events
    Private m_deactivate_events
    Private m_activation_time
    Private m_enable_events
    Private m_disable_events
    Private m_activation_switches
    Private m_action_cb
    Private m_enabled
    Private m_active
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "active":
                GetValue = m_active
        End Select
    End Property

    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "DiverterEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "DiverterEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let ActivateEvents(value) 
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_activate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let DeactivateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_deactivate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ActivationTime(value) : Set m_activation_time = CreateGlfInput(value) : End Property
    Public Property Let ActivationSwitches(value) : m_activation_switches = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "diverter_" & name
        m_enable_events = Array()
        m_disable_events = Array()
        Set m_activate_events = CreateObject("Scripting.Dictionary")
        Set m_deactivate_events = CreateObject("Scripting.Dictionary")
        m_activation_switches = Array()
        Set m_activation_time = CreateGlfInput(0)
        m_debug = False
        m_enabled = False
        m_active = False
        glf_diverters.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_activate_events.Keys()
            AddPinEventListener m_activate_events(evt).EventName, m_name & "_" & evt & "_activate", "DiverterEventHandler", 1000, Array("activate", Me, m_activate_events(evt))
        Next
        For Each evt in m_deactivate_events.Keys()
            AddPinEventListener m_deactivate_events(evt).EventName, m_name & "_" & evt & "_deactivate", "DiverterEventHandler", 1000, Array("deactivate", Me, m_deactivate_events(evt))
        Next
        For Each evt in m_activation_switches
            AddPinEventListener evt & "_active", m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_activate_events.Keys()
            RemovePinEventListener m_activate_events(evt).EventName, m_name & "_" & evt & "_activate"
        Next
        For Each evt in m_deactivate_events.Keys()
            RemovePinEventListener m_deactivate_events(evt).EventName, m_name & "_" & evt & "_deactivate"
        Next
        For Each evt in m_activation_switches
            RemovePinEventListener evt & "_active", m_name & "_activate"
        Next
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
    End Sub

    Public Sub Activate()
        Log "Activating"
        m_active = True
        GetRef(m_action_cb)(1)
        If m_activation_time.Value > 0 Then
            SetDelay m_name & "_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_activation_time.Value
        End If
        DispatchPinEvent m_name & "_activating", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        m_active = False
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_deactivating", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DiverterEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim diverter : Set diverter = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set DiverterEventHandler = kwargs
                Else
                    DiverterEventHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "enable"
            diverter.Enable
        Case "disable"
            diverter.Disable
        Case "activate"
            diverter.Activate
        Case "deactivate"
            diverter.Deactivate
    End Select
    If IsObject(args(1)) Then
        Set DiverterEventHandler = kwargs
    Else
        DiverterEventHandler = kwargs
    End If
End Function