

Class GlfComboSwitches

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_switch_1
    Private m_switch_2
    Private m_events_when_both
    Private m_events_when_inactive
    Private m_events_when_one
    Private m_events_when_switch_1
    Private m_events_when_switch_2
    Private m_hold_time
    Private m_max_offset_time
    Private m_release_time

    Private m_switch_1_active
    Private m_switch_2_active

    Private m_switch_state

    Public Property Get Name() : Name = m_name : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get GetValue(value)
        Select Case value
            'Case ""
            '    GetValue = m_ticks
        End Select
    End Property

    Public Property Let Switch1(value): m_switch_1 = value: End Property
    Public Property Let Switch2(value): m_switch_2 = value: End Property
    Public Property Let HoldTime(value): Set m_hold_time = CreateGlfInput(value): End Property
    Public Property Let MaxOffsetTime(value): Set m_max_offset_time = CreateGlfInput(value): End Property
    Public Property Let ReleaseTime(value): Set m_release_time = CreateGlfInput(value): End Property
    Public Property Let EventsWhenBoth(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_both.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenInactive(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_inactive.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenOne(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_one.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenSwitch1(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_switch_1.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenSwitch2(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_switch_2.Add newEvent.Raw, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = "combo_switch_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
    
        m_switch_1 = Empty
        m_switch_2 = Empty
        Set m_events_when_both = CreateObject("Scripting.Dictionary")
        Set m_events_when_inactive = CreateObject("Scripting.Dictionary")
        Set m_events_when_one = CreateObject("Scripting.Dictionary")
        Set m_events_when_switch_1 = CreateObject("Scripting.Dictionary")
        Set m_events_when_switch_2 = CreateObject("Scripting.Dictionary")
        Set m_hold_time = CreateGlfInput(0)
        Set m_max_offset_time = CreateGlfInput(-1)
        Set m_release_time = CreateGlfInput(0)

        m_switch_1_active = 0
        m_switch_2_active = 0

        m_switch_state = Empty
        Set m_base_device = (new GlfBaseModeDevice)(mode, "combo_switch", Me)

        glf_combo_switches.Add name, Me

        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating Combo Switch"
        AddPinEventListener m_switch_1 & "_active" , m_name & "_switch1_active" , "ComboSwitchEventHandler", m_priority, Array("switch1_active", Me, m_switch_1)
        AddPinEventListener m_switch_1 & "_inactive" , m_name & "_switch1_inactive" , "ComboSwitchEventHandler", m_priority, Array("switch1_inactive", Me, m_switch_1)
        AddPinEventListener m_switch_2 & "_active" , m_name & "_switch2_active" , "ComboSwitchEventHandler", m_priority, Array("switch2_active", Me, m_switch_2)
        AddPinEventListener m_switch_2 & "_inactive" , m_name & "_switch2_inactive" , "ComboSwitchEventHandler", m_priority, Array("switch2_inactive", Me, m_switch_2)
    End Sub

    Public Sub Deactivate()
        Log "Deactivating Combo Switch"
        RemovePinEventListener m_switch_1 & "_active" , m_name & "_switch1_active"
        RemovePinEventListener m_switch_1 & "_inactive" , m_name & "_switch1_inactive"
        RemovePinEventListener m_switch_2 & "_active" , m_name & "_switch2_active"
        RemovePinEventListener m_switch_2 & "_inactive" , m_name & "_switch2_inactive"
        RemoveDelay m_name & "_" & "switch_1_inactive"
        RemoveDelay m_name & "_" & "switch_2_inactive"
        RemoveDelay m_name & "_" & "switch_1_active"
        RemoveDelay m_name & "_" & "switch_1_active"
        RemoveDelay m_name & "_" & "switch_2_only"
        RemoveDelay m_name & "_" & "switch_1_only"
    End Sub

    Public Sub Switch1WentActive(switch_name)
        Log "switch_1 just went active"
        RemoveDelay m_name & "_" & "switch_1_inactive"

        If m_switch_1_active > 0 Then
            Exit Sub
        End If

        If m_hold_time.Value() = 0 Then
            ActivateSwitches1 switch_name
        Else
            SetDelay m_name & "_switch_1_active", "ComboSwitchEventHandler", Array(Array("activate_switch1", Me, switch_name),Null), m_hold_time.Value()
        End If
    End Sub

    Public Sub Switch2WentActive(switch_name)
        Log "switch_2 just went active"
        RemoveDelay m_name & "_" & "switch_2_inactive"

        If m_switch_2_active > 0 Then
            Exit Sub
        End If

        If m_hold_time.Value() = 0 Then
            ActivateSwitches2 switch_name
        Else
            SetDelay m_name & "_switch_2_active", "ComboSwitchEventHandler", Array(Array("activate_switch2", Me, switch_name),Null), m_hold_time.Value()
        End If
    End Sub

    Public Sub Switch1WentInactive(switch_name)
        Log "switch_1 just went inactive"
        RemoveDelay m_name & "_" & "switch_1_active"

        If m_release_time.Value() = 0 Then
            ReleaseSwitch1 switch_name
        Else
            SetDelay m_name & "_switch_1_inactive", "ComboSwitchEventHandler", Array(Array("release_switch1", Me, switch_name),Null), m_release_time.Value()
        End If
    End Sub

    Public Sub Switch2WentInactive(switch_name)
        Log "switch_2 just went inactive"
        RemoveDelay m_name & "_" & "switch_2_active"

        If m_release_time.Value() = 0 Then
            ReleaseSwitch2 switch_name
        Else
            SetDelay m_name & "_switch_2_inactive", "ComboSwitchEventHandler", Array(Array("release_switch2", Me, switch_name),Null), m_release_time.Value()
        End If
    End Sub

    Public Sub ActivateSwitches1(switch_name)
        Log "Switch_1 has passed the hold time and is now active"
        m_switch_1_active = gametime
        RemoveDelay m_name & "_" & "switch_2_only"

        If m_switch_2_active > 0 Then
            If m_max_offset_time.Value() >= 0 And (m_switch_1_active - m_switch_2_active) > m_max_offset_time.Value() Then
                Log "Switches_2 is active, but the max_offset_time=" & m_max_offset_time.Value() & " which is larger than when a Switches_1 switch was first activated, so the state will not switch to both"
                Exit Sub
            End If
            PostSwitchStateEvents "both"
        ElseIf m_max_offset_time.Value()>=0 Then
            SetDelay m_name & "_switch_1_only", "ComboSwitchEventHandler", Array(Array("switch_1_only", Me, switch_name),Null), max_offset_time.Value()
        End If
    End Sub

    Public Sub ActivateSwitches2(switch_name)
        Log "Switch_2 has passed the hold time and is now active"
        m_switch_2_active = gametime
        RemoveDelay m_name & "_" & "switch_1_only"

        If m_switch_1_active > 0 Then
            If m_max_offset_time.Value() >= 0 And (m_switch_2_active - m_switch_1_active) > m_max_offset_time.Value() Then
                Log "Switches_1 is active, but the max_offset_time=" & m_max_offset_time.Value() & " which is larger than when a Switches_2 switch was first activated, so the state will not switch to both"
                Exit Sub
            End If
            PostSwitchStateEvents "both"
        ElseIf m_max_offset_time.Value()>=0 Then
            SetDelay m_name & "_switch_2_only", "ComboSwitchEventHandler", Array(Array("switch_2_only", Me, switch_name),Null), max_offset_time.Value()
        End If
    End Sub

    Public Sub ReleaseSwitch1(switch_name)
        Log "Switches_1 has passed the release time and is now released"
        m_switch_1_active = 0

        If m_switch_2_active > 0 And m_switch_state = "both" Then
            PostSwitchStateEvents "one"
        ElseIf m_switch_state = "one" Then
            PostSwitchStateEvents "inactive"
        End If
    End Sub

    Public Sub ReleaseSwitch2(switch_name)
        Log "Switches_2 has passed the release time and is now released"
        m_switch_2_active = 0

        If m_switch_1_active > 0 And m_switch_state = "both" Then
            PostSwitchStateEvents "one"
        ElseIf m_switch_state = "one" Then
            PostSwitchStateEvents "inactive"
        End If
    End Sub

    Public Sub PostSwitchStateEvents(state)
        If m_switch_state = state Then
            Exit Sub
        End If
        m_switch_state = state
        Log "New State " & state

        Dim evt
        Select Case state
            Case "both"
                For Each evt in m_events_when_both.Keys
                    If m_events_when_both(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_both(evt).EventName, Null
                    End If
                Next
            Case "one"
                For Each evt in m_events_when_one.Keys
                    If m_events_when_one(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_one(evt).EventName, Null
                    End If
                Next
            Case "inactive"
                For Each evt in m_events_when_inactive.Keys
                    If m_events_when_inactive(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_inactive(evt).EventName, Null
                    End If
                Next
            Case "switches_1"
                For Each evt in m_events_when_switch_1.Keys
                    If m_events_when_switch_1(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_switch_1(evt).EventName, Null
                    End If
                Next
            Case "switches_2"
                For Each evt in m_events_when_switch_2.Keys
                    If m_events_when_switch_2(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_switch_2(evt).EventName, Null
                    End If
                Next
        End Select

    End Sub

    Public Function ToYaml()
        Dim yaml, key, x
        yaml = "  " & Replace(m_name, "combo_switch_", "") & ":" & vbCrLf
        If Not IsEmpty(m_switch_1) Then
            yaml = yaml & "    switches_1: " & m_switch_1 & vbCrLf
        End If
        If Not IsEmpty(m_switch_2) Then
            yaml = yaml & "    switches_2: " & m_switch_2 & vbCrLf
        End If
        If UBound(m_events_when_both.Keys()) > -1 Then
            yaml = yaml & "    events_when_both: " & vbCrLf
            For Each key in m_events_when_both.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_inactive.Keys()) > -1 Then
            yaml = yaml & "    events_when_inactive: " & vbCrLf
            For Each key in m_events_when_inactive.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_switch_1.Keys()) > -1 Then
            yaml = yaml & "    events_when_switches_1: " & vbCrLf
            For Each key in m_events_when_switch_1.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_switch_2.Keys()) > -1 Then
            yaml = yaml & "    events_when_switches_2: " & vbCrLf
            For Each key in m_events_when_switch_2.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If UBound(m_events_when_switch_2.Keys()) > -1 Then
            yaml = yaml & "    events_when_switches_2: " & vbCrLf
            For Each key in m_events_when_switch_2.Keys()
                yaml = yaml & "      - " & key & vbCrLf
            Next
        End If
        If m_hold_time.Raw <> 0 Then
            yaml = yaml & "    hold_time: " & m_hold_time.Raw & vbCrLf
        End If
        If m_max_offset_time.Raw <> 0 Then
            yaml = yaml & "    max_offset_time: " & m_max_offset_time.Raw & vbCrLf
        End If
        If m_release_time.Raw <> 0 Then
            yaml = yaml & "    release_time: " & m_release_time.Raw & vbCrLf
        End If

        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function ComboSwitchEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim combo_switch : Set combo_switch = ownProps(1)
    Dim switch_name : switch_name = ownProps(2)
    Select Case evt
        Case "switch1_active"
            combo_switch.Switch1WentActive switch_name
        Case "switch2_active"
            combo_switch.Switch2WentActive switch_name
        Case "switch1_inactive"
            combo_switch.Switch1WentInactive switch_name
        Case "switch2_inactive"
            combo_switch.Switch2WentInactive switch_name
        Case "activate_switch1"
            combo_switch.ActivateSwitches1 switch_name
        Case "activate_switch2"
            combo_switch.ActivateSwitches2 switch_name
        Case "release_switch1"
            combo_switch.ReleaseSwitch1 switch_name
        Case "release_switch2"
            combo_switch.ReleaseSwitch2 switch_name
        Case "switch_1_only"
            combo_switch.PostSwitchStateEvents "switches_1"
        Case "switch_2_only"
            combo_switch.PostSwitchStateEvents "switches_2"
    End Select
    If IsObject(args(1)) Then
        Set ComboSwitchEventHandler = kwargs
    Else
        ComboSwitchEventHandler = kwargs
    End If
End Function
