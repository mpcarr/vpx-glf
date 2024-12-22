
Class GlfSequenceShots

    Private m_name
    Private m_command_name
    Private m_lock_device
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_cancel_events
    Private m_cancel_switches
    Private m_delay_event_list
    Private m_delay_switch_list
    Private m_event_sequence
    Private m_sequence_timeout
    Private m_switch_sequence
    Private m_start_event
    Private m_sequence_count
    Private m_active_delays
    Private m_active_sequences
    Private m_sequence_events
    Private m_start_time

    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    
    Public Property Get GetValue(value)
        Select Case value
            'Case "":
            '    GetValue = 
        End Select
    End Property

    Public Property Let CancelEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_cancel_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CancelSwitches(value): m_cancel_switches = value: End Property
    Public Property Let DelayEventList(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_delay_event_list.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let DelaySwitchList(value): m_delay_switch_list = value: End Property
    Public Property Let EventSequence(value)
        m_event_sequence = value
        If m_sequence_count = 0 Then
            m_sequence_events = value
        Else
            Redim Preserve m_sequence_events(m_sequence_count+UBound(value))
            Dim i
            For i = 0 To UBound(value)
                m_sequence_events(m_sequence_count + i) = value(i)
            Next
        End If
        m_start_event = value(0)
        m_sequence_count = m_sequence_count + (UBound(m_sequence_events)+1)
    End Property
    Public Property Let SequenceTimeout(value): Set m_sequence_timeout = CreateGlfInput(value): End Property
    Public Property Let SwitchSequence(value)
        m_switch_sequence = value
        If m_sequence_count = 0 Then
            m_start_event = value(0) & "_active"
        End If
        Redim Preserve m_sequence_events(m_sequence_count+UBound(value))
        Dim i
        For i = 0 To UBound(value)
            m_sequence_events(m_sequence_count + i) = value(i) & "_active"
        Next
        m_sequence_count = m_sequence_count + (UBound(m_sequence_events)+1)
    End Property

	Public default Function init(name, mode)
        m_name = "sequence_shot_" & name
        m_command_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        Set m_cancel_events = CreateObject("Scripting.Dictionary")
        Set m_delay_event_list = CreateObject("Scripting.Dictionary")
        Set m_active_sequences = CreateObject("Scripting.Dictionary")
        Set m_active_delays = CreateObject("Scripting.Dictionary")
        
        m_sequence_events = Array()
        m_cancel_switches = Array()
        m_start_time = 0
        m_event_sequence = Array()
        m_switch_sequence = Array()
        Set m_sequence_timeout = CreateGlfInput(0)
        m_sequence_count = 0
        m_start_event = Empty
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "sequence_shot_", Me)
        
        Set Init = Me
	End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_event_sequence
            AddPinEventListener evt, m_name & "_" & evt & "_advance", "SequenceShotsHandler", m_priority, Array("advance", Me, evt)
        Next
        For Each evt in m_switch_sequence
            AddPinEventListener evt & "_active", m_name & "_" & evt & "_advance", "SequenceShotsHandler", m_priority, Array("advance", Me, evt & "_active")
        Next
        For Each evt in m_cancel_events.Keys
            AddPinEventListener m_cancel_events(evt).EventName, m_name & "_" & evt & "_cancel", "SequenceShotsHandler", m_priority, Array("cancel_event", Me, m_cancel_events(evt))
        Next
        For Each evt in m_delay_event_list.Keys
            AddPinEventListener m_delay_event_list(evt).EventName, m_name & "_" & evt & "_delay", "SequenceShotsHandler", m_priority, Array("delay_event", Me, m_delay_event_list(evt))
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_event_sequence
            RemovePinEventListener evt, m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_switch_sequence
            RemovePinEventListener evt & "_active", m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_cancel_events.Keys
            RemovePinEventListener m_cancel_events(evt).EventName, m_name & "_" & evt & "_cancel"
        Next
        For Each evt in  m_delay_event_list.Keys
            RemovePinEventListener m_delay_event_list(evt).EventName, m_name & "_" & evt & "_delay"
        Next
    End Sub

    Sub SequenceAdvance(event_name)
        ' Since we can track multiple simultaneous sequences (e.g. two balls
        ' going into an orbit in a row), we first have to see whether this
        ' switch is starting a new sequence or continuing an existing one

        Log "Sequence advance: " & event_name

        If event_name = m_start_event Then
            If m_sequence_count > 1 Then
                ' start a new sequence
                StartNewSequence()
            ElseIf UBound(m_active_delays.Keys) = -1 Then
                ' if it only has one step it will finish right away
                Completed()
            End If
        Else
            ' Get the seq_id of the first sequence this switch is next for.
            ' This is not a loop because we only want to advance 1 sequence
            Dim k, seq
            seq = Null
            For Each k In m_active_sequences.Keys
                Log m_active_sequences(k).NextEvent
                If m_active_sequences(k).NextEvent = event_name Then
                    Set seq = m_active_sequences(k)
                    Exit For
                End If
            Next

            If Not IsNull(seq) Then
                ' advance this sequence
                AdvanceSequence(seq)
            End If
        End If
    End Sub

    Public Sub StartNewSequence()
        ' If the sequence hasn't started, make sure we're not within the
        ' delay_switch hit window

        If UBound(m_active_delays.Keys)>-1 Then
            Log "There's a delay timer in effect. Sequence will not be started."
            Exit Sub
        End If

        'record start time
        m_start_time = gametime

        ' create a new sequence
        Dim seq_id : seq_id = "seq_" & glf_SeqCount
        glf_SeqCount = glf_SeqCount + 1

        Dim next_event : next_event = m_sequence_events(1)

        Log "Setting up a new sequence. Next: " & next_event

        m_active_sequences.Add seq_id, (new GlfActiveSequence)(seq_id, 0, next_event)

        ' if this sequence has a time limit, set that up
        If m_sequence_timeout.Value > 0 Then
            Log "Setting up a sequence timer for " & m_sequence_timeout.Value
            SetDelay seq_id, "SequenceShotsHandler" , Array(Array("seq_timeout", Me, seq_id),Null), m_sequence_timeout.Value
        End If
    End Sub

    Public Sub AdvanceSequence(sequence)
        ' Remove this sequence from the list
        If sequence.CurrentPositionIndex = (m_sequence_count - 2) Then  ' complete
            Log "Sequence complete!"
            RemoveDelay sequence.SeqId
            m_active_sequences.Remove sequence.SeqId
            Completed()
        Else
            Dim current_position_index : current_position_index = sequence.CurrentPositionIndex + 1
            Dim next_event : next_event = m_sequence_events(current_position_index + 1)
            Log "Advancing the sequence. Next: " & next_event
            sequence.CurrentPositionIndex = current_position_index
            sequence.NextEvent = next_event
        End If
    End Sub

    Public Sub Completed()
        'measure the elapsed time between start and completion of the sequence
        Dim elapsed
        If m_start_time > 0 Then
            elapsed = gametime - m_start_time
        Else
            elapsed = 0
        End If

        'Post sequence complete event including its elapsed time to complete.
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "elapsed", elapsed
        End With
        DispatchPinEvent m_command_name & "_hit", kwargs
    End Sub

    Public Sub ResetAllSequences()
        'Reset all sequences."""
        Dim k
        For Each k in m_active_sequences.Keys
            RemoveDelay m_active_sequences(k).SeqId
        Next

        m_active_sequences.RemoveAll()
    End Sub

    Public Sub DelayEvent(delay, name)
        Log "Delaying sequence by " & delay
        SetDelay name & "_delay_timer", "SequenceShotsHandler" , Array(Array("release_delay", Me, name),Null), delay
        m_active_delays.Add name, True
    End Sub

    Public Sub ReleaseDelay(name)
        m_active_delays.Remove name
    End Sub

    Public Sub FireSequenceTimeout(seq_id)
        Log "Sequence " & seq_id & " timeouted"
        m_active_sequences.Remove seq_id
        DispatchPinEvent m_name & "_timeout", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Class GlfActiveSequence

    Private m_next_event, m_seq_id, m_idx

    Public Property Get SeqId(): SeqId = m_seq_id: End Property
    Public Property Get NextEvent(): NextEvent = m_next_event: End Property
    Public Property Let NextEvent(value): m_next_event = value: End Property
    Public Property Get CurrentPositionIndex(): CurrentPositionIndex = m_idx: End Property
    Public Property Let CurrentPositionIndex(value): m_idx = value: End Property

    Public default Function init(seq_id, idx, next_event)
        m_seq_id = seq_id
        m_idx = idx
        m_next_event = next_event
        Set Init = Me
    End Function

End Class

Function SequenceShotsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim sequence_shot : Set sequence_shot = ownProps(1)
    Select Case evt
        Case "advance"
            sequence_shot.SequenceAdvance ownProps(2)
        Case "cancel"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            sequence_shot.ResetAllSequences
        Case "delay"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            sequence_shot.DelayEvent glfEvent.Delay, glfEvent.EventName
        Case "seq_timeout"
            sequence_shot.FireSequenceTimeout ownProps(2)
        Case "release_delay"
            sequence_shot.ReleaseDelay ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set SequenceShotsHandler = kwargs
    Else
        SequenceShotsHandler = kwargs
    End If
End Function