
Class GlfSegmentDisplayPlayer

    Private m_priority
    Private m_mode
    Private m_name
    Private m_debug
    Private m_events
    private m_base_device

    Public Property Get Name() : Name = "segment_player" : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    
    Public Property Get EventName(value)
        Dim newEvent : Set newEvent = (new GlfEvent)(value)

        If m_events.Exists(newEvent.Raw) Then
            Set EventName = m_events(newEvent.Raw)
        Else
            Dim new_segment_event : Set new_segment_event = (new GlfSegmentDisplayPlayerEvent)(newEvent)
            m_events.Add newEvent.Raw, new_segment_event
            Set EventName = new_segment_event
        End If
    End Property

	Public default Function init(mode)
        m_name = "segment_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "segment_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).GlfEvent.EventName, m_mode & "_segment_player_play", "SegmentPlayerEventHandler", m_priority+m_events(evt).GlfEvent.Priority, Array("play", Me, m_events(evt), m_events(evt).GlfEvent.EventName)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).GlfEvent.EventName, m_mode & "_segment_player_play"
            PlayOff m_events(evt).GlfEvent.EventName, m_events(evt)
        Next
    End Sub

    Public Sub Play(evt, segment_event)
        Dim i
        For i=0 to UBound(segment_event.Displays())
            SegmentPlayerCallbackHandler evt, segment_event.Displays()(i), m_mode, m_priority
        Next
    End Sub

    Public Sub PlayOff(evt, segment_event)
        Dim i, segment_item
        For i=0 to UBound(segment_event.Displays())
            Set segment_item = segment_event.Displays()(i)
            Dim key
            key = m_mode & "." & "segment_player_player." & segment_item.Display
            If Not IsEmpty(segment_item.Key) Then
                key = key & segment_item.Key
            End If
            Dim display : Set display = glf_segment_displays(segment_item.Display)
            RemoveDelay key
            display.RemoveTextByKey key    
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfSegmentDisplayPlayerEvent

    Private m_items
    Private m_event

    Public Property Get Displays() : Displays = m_items.Items() : End Property

    Public Property Get GlfEvent() : Set GlfEvent = m_event : End Property

    Public Property Get Display(value)
        If m_items.Exists(value) Then
            Set Display = m_items(value)
        Else
            Dim new_item : Set new_item = (new GlfSegmentPlayerEventItem)()
            new_item.Display = value
            m_items.add value, new_item
            Set Display = new_item
        End If
    End Property

    Public default Function init(evt)
        Set m_items = CreateObject("Scripting.Dictionary")
        Set m_event = evt
        Set Init = Me
	End Function

End Class

Class GlfSegmentPlayerEventItem
	
    private m_display
    private m_text
    private m_priority
    private m_action
    private m_expire
    private m_flash_mask
    private m_flashing
    private m_key
    private m_transition
    private m_transition_out
    private m_color

    Public Property Get Display() : Display = m_display : End Property
    Public Property Let Display(input) : m_display = input : End Property
    
    Public Property Get Text()
        If Not IsNull(m_text) Then
            Set Text = m_text
        Else
            Text = Null
        End If
    End Property
    Public Property Let Text(input) 
        Set m_text = (new GlfInput)(input)
    End Property

    Public Property Get Priority() : Priority = m_priority : End Property
    Public Property Let Priority(input) : m_priority = input : End Property
        
    Public Property Get Action() : Action = m_action : End Property
    Public Property Let Action(input) : m_action = input : End Property
                
    Public Property Get Expire() : Expire = m_expire : End Property
    Public Property Let Expire(input) : m_expire = input : End Property

    Public Property Get FlashMask() : FlashMask = m_flash_mask : End Property
    Public Property Let FlashMask(input) : m_flash_mask = input : End Property
                        
    Public Property Get Flashing() : Flashing = m_flashing : End Property
    Public Property Let Flashing(input) : m_flashing = input : End Property
                            
    Public Property Get Key() : Key = m_key : End Property
    Public Property Let Key(input) : m_key = input : End Property

    Public Property Get Color() : Color = m_color : End Property
    Public Property Let Color(input) : m_color = input : End Property

    Public Property Get HasTransition() : HasTransition = Not IsNull(m_transition) : End Property    
    Public Property Get HasTransitionOut() : HasTransitionOut = Not IsNull(m_transition_out) : End Property

    Public Property Get Transition()
        If IsNull(m_transition) Then
            Set m_transition = (new GlfSegmentPlayerTransition)()
            Set Transition = m_transition   
        Else
            Set Transition = m_transition
        End If
    End Property

    Public Property Get TransitionOut()
        If IsNull(m_transition_out) Then
            Set m_transition_out = (new GlfSegmentPlayerTransition)()
            Set TransitionOut = m_transition_out   
        Else
            Set TransitionOut = m_transition_out
        End If
    End Property
                                
	Public default Function init()
        m_display = Empty
        m_text = Null
        m_priority = 0
        m_action = "add"
        m_expire = 0
        m_flash_mask = Empty
        m_flashing = "no_flash"
        m_key = Empty
        m_transition = Null
        m_transition_out = Null
        m_color = Rgb(255,255,255)
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        If Not IsEmpty(m_display) Then
            yaml = yaml & "    " & m_display & ": " & vbCrLf
        End If
        If Not IsNull(m_text) Then
            yaml = yaml & "    " & m_text.Raw() & ": " & vbCrLf
        End If
        If m_priority > 0 Then
            yaml = yaml & "    " & m_priority & ": " & vbCrLf
        End If
        If m_action <> "add" Then
            yaml = yaml & "    " & m_action & ": " & vbCrLf
        End If
        If m_expire > 0 Then
            yaml = yaml & "    " & m_expire & ": " & vbCrLf
        End If
        If Not IsEmpty(m_flash_mask) Then
            yaml = yaml & "    " & m_flash_mask & ": " & vbCrLf
        End If
        If m_flashing <> "not_set" Then
            yaml = yaml & "    " & m_flashing & ": " & vbCrLf
        End If
        If Not IsEmpty(m_key) Then
            yaml = yaml & "    " & m_key & ": " & vbCrLf
        End If
        If Not IsEmpty(m_color) Then
            yaml = yaml & "    " & m_color & ": " & vbCrLf
        End If
        If Not IsNull(m_transition) Then
            yaml = yaml & m_transition.ToYaml()
        End If
        If Not IsNull(m_transition_out) Then
            yaml = yaml & m_transition_out.ToYaml()
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfSegmentPlayerTransition
	
    private m_type
    private m_text
    private m_direction

    Public Property Get TransitionType() : TransitionType = m_type : End Property
    Public Property Let TransitionType(input) : m_type = input : End Property
    
    Public Property Get Text() : Text = m_text : End Property
    Public Property Let Text(input) : m_text = input : End Property

    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let Direction(input) : m_direction = input : End Property                          

	Public default Function init()
        m_type = "push"
        m_text = Empty
        m_direction = "right"
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    transition:" & vbCrLf
        yaml = yaml & "      " & m_type & ": " & vbCrLf
        yaml = yaml & "      " & m_direction & ": " & vbCrLf
        yaml = yaml & "      " & m_text & ": " & vbCrLf
        ToYaml = yaml
    End Function

End Class

Function SegmentPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim SegmentPlayer : Set SegmentPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            SegmentPlayer.Activate
        Case "deactivate"
            SegmentPlayer.Deactivate
        Case "play"
            If ownProps(2).GlfEvent.Evaluate() Then
                SegmentPlayer.Play ownProps(3), ownProps(2)
            End If
        Case "remove"
            RemoveDelay ownProps(2)
            SegmentPlayer.RemoveTextByKey ownProps(2)
    End Select
    SegmentPlayerEventHandler = Null
End Function


Function SegmentPlayerCallbackHandler(evt, segment_item, mode, priority)

    If IsObject(segment_item) Then
        'Shot Text on a display
        Dim key
        key = mode & "." & "segment_player_player." & segment_item.Display
        If Not IsEmpty(segment_item.Key) Then
            key = key & segment_item.Key
        End If

        Dim display : Set display = glf_segment_displays(segment_item.Display)
        
        If segment_item.Action = "add" Then
            RemoveDelay key
            Dim transition, transition_out : transition = Null : transition_out = Null
            If segment_item.HasTransition() Then
                Set transition = segment_item.Transition
            End If
            If segment_item.HasTransitionOut() Then
                Set transition_out = segment_item.TransitionOut
            End If
            display.AddTextEntry segment_item.Text, segment_item.Color, segment_item.Flashing, segment_item.FlashMask, transition, transition_out, priority + segment_item.Priority, key
                                
            If segment_item.Expire > 0 Then
                SetDelay key & "_expire", "SegmentPlayerEventHandler",  Array(Array("remove", display, key)), segment_item.Expire
            End If

        ElseIf segment_item.Action = "remove" Then
            RemoveDelay key
            display.RemoveTextByKey key        
        ElseIf segment_item.Action = "flash" Then
            display.SetFlashing "all"
        ElseIf segment_item.Action = "flash_match" Then
            display.SetFlashing "match"
        ElseIf segment_item.Action = "flash_mask" Then
            display.SetFlashingMask segment_item.FlashMask
        ElseIf segment_item.Action = "no_flash" Then
            display.SetFlashing "no_flash"
        ElseIf segment_item.Action = "set_color" Then
            If Not IsNull(segment_item.Color) Then
                display.SetColor segment_item.Color
            End If
        End If
    End If

End Function