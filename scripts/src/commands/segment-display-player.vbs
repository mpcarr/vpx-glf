
Class GlfSegmentPlayer

    Private m_priority
    Private m_mode
    Private m_name
    Private m_events
    private m_base_device

    Public Property Get Name() : Name = "segment_player" : End Property

    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get Events(name)
        If m_events.Exists(name) Then
            Set Events = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfSegmentPlayerEventItem)()
            m_events.Add name, new_event
            Set Events = new_event
        End If
    End Property

	Public default Function init(mode)
        m_name = "segment_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "segment_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_segment_player_play", "SegmentPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_segment_player_play"
            PlayOff evt, m_events(evt)
        Next
    End Sub

    Public Sub Play(evt, segment_item)
        SegmentPlayerCallbackHandler evt, segment_item, m_mode, m_priority
    End Sub

    Public Sub PlayOff(evt, segment_item)
        SegmentPlayerCallbackHandler evt, Null, m_mode, m_priority
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

    Public Property Get Display() : Display = m_display : End Property
    Public Property Let Display(input) : m_display = input : End Property
    
    Public Property Get Text() : Text = m_text : End Property
    Public Property Let Text(input) : m_text = input : End Property

    Public Property Get Priority() : Priority = m_priority : End Property
    Public Property Let Priority(input) : m_priority = input : End Property
        
    Public Property Get Action() : Action = m_action : End Property
    Public Property Let Action(input) : m_action = input : End Property
                
    Public Property Get Expire() : Expire = m_expire : End Property
    Public Property Let Expire(input) : m_events = input : End Property

    Public Property Get FlashMask() : FlashMask = m_flash_mask : End Property
    Public Property Let FlashMask(input) : m_flash_mask = input : End Property
                        
    Public Property Get Flashing() : Flashing = m_flashing : End Property
    Public Property Let Flashing(input) : m_flashing = input : End Property
                            
    Public Property Get Key() : Key = m_key : End Property
    Public Property Let Key(input) : m_key = input : End Property

    Public Property Get Transition()
        If IsNull(m_transition) Then
            Set m_transition = (new GlfSegmentPlayerTransition)()
            Set Transition = m_transition   
        Else
            Set Transition = m_transition
        End If
    End Property
                                
	Public default Function init()
        m_display = Empty
        m_text = Empty
        m_priority = 0
        m_action = "add"
        m_expire = 0
        m_flash_mask = Empty
        m_flashing = "not_set"
        m_key = Empty
        m_transition = Null
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        If Not IsEmpty(m_display) Then
            yaml = yaml & "    " & m_display & ": " & vbCrLf
        End If
        If Not IsEmpty(m_text) Then
            yaml = yaml & "    " & m_text & ": " & vbCrLf
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
        If Not IsNull(m_transition) Then
            yaml = yaml & m_transition.ToYaml()
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
        m_direction = "push"
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
            SegmentPlayer.Play ownProps(3), ownProps(2)
    End Select
    SegmentPlayerEventHandler = Null
End Function


Function SegmentPlayerCallbackHandler(evt, segment_item, mode, priority)
    'Shot Text on a display
    Dim key
    key = mode & "." & "segment_player_player." & segment_item.Display
    
    If Not IsEmpty(segment_item.Key) Then
        key = key & segment_item.Key
    End If

    Dim display : Set display = glf_segment_displays(segment_item.Display)
    
    If segment_item.Action = "add" Then
        RemoveDelay key
        ' add text
        s = TransitionManager.validate_config(s, self.machine.config_validator)
        
           
        display.AddTextEntry segment_item.Text, segment_item.Color, segment_item.Flashing, segment_item.FlashMask, segment_item.Transition, segment_item.TransitionOut, segment_item.Priority, segment_item.Key
                            
        If segment_item.Expire > 0 Then
            'TODO Add delay for remove
            'SetDelay
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


End Function