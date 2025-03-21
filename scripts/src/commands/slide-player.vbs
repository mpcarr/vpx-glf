

Class GlfSlidePlayer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "slide_player" : End Property

    Public Property Get EventName(name)
        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_slide : Set new_slide = (new GlfSlidePlayerItem)()
        m_eventValues.Add newEvent.Raw, new_slide
        Set EventName = new_slide
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "slide_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "slide_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_slide_player_play", "SlidePlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_slide_player_play"
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            'Fire Slide
            bcpController.PlaySlide m_eventValues(evt).Slide, m_mode, m_events(evt).EventName, m_priority
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_slide_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function SlidePlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim slide_player : Set slide_player = ownProps(1)
    Select Case evt
        Case "activate"
            slide_player.Activate
        Case "deactivate"
            slide_player.Deactivate
        Case "play"
            slide_player.Play(ownProps(2))
    End Select
    If IsObject(args(1)) Then
        Set SlidePlayerEventHandler = kwargs
    Else
        SlidePlayerEventHandler = kwargs
    End If
End Function

Class GlfSlidePlayerItem
	Private m_slide, m_action, m_expire, m_max_queue_time, m_method, m_priority, m_target, m_tokens
    
    Public Property Get Slide(): Slide = m_slide: End Property
    Public Property Let Slide(input)
        m_slide = input
    End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        m_action = input
    End Property

    Public Property Get Expire(): Expire = m_expire: End Property
    Public Property Let Expire(input)
        m_expire = input
    End Property

    Public Property Get MaxQueueTime(): MaxQueueTime = m_max_queue_time: End Property
    Public Property Let MaxQueueTime(input)
        m_max_queue_time = input
    End Property

    Public Property Get Method(): Method = m_method: End Property
    Public Property Let Method(input)
        m_method = input
    End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input)
        m_priority = input
    End Property

    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input)
        m_target = input
    End Property

    Public Property Get Tokens(): Tokens = m_tokens: End Property
    Public Property Let Tokens(input)
        m_tokens = input
    End Property

	Public default Function init()
        m_action = "play"
        m_slide = Empty
        m_expire = Empty
        m_max_queue_time = Empty
        m_method = Empty
        m_priority = Empty
        m_target = Empty
        m_tokens = Empty
        Set Init = Me
	End Function

End Class
