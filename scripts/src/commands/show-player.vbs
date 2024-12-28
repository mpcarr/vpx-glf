
Class GlfShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_eventValues
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "show_player" : End Property
    Public Property Get EventShows() : EventShows = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_show : Set new_show = (new GlfShowPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_show
        Set EventName = new_show
        
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "show_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName , m_mode & "_" & m_eventValues(evt).Key & "_show_player_play", "ShowPlayerEventHandler", m_priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_show_player_play"
            PlayOff m_eventValues(evt).Key
        Next
    End Sub

    Public Sub Play(evt)
        If m_events(evt).Evaluate() Then
            If m_eventValues(evt).Action = "stop" Then
                PlayOff m_eventValues(evt).Key
            Else
                Dim new_running_show
                Set new_running_show = (new GlfRunningShow)(m_name & "_" & m_eventValues(evt).Key, m_eventValues(evt).Key, m_eventValues(evt), m_priority, Null, Null)
            End If
        End If
    End Sub

    Public Sub PlayOff(key)
        If glf_running_shows.Exists(m_name & "_" & key) Then 
            glf_running_shows(m_name & "_" & key).StopRunningShow()
        End If
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
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function ShowPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ShowPlayer : Set ShowPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            ShowPlayer.Activate
        Case "deactivate"
            ShowPlayer.Deactivate
        Case "play"
            ShowPlayer.Play ownProps(2)
    End Select
    ShowPlayerEventHandler = Null
End Function

Class GlfShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms, m_duration, m_priority, m_internal_cache_id
  
	Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property

    Public Property Get Show()
        If IsNull(m_show) Then
            Show = Null
        Else
            Set Show = m_show
        End If
    End Property
	Public Property Let Show(input)
        'msgbox "input:" & input
        If glf_shows.Exists(input) Then
            Set m_show = glf_shows(input)
        End If
    End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property        

    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property        
  
    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    " & m_show.Name &": " & vbCrLf
        If m_action <> "play" Then
            yaml = yaml & "      action: " & m_action & vbCrLf
        End If
        If m_key <> "" Then
            yaml = yaml & "      key: " & m_key & vbCrLf
        End If
        If m_priority <> 0 Then
            yaml = yaml & "      priority: " & m_priority & vbCrLf
        End If
        If m_loops > -1 Then
            yaml = yaml & "      loops: " & m_loops & vbCrLf
        End If
        If m_speed <> 1 Then
            yaml = yaml & "      speed: " & m_speed & vbCrLf
        End If
        If UBound(m_tokens.Keys) > -1 Then
            yaml = yaml & "      show_tokens: " & vbCrLf
            Dim key
            For Each key in m_tokens.Keys
                yaml = yaml & "        " & key & ": " & m_tokens(key) & vbCrLf
            Next
        End If
        If m_syncms > 0 Then
            yaml = yaml & "      sync_ms: " & m_syncms & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        m_action = "play"
        m_key = ""
        m_priority = 0
        m_loops = -1
        m_internal_cache_id = -1
        m_speed = 1
        m_syncms = 0
        m_show = Null
        Set m_tokens = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class
