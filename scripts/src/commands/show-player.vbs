
Class GlfShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_show_cache
    Private m_debug
    Private m_name
    Private m_value

    Public Property Get Events(name)
        If m_events.Exists(name) Then
            Set Events = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfShowPlayerItem)()
            m_events.Add name, new_event
            Set Events = new_event
        End If
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = True
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_show_cache = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "show_player_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "show_player_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            If IsObject(m_events(evt)) Then
                AddPinEventListener Replace(evt, "_" & m_events(evt).Key, "") , m_mode & "_" & m_events(evt).Key & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
            Else
                AddPinEventListener Replace(evt, "_" & m_events(evt), "") , m_mode & "_" & m_events(evt) & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
            End If
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            If IsObject(m_events(evt)) Then
                RemovePinEventListener Replace(evt, "_" & m_events(evt).Key, ""), m_mode & "_" & m_events(evt).Key & "_show_player_play"
            Else
                RemovePinEventListener Replace(evt, "_" & m_events(evt), ""), m_mode & "_" & m_events(evt) & "_show_player_play"
            End If
            PlayOff m_events(evt).Key
        Next
    End Sub

    Public Sub Play(evt, show)
        
        If show.Action = "stop" Then
            PlayOff show.Key
        Else
            If m_show_cache.Exists(show.Key) Then
                lightCtrl.AddLightSeq m_name & "_" & show.Key, show.Key, m_show_cache(show.Key), show.Loops, show.Speed, Null, m_priority, 0
            Else
                Dim cachedShow : cachedShow = Glf_ConvertShow(show.Show, show.Tokens)
                m_show_cache.Add show.Key, cachedShow
                lightCtrl.AddLightSeq m_name & "_" & show.Key, show.Key, cachedShow, show.Loops, show.Speed, Null, m_priority, 0
            End If
        End If
    End Sub

    'Public Sub StopShow(evt, key)
    '    m_events.Add evt & "_" & key & ".stop", key & ".stop"
    'End Sub

    Public Sub PlayOff(key)
        lightCtrl.RemoveLightSeq m_name & "_" & key, key
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_show_player", message
        End If
    End Sub
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
            ShowPlayer.Play ownProps(3), ownProps(2)
    End Select
    ShowPlayerEventHandler = Null
End Function

Class GlfShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Show(): Show = m_show: End Property
	Public Property Let Show(input): m_show = input: End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property        

    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property        
  
	Public default Function init()
        m_action = "play"
        m_key = ""
        m_loops = -1
        m_speed = 1
        m_syncms = 0
        Set m_tokens = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class
