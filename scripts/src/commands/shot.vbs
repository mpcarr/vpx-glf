
Class GlfShot

    Private m_name
    Private m_mode
    Private m_profile
    Private m_enable_events
    Private m_disable_events
    Private m_switches
    Private m_tokens
    Private m_hit_events
    Private m_start_enabled
    Private m_priority
    Private m_show_cache
    Private m_debug

    Private m_state

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get ShotKey(): ShotKey = m_name & "_" & m_profile: End Property
    Public Property Get Tokens() : Set Tokens = m_tokens : End Property 
    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let Profile(value) : m_profile = value : End Property
    Public Property Let Switches(value) : m_switches = value : End Property
    Public Property Let StartEnabled(value) : m_start_enabled = value : End Property
    Public Property Let HitEvents(value) : m_hit_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "shot_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_profile = ""
        m_state = -1
        m_switches = Array()
        m_start_enabled = True
        m_hit_events = Array()
        Set m_tokens = CreateObject("Scripting.Dictionary")
        m_enable_events = Array()
        Set m_show_cache = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "ShotEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "ShotEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        m_state = 0
        If m_start_enabled Then
            Enable()
        Else
            Dim evt
            For Each evt in m_enable_events
                AddPinEventListener evt, m_mode & "_" & m_name & "_enable", "ShotEventHandler", m_priority, Array("enable", Me)
            Next
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
        m_state = -1
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_mode & "_" & m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_switches
            AddPinEventListener evt & "_active", m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me)
        Next
        For Each evt in m_hit_events
            AddPinEventListener evt, m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me)
        Next
        'Play the show for the active state
        PlayShowForState(m_state)
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_switches
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        For Each evt in m_hit_events
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        Dim x
        For x=0 to Glf_ShotProfiles(m_profile).StatesCount()
            StopShowForState(x)
        Next
    End Sub

    Private Sub StopShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        'If IsObject(profileState) Then
            lightCtrl.RemoveLightSeq m_mode & "_" & m_name, profileState.Key
        'End If
    End Sub

    Private Sub PlayShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        If IsObject(profileState) Then
            If IsArray(profileState.Show) Then
                If m_show_cache.Exists(profileState.Key) Then
                    lightCtrl.AddLightSeq m_mode & "_" & m_name, profileState.Key, m_show_cache(profileState.Key), profileState.Loops, 180/profileState.Speed, Null, m_priority
                Else
                    Dim key
                    Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
                    Dim profileStateTokens : Set profileStateTokens = profileState.Tokens
                    For Each key In profileStateTokens.Keys
                        mergedTokens.Add key, profileStateTokens(key)
                    Next
                    Dim shotTokens : Set shotTokens = m_tokens
                    For Each key In shotTokens.Keys
                        If mergedTokens.Exists(key) Then
                            mergedTokens(key) = shotTokens(key)
                        Else
                            mergedTokens.Add key, shotTokens(key)
                        End If
                    Next
                    Dim show : show = Glf_ConvertShow(profileState.Show, mergedTokens)
                    m_show_cache.Add profileState.Key, show
                    lightCtrl.AddLightSeq m_mode & "_" & m_name, profileState.Key, show, profileState.Loops, 180/profileState.Speed, Null, m_priority
                End If
            End If
        End If
    End Sub

    Public Sub Hit(evt)

        If Glf_ShotProfiles(m_profile).StatesCount() = m_state Then
            If Glf_ShotProfiles(m_profile).ProfileLoop Then
                StopShowForState(m_state)
                m_state = 0
                PlayShowForState(m_state)
            End If
        Else
            StopShowForState(m_state)
            m_state = m_state + 1
            PlayShowForState(m_state)
        End If
        
        If Glf_ShotProfiles(m_profile).Block Then
            Glf_EventBlocks(evt).Add Name, True
        Else
            Glf_EventBlocks(evt).Add ShotKey, True
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function ShotEventHandler(args)
    Dim ownProps, kwargs, e
    ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    e = args(2)
    Dim evt : evt = ownProps(0)
    Dim shot : Set shot = ownProps(1)
    Select Case evt
        Case "activate"
            shot.Activate
        Case "deactivate"
            shot.Deactivate
        Case "enable"
            shot.Enable
        Case "hit"
            If Not Glf_EventBlocks(e).Exists(shot.Name) And Not Glf_EventBlocks(e).Exists(shot.ShotKey) Then
                shot.Hit e
            End If
    End Select
    If IsObject(args(1)) Then
        Set ShotEventHandler = kwargs
    Else
        ShotEventHandler = kwargs
    End If
End Function