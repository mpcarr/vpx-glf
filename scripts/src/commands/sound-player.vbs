
Class GlfSoundPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_eventValues
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "sound_player" : End Property
    Public Property Get EventSounds() : EventSounds = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_sound : Set new_sound = (new GlfSoundPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_sound
        Set EventName = new_sound
        
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "sound_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "sound_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_sound_player_play", "SoundPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_sound_player_play"
            PlayOff evt
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            If m_eventValues(evt).Action = "stop" Then
                PlayOff evt
            Else
                glf_sound_buses(m_eventValues(evt).Sound.Bus).Play m_eventValues(evt)
            End If
        End If
    End Function

    Public Sub PlayOff(evt)
        glf_sound_buses(m_eventValues(evt).Sound.Bus).StopSoundWithKey m_eventValues(evt).Sound.File
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt, key
        If UBound(m_eventValues.Keys) > -1 Then
            For Each key in m_eventValues.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_eventValues(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function SoundPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim sound_player : Set sound_player = ownProps(1)
    Select Case evt
        Case "activate"
            sound_player.Activate
        Case "deactivate"
            sound_player.Deactivate
        Case "play"
            'Dim block_queue
            sound_player.Play(ownProps(2))
            'If Not IsEmpty(block_queue) Then
            '    kwargs.Add "wait_for", block_queue
            'End If
    End Select
    If IsObject(args(1)) Then
        Set SoundPlayerEventHandler = kwargs
    Else
        SoundPlayerEventHandler = kwargs
    End If
End Function


Class GlfSoundPlayerItem
	Private m_sound, m_action, m_key, m_volume, m_loops
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property

    Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Sound()
        If IsNull(m_sound) Then
            Sound = Null
        Else
            Set Sound = m_sound
        End If
    End Property
	Public Property Let Sound(input)
        If glf_sounds.Exists(input) Then
            Set m_sound = glf_sounds(input)
        End If
    End Property
  
	Public default Function init()
        m_action = "play"
        m_sound = Null
        m_key = Empty
        m_volume = Empty
        m_loops = Empty
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    " & Sound.NameRaw & ": " & vbCrLf
        If Not IsEmpty(m_key) Then
            yaml = yaml & "      key: " & m_key & vbCrLf
        End If
        yaml = yaml & "      action: " & m_action & vbCrLf
        If Not IsEmpty(m_volume) Then
            yaml = yaml & "      volume: " & m_volume & vbCrLf
        End If
        If Not IsEmpty(m_loops) Then
            yaml = yaml & "      loops: " & m_loops & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class
