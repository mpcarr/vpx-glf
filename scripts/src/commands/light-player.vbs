
Class GlfLightPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug
    Private m_name
    Private m_value

    Public Property Let Events(value)
        Set m_events = value
        Dim evt
        For Each evt in m_events
            lightCtrl.CreateSeqRunner m_name & "_" & evt, m_priority
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_name = "light_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "light_player_activate", "LightPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "light_player_deactivate", "LightPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_light_player_play"
            PlayOff evt, m_events(evt)
        Next
    End Sub

    Public Sub Add(evt, lights)
        
        Dim light, lightsCount, x,tagLight, tagLights, lightParts
        lightsCount = 0
        For Each light in lights
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            Else
                If IsNull(lightCtrl.GetLightIdx(lightParts)) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts)
                    Log "Tag Lights: " & Join(tagLights)
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            End If
        Next
        Log "Adding " & lightsCount & " lights for event: " & evt 
        Dim seqArray
        ReDim seqArray(lightsCount-1)
        x=0
        For Each light in lights
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        If UBound(lightParts) >=1 Then
                            seqArray(x) = tagLight & "|100|"&lightParts(1)
                        Else
                            seqArray(x) = tagLight & "|100"
                        End If
                        x=x+1
                    Next
                Else
                    If UBound(lightParts) >= 1 Then
                        seqArray(x) = lightParts(0) & "|100|"&lightParts(1)
                    Else
                        seqArray(x) = lightParts(0) & "|100"
                    End If
                    x=x+1
                End If
            Else
                If IsNull(lightCtrl.GetLightIdx(lightParts)) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts)
                    For Each tagLight in tagLights
                        seqArray(x) = tagLight & "|100"
                        x=x+1
                    Next
                Else
                    seqArray(x) = lightParts & "|100"
                    x=x+1
                End If
            End If
        Next
        Log "Light List: " & Join(seqArray)
        m_events.Add evt, Array(lights, Array(seqArray))
    End Sub

    Public Sub Play(evt, lights)
        LightPlayerCallbackHandler evt, lights(1), m_name, m_priority
    End Sub

    Public Sub PlayOff(evt, lights)
        LightPlayerCallbackHandler evt, Null, m_name, m_priority
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_light_player", message
        End If
    End Sub
End Class

Function LightPlayerCallbackHandler(key, lights, mode, priority)
    If IsNull(lights) Then
        lightCtrl.RemoveLightSeq mode & "_" & key, key
        glf_debugLog.WriteToLog "LightPlayer", "Removing Light Seq" & mode & "_" & key
    Else
        If UBound(lights) = -1 Then
            Exit Function
        End If
        If IsArray(lights) Then
            'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key    
        Else
            glf_debugLog.WriteToLog "LightPlayer", "Lights not an array!?"
        End If
        Dim light, lightParts
        'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key
        lightCtrl.AddLightSeq mode & "_" & key, key, lights, -1, 1, Null, priority, 0,100000
        
        'TODO - Refactor this, this is the light fading. need to handle this differently
        'For Each light in lights(0)
        '    lightParts = Split(light, "|")
        '    If IsArray(lightParts) Then
        '        If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
        '           Dim tagLight, tagLights
        '            tagLights = lightCtrl.GetLightsForTag(lightParts(0))
        '            For Each tagLight in tagLights
        '                ProcessLight tagLight, lightParts, key
        '                If UBound(lightParts) = 2 Then
        '                    lightCtrl.FadeLightToColor tagLight, lightParts(1), lightParts(2), mode & "_" & key & "_" & tagLight, priority
        '                End If
        '            Next
        '        Else
        '            If UBound(lightParts) = 2 Then
        '                lightCtrl.FadeLightToColor lightParts(0), lightParts(1), lightParts(2), mode & "_" & key & "_" & lightParts(0), priority
        '            End If
        '        End If
        '    End If
        'Next
    End If
End Function

Function LightPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim LightPlayer : Set LightPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            LightPlayer.Activate
        Case "deactivate"
            LightPlayer.Deactivate
        Case "play"
            LightPlayer.Play ownProps(3), ownProps(2)
    End Select
    LightPlayerEventHandler = Null
End Function

