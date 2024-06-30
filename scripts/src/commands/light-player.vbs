
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
        Dim light, lightParts
        lightCtrl.AddLightSeq m_name & "_" & evt, evt, lights(1), -1, 180, Null, m_priority
        For Each light in lights(0)
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    Dim tagLight, tagLights
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        ProcessLight tagLight, lightParts, evt
                    Next
                Else
                    ProcessLight lightParts(0), lightParts, evt
                End If
            End If
        Next
    End Sub

    Public Sub PlayOff(evt, lights)
        Dim light
        lightCtrl.RemoveLightSeq m_name & "_" & evt, evt
    End Sub

    Private Sub ProcessLight(light_name, light_props, evt)
        If UBound(light_props) = 2 Then
            lightCtrl.FadeLightToColor light_name, light_props(1), light_props(2), m_name & "_" & evt & "_" & light_name, m_priority
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_light_player", message
        End If
    End Sub
End Class

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

