
Class ShowPlayer

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

	Public default Function init(name, mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "show_player_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "show_player_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_show_player_play"
            If varType(m_events(evt)) = 8 Then
                Dim showControl : showControl = Split(m_events(evt), ".")
                If showControl(1) = "stop" Then
                    PlayOff evt, showControl(0)
                End If
            Else
                PlayOff evt, m_events(evt).Key
            End If
        Next
    End Sub

    Public Sub Add(evt, show)
        
        If vartype(show) = 8 Then
            m_events.Add evt, show
        Else
            Dim showStep, light, lightsCount, x,tagLight, tagLights, lightParts
            lightsCount = 0
            For Each showStep in show.Show
                For Each light in showStep
                    lightParts = Split(light, "|")
                    If IsArray(lightParts) Then
                        If IsNull(IsToken(lightParts(0))) And IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                            tagLights = lightCtrl.GetLightsForTag(lightParts(0))
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
                For Each light in showStep
                    lightParts = Split(light, "|")
                    If IsArray(lightParts) Then
                        If IsNull(IsToken(lightParts(0))) And IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                            tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                            For Each tagLight in tagLights
                                If UBound(lightParts) >=1 Then
                                    seqArray(x) = tagLight & "|100|"&lightParts(2)
                                Else
                                    seqArray(x) = tagLight & "|100"
                                End If
                                x=x+1
                            Next
                        Else
                            If UBound(lightParts) >= 1 Then
                                seqArray(x) = lightParts(0) & "|100|"&lightParts(2)
                            Else
                                seqArray(x) = lightParts(0) & "|100"
                            End If
                            x=x+1
                        End If
                    End If
                Next
                showStep = seqArray
                Log "Light List: " & Join(seqArray)
            Next
            m_events.Add evt, show
        End If
    End Sub

    Public Sub Play(evt, show)
        If varType(show) = 8 Then
            Dim showControl : showControl = Split(show, ".")
            If showControl(1) = "stop" Then
                For Each evt In m_events.Keys()
                    If IsObject(m_events(evt)) Then
                        PlayOff evt, m_events(evt).Key
                    End If
                Next
            End If
        Else
            lightCtrl.AddLightSeq m_name & "_" & evt, show.Key, show.Show, show.Loops, 180/show.Speed, show.Tokens
        End If
    End Sub

    Public Sub PlayOff(evt, key)
        lightCtrl.RemoveLightSeq m_name & "_" & evt, key
    End Sub

    Private Function IsToken(mainString)
        ' Check if the string contains an opening parenthesis and ends with a closing parenthesis
        If InStr(mainString, "(") > 0 And Right(mainString, 1) = ")" Then
            ' Extract the substring within the parentheses
            Dim startPos, subString
            startPos = InStr(mainString, "(")
            subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
            IsToken = subString
        Else
            IsToken = Null
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_mode & "_show_player", message
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

Class ShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens
  
	Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Show(): Show = m_show: End Property
	Public Property Let Show(input): m_show = input: End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property
  
	Public default Function init()
        If IsEmpty(m_tokens) Then
            Set m_tokens = CreateObject("Scripting.Dictionary")
        End If
	    Set Init = Me
	End Function

    Public Sub AddToken(token, value)
        m_tokens.Add token, value
    End Sub
End Class