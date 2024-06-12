
Class ShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Let Events(value) : Set m_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", "show_player_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "show_player_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_show_player_play", "ShowPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_show_player_play"
        Next
    End Sub

    Public Sub Play(showItem)
        Log "Playing " & showItem.Name
        Dim show_step, stepIdx, lastTime
        stepIdx = 0
        lastTime = 125
        For Each show_step in showItem.Show
            lastTime = lastTime + show_step.Time
            SetDelay m_mode & "_show_player_play_step_" & stepIdx, "ShowPlayerEventHandler", Array(Array("play_step", Me), show_step), lastTime            
            stepIdx = stepIdx + 1
        Next
    End Sub

    Public Sub PlayStep(showStep)
        Dim light
        Log "Playing Step"
        For Each light in showStep.Lights
            If light(1) = "off" Then
                lightCtrl.LightOff light(0)
            Else
                If UBound(light) = 2 Then
                    lightCtrl.LightOn light(0)
                    lightCtrl.FadeLightToColor light(0), light(1), light(2)
                Else
                    lightCtrl.LightOnWithColor light(0), light(1)
                End If
            End If
        Next
    End Sub

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
            ShowPlayer.Play ownProps(2)
        Case "play_step"
            Dim show_step : Set show_step = args(1)
            ShowPlayer.PlayStep show_step
    End Select
    ShowPlayerEventHandler = Null
End Function

Class ShowPlayerItem

    Private m_name
    Private m_priority
    Private m_mode
    Private m_show
    Private m_speed
    Private m_tokens
    Private m_debug

    Private m_value

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Show(): Show = m_show: End Property

    Public Property Let Speed(value) : m_speed = value : End Property
    Public Property Let Tokens(value) : m_tokens = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode, show)
        m_mode = mode.Name
        m_name = m_mode & "_show_player_" & name
        m_priority = mode.Priority
        m_show = show
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        'Dim evt
        'For Each evt In m_events.Keys()
        '    AddPinEventListener evt, m_mode & "_show_player_play", "ShowPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        'Next
    End Sub

    Public Sub Deactivate()
        'Dim evt
        'For Each evt In m_events.Keys()
        '    RemovePinEventListener evt, m_mode & "_show_player_play"
        'Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Class ShowPlayerLightStep 

    Private m_time
    Private m_lights
    Private m_debug

    Public Property Get Time(): Time = m_time: End Property
    Public Property Get Lights(): Lights = m_lights: End Property

    Public Property Let Time(value) : m_time = value : End Property
    Public Property Let Lights(value) : m_lights = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(time, lights)
        m_time = time
        m_lights = lights
        m_debug = False
        Set Init = Me
	End Function

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

