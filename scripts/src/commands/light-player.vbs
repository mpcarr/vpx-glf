
Class LightPlayer

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
        
        AddPinEventListener m_mode & "_starting", "light_player_activate", "LightPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "light_player_deactivate", "LightPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_light_player_play"
        Next
    End Sub

    Public Sub Play(lights)
        Dim light
        For Each light in lights
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
            debugLog.WriteToLog m_mode & "_light_player", message
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
            LightPlayer.Play ownProps(2)
    End Select
    LightPlayerEventHandler = Null
End Function

