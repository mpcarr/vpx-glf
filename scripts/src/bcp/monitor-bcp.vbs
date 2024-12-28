

'*****************************************************************************************************************************************
'  Glf Monitor Bcp Controller
'*****************************************************************************************************************************************

Class GlfMonitorBcpController

    Private m_bcpController, m_connected, m_isInMonitor

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_server.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        If Err Then MsgBox("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

    Public Function IsInMonitior()
        If m_connected = True And m_isInMonitor = True Then
            IsInMonitior=True
        Else
            IsInMonitior=False
        End If
    End Function

	Public Sub Send(commandMessage)
		If m_connected = True Then
            m_bcpController.Send commandMessage
        End If
	End Sub

    Public Function GetMessages
		If m_connected Then
            GetMessages = m_bcpController.GetMessages
        End If
	End Function

    Public Sub Reset()
		If m_connected Then
            m_bcpController.Send "reset"
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""monitor"": {""action"": ""play"", ""expire"": 0}}, ""context"": """", ""priority"": 1}"
            
            Dim mode
            For Each mode in glf_modes.Items()
                glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """&mode.Status&""", ""debug"": " & mode.IsDebug & "}," 
            Next
            m_isInMonitor = True
        End If
	End Sub

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
        End If
    End Sub
End Class

Sub Glf_MonitorModeUpdate(mode)
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If
    glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """&mode.Status&""", ""debug"": " & mode.IsDebug & "},"
    Dim config_item
    For Each config_item in mode.BallSavesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.CountersItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.TimersItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.MultiballLocksItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.MultiballsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ModeShots()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ShotGroupsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.BallHoldsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.SequenceShotsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ModeStateMachines()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    If Not IsNull(mode.LightPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.LightPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.LightPlayer.Name & """},"
    End If
    If Not IsNull(mode.EventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.EventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.EventPlayer.Name & """},"
    End If
    If Not IsNull(mode.RandomEventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.RandomEventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.RandomEventPlayer.Name & """},"
    End If
    If Not IsNull(mode.ShowPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.ShowPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.ShowPlayer.Name & """},"
    End If
    If Not IsNull(mode.SegmentDisplayPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.SegmentDisplayPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.SegmentDisplayPlayer.Name & """},"
    End If
    If Not IsNull(mode.VariablePlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.VariablePlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.VariablePlayer.Name & """},"
    End If
End Sub

Sub Glf_MonitorPlayerStateUpdate(key, value)
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If    
    glf_monitor_player_state = glf_monitor_player_state & "{""key"": """&key&""", ""value"": """&value&"""},"
End Sub

Sub Glf_MonitorEventStream(label, message)
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If
    glf_monitor_event_stream = glf_monitor_event_stream & "{""label"": """&label&""", ""message"": """&message&"""},"
End Sub


Sub Glf_MonitorBcpUpdate()
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If

    'Send Updates
    If glf_debugBcpController.IsInMonitior Then
        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_player_state"",""changes"": ["&glf_monitor_player_state&"]}"
        glf_monitor_player_state = ""

        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_monitor_modes"",""changes"": ["&glf_monitor_modes&"]}"
        glf_monitor_modes = ""

        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_event_stream"",""changes"": ["&glf_monitor_event_stream&"]}"
        glf_monitor_event_stream = ""
    End If
    

    Dim messages : messages = glf_debugBcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter, eventName
        For Each message in messages
            debug.print(message.Command)
            Select Case message.Command
                case "hello"
                    glf_debugBcpController.Reset
                case "trigger"
                    eventName = message.GetValue("name")
                    debug.print eventName
                    If eventName = "glf_monitor_debug_mode" Then
                        Dim mode_name : mode_name = message.GetValue("mode")
                        If Not IsNull(GlfModes(mode_name)) Then
                            debug.print("got mode")
                            If GlfModes(mode_name).IsDebug = 1 Then
                                debug.print("Turning off debug")
                                GlfModes(mode_name).Debug = False
                            Else
                                debug.print("Turning on debug")
                                GlfModes(mode_name).Debug = True
                            End If
                        End If
                    End If
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************a