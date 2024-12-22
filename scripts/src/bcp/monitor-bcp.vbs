

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
            
            Dim mode, mode_json
            mode_json = ""
            For Each mode in glf_modes.Items()
                mode_json ="{""mode"": """&mode.Name&""", ""value"": """&mode.Status&""", ""debug"": " & mode.IsDebug & "}," 
            Next
            m_bcpController.Send "glf_monitor?json={""name"": ""glf_monitor_modes"", ""changes"": [" & mode_json & "]}"
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
'*****************************************************************************************************************************************