

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************

Class GlfVpxBcpController

    Private m_bcpController, m_connected, m_mode_list

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_controller.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        useBcp = True
        If backboxCommand = "" Then
            m_bcpController.EnableLogging()
        End If
        m_mode_list = ""
        AddPinEventListener "player_added", "bcp_player_added", "GlfVpxBcpControllerEventHandler", 50, Array("player_added")
        AddPinEventListener "next_player", "bcp_player_next_player", "GlfVpxBcpControllerEventHandler", 50, Array("next_player")
        AddPinEventListener "game_started", "bcp_player_next_player", "GlfVpxBcpControllerEventHandler", 50, Array("next_player")
        AddPinEventListener "ball_ended", "bcp_player_ball_end", "GlfVpxBcpControllerEventHandler", 50, Array("ball_end")
        If Err Then MsgBox("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
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
            m_bcpController.Send "trigger?name=windowtitle"
        End If
	End Sub
    
    Public Sub PlaySlide(slide, context, calling_context, action, expire, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""" & slide & """: {""action"": """ & action & """, ""expire"": " & expire & "}}, ""context"": """ & context & """, ""calling_context"": """ & calling_context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub PlaySound(sound, context, calling_context, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""sounds_play"", ""settings"": {""" & sound & """: {""action"": ""play"", ""expire"": 0}}, ""context"": """ & context & """, ""calling_context"": """ & calling_context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub PlayWidget(widget, context, calling_context, priorty, expire)
		If m_connected Then
            If IsEmpty(expire) Then
                m_bcpController.Send "trigger?json={""name"": ""widgets_play"", ""settings"": {""" & widget & """: {""action"": ""play"", ""expire"": null, ""x"": 0, ""y"": 0}}, ""context"": """ & context & """, ""calling_context"": """ & calling_context & """, ""priority"": " & priorty & "}"
            Else
                m_bcpController.Send "trigger?json={""name"": ""widgets_play"", ""settings"": {""" & widget & """: {""action"": ""play"", ""expire"": " & expire & " , ""x"": 0, ""y"": 0}}, ""context"": """ & context & """, ""calling_context"": """ & calling_context & """, ""priority"": " & priorty & "}"
            End If
        End If
	End Sub

    Public Sub ModeList()
        If m_connected Then
            If m_mode_list <> glf_running_modes Then
                m_bcpController.Send "mode_list?json={""running_modes"": ["&glf_running_modes&"]}"
                m_mode_list = glf_running_modes
            End If
        End If
    End Sub

    Public Sub SlidesClear(context)
        If m_connected Then
            m_bcpController.Send "trigger?name=slides_clear&context=" & context
        End If
    End Sub

    Public Sub ModeStart(name, priority)
        If m_connected Then
            m_bcpController.Send "mode_start?priority=int:" & priority & "&name=" & name
        End If
    End Sub

    Public Sub ModeStop(name)
        If m_connected Then
            m_bcpController.Send "mode_stop?name=" & name
        End If
    End Sub


    Public Sub SendPlayerVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & Getglf_currentPlayerNumber+1
        End If
	End Sub

    Public Sub SendMachineVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "machine_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue))
        End If
	End Sub

    Private Function EncodeVariable(value)
        Dim retValue
        Select Case VarType(value)
            Case vbInteger, vbLong
                retValue = "int:" & value
            Case vbSingle, vbDouble
                retValue = "float:" & value
            Case vbString
                retValue = value
            Case vbBoolean
                retValue = "bool:" & CStr(value)
            Case Else
                retValue = "NoneType:"
        End Select
        EncodeVariable = retValue
    End Function
    
    Private Function VariableVariance(v1, v2)
        Dim retValue
        Select Case VarType(v1)
            Case vbInteger, vbLong, vbSingle, vbDouble
                retValue = Abs(v1 - v2)
            Case Else
                retValue = True 
        End Select
        VariableVariance = retValue
    End Function

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
            useBcp = False
        End If
    End Sub
End Class

Sub Glf_BcpSendPlayerVar(args)
    If useBcp=False Then
        Exit Sub
    End If
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim player_var : player_var = kwargs(0)
    Dim value : value = kwargs(1)
    Dim prevValue : prevValue = kwargs(2)
    bcpController.SendPlayerVariable player_var, value, prevValue
End Sub

Sub Glf_BcpSendMachineVar(key, value, prevValue)
    If useBcp=False Then
        Exit Sub
    End If
    bcpController.SendMachineVariable key, value, prevValue
End Sub

Sub Glf_BcpSendEvent(evt)
    If useBcp=False Then
        Exit Sub
    End If
    bcpController.Send "trigger?name=" & evt
End Sub

Sub Glf_BcpAddPlayer(playerNum)
    If useBcp Then
        bcpController.Send("player_added?player_num=int:"&playerNum)
        dim p
        Select Case playerNum
            Case 1:
                p = "PLAYER 1"
            Case 2:
                p = "PLAYER 2"
            Case 3:
                p = "PLAYER 3"
            Case 4:
                p = "PLAYER 4"
        End Select
        Dim key, player_state
        Set player_state = glf_playerState(p)
        For Each key in player_state.Keys()
            Glf_BcpSendPlayerVar Array(Null, Array(key, player_state(key), Empty))
        Next
    End If
End Sub

Sub Glf_BcpUpdate()
    If useBcp=False Then
        Exit Sub
    End If
    Dim messages : messages = bcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter, eventName
        For Each message in messages
            glf_debugLog.WriteToLog "BCP Update", message.GetValue("RawMessage")
            Select Case message.Command
                case "hello"
                    bcpController.Reset
                case "monitor_start"
                    Dim category : category = message.GetValue("category")
                    If category = "player_vars" Then
                        glf_monitor_player_vars = True
                        'AddPlayerStateEventListener "score", "bcp_player_var_score_0", 0, "Glf_BcpSendPlayerVar", 1000, Null
                        'AddPlayerStateEventListener "current_ball", "bcp_player_var_ball_0", 0, "Glf_BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    eventName = message.GetValue("event")
                    glf_debugLog.WriteToLog "BCP Update: Register Trigger", eventName
                case "trigger"
                    eventName = message.GetValue("name")
                    glf_debugLog.WriteToLog "BCP Update: Trigger", eventName
                    DispatchPinEvent eventName, Null
            End Select
        Next
    End If
    bcpController.ModeList()
End Sub

Function GlfVpxBcpControllerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Select Case evt
        Case "player_added"
            bcpController.Send "player_added?player_num=int:" & kwargs("num")
        Case "next_player"
            Dim p_num : p_num = Getglf_currentPlayerNumber() + 1
            bcpController.Send "player_turn_start?player_num=int:" & p_num
            bcpController.Send "ball_start?player_num=int:" & p_num & "&ball=int:" & GetPlayerState("ball")
            bcpController.SendPlayerVariable "number", p_num, p_num-1
            'bcpController.SendPlayerVariable "number", 1, 0
        Case "ball_end"
            bcpController.Send "ball_end"
    End Select
    If IsObject(args(1)) Then
        Set GlfVpxBcpControllerEventHandler = kwargs
    Else
        GlfVpxBcpControllerEventHandler = kwargs
    End If
    
End Function

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************