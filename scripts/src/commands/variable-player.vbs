Class GlfVariablePlayer

    Private m_priority
    Private m_name
    Private m_mode
    Private m_events
    Private m_debug
    private m_base_device

    Private m_value

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property 
    Public Property Get EventName(name)
        Dim newEvent : Set newEvent = (new GlfVariablePlayerEvent)(name)
        m_events.Add newEvent.BaseEvent.Raw, newEvent
        Set EventName = newEvent
    End Property
   
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "variable_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "variable_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_" & evt & "_variable_player_play", "VariablePlayerEventHandler", m_priority+m_events(evt).BaseEvent.Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_" & evt & "_variable_player_play"
        Next
    End Sub

    Public Sub Play(evt)
        Log "Playing: " & evt
        If m_events(evt).BaseEvent.Evaluate() = False Then
            Exit Sub
        End If
        Dim vKey, v
        For Each vKey in m_events(evt).Variables.Keys
            Set v = m_events(evt).Variable(vKey)
            Dim varValue : varValue = v.VariableValue
            Select Case v.Action
                Case "add"
                    Log "Add Variable " & vKey & ". New Value: " & CStr(GetPlayerState(vKey) + varValue) & " Old Value: " & CStr(GetPlayerState(vKey))
                    SetPlayerState vKey, GetPlayerState(vKey) + varValue
                Case "add_machine"
                    Log "Add Machine Variable " & vKey & ". New Value: " & CStr(GetPlayerState(vKey) + varValue) & " Old Value: " & CStr(GetPlayerState(vKey))
                    'SetPlayerState vKey, GetPlayerState(vKey) + varValue
                    glf_machine_vars(vkey).Value = glf_machine_vars(vkey).Value + varValue
                Case "set"
                    Log "Setting Variable " & vKey & ". New Value: " & CStr(varValue)
                    SetPlayerState vKey, varValue
                Case "set_machine"
                    Log "Setting Machine Variable " & vKey & ". New Value: " & CStr(varValue)
                    glf_machine_vars(vkey).Value = varValue
        End Select
        Next
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim key
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_variable_player", message
        End If
    End Sub
End Class

Class GlfVariablePlayerEvent

    Private m_event
	Private m_variables

    Public Property Get BaseEvent() : Set BaseEvent = m_event : End Property
  
	Public Property Get Variable(name)
        If m_variables.Exists(name) Then
            Set Variable = m_variables(name)
        Else
            Dim new_variable : Set new_variable = (new GlfVariablePlayerItem)()
            m_variables.Add name, new_variable
            Set Variable = new_variable
        End If
    End Property
    
    Public Property Get Variables(): Set Variables = m_variables End Property

	Public default Function init(evt)
        Set m_event = (new GlfEvent)(evt)
        Set m_variables = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        Dim key
        If UBound(m_variables.Keys) > -1 Then
            For Each key in m_variables.keys
                yaml = yaml & "    " & key & ": " & vbCrLf
                yaml = yaml & m_variables(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfVariablePlayerItem
	Private m_block, m_show, m_float, m_int, m_string, m_player, m_action, m_type
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Block(): Block = m_block End Property
    Public Property Let Block(input): m_block = input End Property

	Public Property Let Float(input): Set m_float = CreateGlfInput(input): m_type = "float" : End Property
  
	Public Property Let Int(input): Set m_int = CreateGlfInput(input): m_type = "int" : End Property
  
	Public Property Let String(input) : Set m_string = CreateGlfInput(input) : m_type = "string" : End Property

    Public Property Get VariableType(): VariableType = m_type: End Property
    Public Property Get VariableValue()
        Select Case m_type
            Case "float"
                VariableValue = m_float.Value()
            Case "int"
                VariableValue = m_int.Value()
            Case "string"
                VariableValue = m_string.Value()
            Case Else
                VariableValue = Empty
        End Select
    End Property

    Public Property Get Player(): Player = m_player: End Property
    Public Property Let Player(input): m_player = input: End Property

	Public default Function init()
        m_action = "add"
        m_type = Empty
        m_block = False
        m_float = Empty
        m_int = Empty
        m_string = Empty
        m_player = Empty
	    Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "      "& "action" & ": " & m_action & vbCrLf
        If m_block = True Then
            yaml = yaml & "      "& "block" & ": true" & vbCrLf
        End If
        If Not IsEmpty(m_int) Then
            yaml = yaml & "      "& "int" & ": " & m_int.Raw & vbCrLf
        End If
        If Not IsEmpty(m_float) Then
            yaml = yaml & "      "& "float" & ": " & m_float.Raw & vbCrLf
        End If
        If Not IsEmpty(m_string) Then
            yaml = yaml & "      "& "string" & ": " & m_string.Raw & vbCrLf
        End If
        If Not IsEmpty(m_player) Then
            yaml = yaml & "      "& "player" & ": " & m_strinm_playerg & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class


Function CreateMachineVar(name)
	Dim machine_var : Set machine_var = (new GlfMachineVars)(name)
	Set CreateMachineVar = machine_var
End Function

Class GlfMachineVars

    Private m_name
	Private m_persist
    Private m_value_type
    Private m_value

    Public Property Get Name(): Name = m_name : End Property
    Public Property Let Name(input): m_name = input : End Property

    Public Property Get Value(): Value = m_value : End Property
    Public Property Let Value(input): m_value = input : End Property

    Public Property Let InitialValue(input)
        m_value = input
    End Property

    Public Property Get Persist(): Persist = m_persist : End Property
    Public Property Let Persist(input): m_persist = input : End Property

    Public Property Get ValueType(): ValueType = m_value_type : End Property
    Public Property Let ValueType(input): m_value_type = input : End Property

    Public Function GetValue()
        GetValue = m_value
    End Function

	Public default Function init(name)
        m_name = name
        m_persist = True
        m_value_type = "int"
        m_value = 0
        glf_machine_vars.Add name, Me
	    Set Init = Me
	End Function
End Class

Function VariablePlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Dim variablePlayer : Set variablePlayer = ownProps(1)
    Select Case evt
        Case "play"
            variablePlayer.Play ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set VariablePlayerEventHandler = kwargs
    Else
        VariablePlayerEventHandler = kwargs
    End If
    
End Function