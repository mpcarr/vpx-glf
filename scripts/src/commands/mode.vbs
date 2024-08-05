
Class Mode

    Private m_name
    Private m_start_events
    Private m_stop_events
    private m_priority
    Private m_debug

    Private m_ballsaves
    Private m_counters
    Private m_multiball_locks
    Private m_multiballs
    Private m_shots
    Private m_shot_groups
    Private m_ballholds
    Private m_timers
    Private m_lightplayer
    Private m_showplayer
    Private m_variableplayer
    Private m_eventplayer

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Get Debug(): Debug = m_debug: End Property
    Public Property Get LightPlayer(): Set LightPlayer = m_lightplayer: End Property
    Public Property Get ShowPlayer(): Set ShowPlayer = m_showplayer: End Property
    Public Property Get EventPlayer() : Set EventPlayer = m_eventplayer: End Property
    Public Property Get VariablePlayer(): Set VariablePlayer = m_variableplayer: End Property
    

    Public Property Get BallSaves(name)
        If m_ballsaves.Exists(name) Then
            Set BallSaves = m_ballsaves(name)
        Else
            Dim new_ballsave : Set new_ballsave = (new BallSave)(name, Me)
            m_ballsaves.Add name, new_ballsave
            Set BallSaves = new_ballsave
        End If
    End Property

    Public Property Get Timers(name)
        If m_timers.Exists(name) Then
            Set Timers = m_timers(name)
        Else
            Dim new_timer : Set new_timer = (new GlfTimer)(name, Me)
            m_timers.Add name, new_timer
            Set Timers = new_timer
        End If
    End Property

    Public Property Get Counters(name)
        If m_counters.Exists(name) Then
            Set Counters = m_counters(name)
        Else
            Dim new_counter : Set new_counter = (new GlfCounter)(name, Me)
            m_counters.Add name, new_counter
            Set Counters = new_counter
        End If
    End Property

    Public Property Get MultiballLocks(name)
        If m_multiball_locks.Exists(name) Then
            Set MultiballLocks = m_multiball_locks(name)
        Else
            Dim new_multiball_lock : Set new_multiball_lock = (new GlfMultiballLocks)(name, Me)
            m_multiball_locks.Add name, new_multiball_lock
            Set MultiballLocks = new_multiball_lock
        End If
    End Property

    Public Property Get Multiballs(name)
        If m_multiballs.Exists(name) Then
            Set Multiballs = m_multiballs(name)
        Else
            Dim new_multiball : Set new_multiball = (new GlfMultiballs)(name, Me)
            m_multiballs.Add name, new_multiball
            Set Multiballs = new_multiball
        End If
    End Property

    Public Property Get Shots(name)
        If m_shots.Exists(name) Then
            Set Shots = m_shots(name)
        Else
            Dim new_shot : Set new_shot = (new GlfShot)(name, Me)
            m_shots.Add name, new_shot
            Set Shots = new_shot
        End If
    End Property

    Public Property Get ShotGroups(name)
        If m_shot_groups.Exists(name) Then
            Set ShotGroups = m_shot_groups(name)
        Else
            Dim new_shot_group : Set new_shot_group = (new GlfShotGroup)(name, Me)
            m_shot_groups.Add name, new_shot_group
            Set ShotGroups = new_shot_group
        End If
    End Property

    Public Property Get BallHolds(name)
        If m_ballholds.Exists(name) Then
            Set BallHolds = m_shots(name)
        Else
            Dim new_ballhold : Set new_ballhold = (new GlfBallHold)(name, Me)
            m_ballholds.Add name, new_ballhold
            Set BallHolds = new_ballhold
        End If
    End Property

    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Name, newEvent
            AddPinEventListener newEvent.EventName, m_name & "_start", "ModeEventHandler", m_priority, Array("start", Me, newEvent)
        Next
    End Property
    
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Name, newEvent
            AddPinEventListener newEvent.EventName, m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me, newEvent)
        Next
        Set newEvent = (new GlfEvent)("ball_ended")
        AddPinEventListener newEvent.EventName, m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me, newEvent)
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_priority = priority
        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")
        Set m_ballsaves = CreateObject("Scripting.Dictionary")
        Set m_counters = CreateObject("Scripting.Dictionary")
        Set m_timers = CreateObject("Scripting.Dictionary")
        Set m_multiball_locks = CreateObject("Scripting.Dictionary")
        Set m_multiballs = CreateObject("Scripting.Dictionary")
        Set m_shots = CreateObject("Scripting.Dictionary")
        Set m_shot_groups = CreateObject("Scripting.Dictionary")
        Set m_ballholds = CreateObject("Scripting.Dictionary")
        Set m_lightplayer = (new GlfLightPlayer)(Me)
        Set m_showplayer = (new GlfShowPlayer)(Me)
        Set m_eventplayer = (new GlfEventPlayer)(Me)
        Set m_variableplayer = (new GlfVariablePlayer)(Me)
        Set Init = Me
	End Function

    Public Sub StartMode()
        Log "Starting"
        DispatchPinEvent m_name & "_starting", Null
        DispatchPinEvent m_name & "_started", Null
        Log "Started"
    End Sub

    Public Sub StopMode()
        Log "Stopping"
        DispatchPinEvent m_name & "_stopping", Null
        DispatchPinEvent m_name & "_stopped", Null
        Log "Stopped"
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        dim yaml, child
        yaml = "#config_version=6" & vbCrLf
        yaml = yaml & "mode:" & vbCrLf

        If UBound(m_start_events.Keys) > -1 Then
            yaml = yaml & "  start_events: "
            x=0
            For Each key in m_start_events.keys
                yaml = yaml & m_start_events(key).Raw
                If x <> UBound(m_start_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_stop_events.Keys) > -1 Then
            yaml = yaml & "  stop_events: "
            x=0
            For Each key in m_stop_events.keys
                yaml = yaml & m_stop_events(key).Raw
                If x <> UBound(m_stop_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        yaml = yaml & "  priority: " & m_priority
        yaml = yaml & vbCrLf
        If UBound(m_ballsaves.Keys)>-1 Then
            yaml = yaml & "ballsaves: " & vbCrLf
            For Each child in m_ballsaves.Keys
                yaml = yaml & m_ballsaves(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        If UBound(m_shots.Keys)>-1 Then
            yaml = yaml & "shots: " & vbCrLf
            For Each child in m_shots.Keys
                yaml = yaml & m_shots(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        If UBound(m_shot_groups.Keys)>-1 Then
            yaml = yaml & "shot_groups: " & vbCrLf
            For Each child in m_shot_groups.Keys
                yaml = yaml & m_shot_groups(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        If UBound(m_eventplayer.Events.Keys)>-1 Then
            yaml = yaml & "event_player: " & vbCrLf
            yaml = yaml & m_eventplayer.ToYaml()
        End If
        yaml = yaml & vbCrLf
        If UBound(m_ballholds.Keys)>-1 Then
            yaml = yaml & "ball_holds: " & vbCrLf
            For Each child in m_ballholds.Keys
                yaml = yaml & m_ballholds(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        
        Dim fso, modesFolder, TxtFileStream
        Set fso = CreateObject("Scripting.FileSystemObject")
        modesFolder = "glf_mpf\modes\" & Replace(m_name, "mode_", "") & "\config"

        If Not fso.FolderExists("glf_mpf") Then
            fso.CreateFolder "glf_mpf"
        End If

        Dim currentFolder
        Dim folderParts
        Dim i
    
        ' Split the path into parts
        folderParts = Split(modesFolder, "\")
        
        ' Initialize the current folder as the root
        currentFolder = folderParts(0)
    
        ' Iterate over each part of the path and create folders as needed
        For i = 1 To UBound(folderParts)
            currentFolder = currentFolder & "\" & folderParts(i)
            If Not fso.FolderExists(currentFolder) Then
                fso.CreateFolder(currentFolder)
            End If
        Next


        
        Set TxtFileStream = fso.OpenTextFile(modesFolder & "\" & Replace(m_name, "mode_", "") & ".yaml", 2, True)
        TxtFileStream.WriteLine yaml
        TxtFileStream.Close

        ToYaml = yaml
    End Function
End Class

Function ModeEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim mode : Set mode = ownProps(1)
    Dim glfEvent
    Select Case evt
        Case "start"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            mode.StartMode
        Case "stop"
            Set glfEvent = ownProps(2)
            If Not IsNull(glfEvent.Condition) Then
                If GetRef(glfEvent.Condition)() = False Then
                    Exit Function
                End If
            End If
            mode.StopMode
        Case "started"
            DispatchPinEvent mode.Name & "_started", Null
        Case "stopped"
            DispatchPinEvent mode.Name & "_stopped", Null
    End Select
    If IsObject(args(1)) Then
        Set ModeEventHandler = kwargs
    Else
        ModeEventHandler = kwargs
    End If
End Function