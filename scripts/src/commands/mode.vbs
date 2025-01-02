
Class Mode

    Private m_name
    Private m_start_events
    Private m_stop_events
    private m_priority
    Private m_debug
    Private m_started

    Private m_ballsaves
    Private m_counters
    Private m_multiball_locks
    Private m_multiballs
    Private m_shots
    Private m_shot_groups
    Private m_ballholds
    Private m_timers
    Private m_lightplayer
    Private m_segment_display_player
    Private m_showplayer
    Private m_variableplayer
    Private m_eventplayer
    Private m_random_event_player
    Private m_shot_profiles
    Private m_sequence_shots
    Private m_state_machines

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Get Status()
        If m_started Then
            Status = "started"
        Else
            Status = "stopped"
        End If
    End Property
    Public Property Get LightPlayer()
        If IsNull(m_lightplayer) Then
            Set m_lightplayer = (new GlfLightPlayer)(Me)
        End If
        Set LightPlayer = m_lightplayer
    End Property
    Public Property Get ShowPlayer()
        If IsNull(m_showplayer) Then
            Set m_showplayer = (new GlfShowPlayer)(Me)
        End If
        Set ShowPlayer = m_showplayer
    End Property
    Public Property Get SegmentDisplayPlayer()
        If IsNull(m_segment_display_player) Then
            Set m_segment_display_player = (new GlfSegmentDisplayPlayer)(Me)
        End If
        Set SegmentDisplayPlayer = m_segment_display_player
    End Property
    Public Property Get EventPlayer() : Set EventPlayer = m_eventplayer: End Property
    Public Property Get RandomEventPlayer() : Set RandomEventPlayer = m_random_event_player : End Property
    Public Property Get VariablePlayer(): Set VariablePlayer = m_variableplayer: End Property

    Public Property Get ShotProfiles(name)
        If m_shot_profiles.Exists(name) Then
            Set ShotProfiles = m_shot_profiles(name)
        Else
            Dim new_shotprofile : Set new_shotprofile = (new GlfShotProfile)(name)
            m_shot_profiles.Add name, new_shotprofile
            Glf_ShotProfiles.Add name, new_shotprofile
            Set ShotProfiles = new_shotprofile
        End If
    End Property

    Public Property Get BallSavesItems() : BallSavesItems = m_ballsaves.Items() : End Property
    Public Property Get BallSaves(name)
        If m_ballsaves.Exists(name) Then
            Set BallSaves = m_ballsaves(name)
        Else
            Dim new_ballsave : Set new_ballsave = (new BallSave)(name, Me)
            m_ballsaves.Add name, new_ballsave
            Set BallSaves = new_ballsave
        End If
    End Property

    Public Property Get TimersItems() : TimersItems = m_timers.Items() : End Property
    Public Property Get Timers(name)
        If m_timers.Exists(name) Then
            Set Timers = m_timers(name)
        Else
            Dim new_timer : Set new_timer = (new GlfTimer)(name, Me)
            m_timers.Add name, new_timer
            Set Timers = new_timer
        End If
    End Property

    Public Property Get CountersItems() : CountersItems = m_counters.Items() : End Property
    Public Property Get Counters(name)
        If m_counters.Exists(name) Then
            Set Counters = m_counters(name)
        Else
            Dim new_counter : Set new_counter = (new GlfCounter)(name, Me)
            m_counters.Add name, new_counter
            Set Counters = new_counter
        End If
    End Property

    Public Property Get MultiballLocksItems() : MultiballLocksItems = m_multiball_locks.Items() : End Property
    Public Property Get MultiballLocks(name)
        If m_multiball_locks.Exists(name) Then
            Set MultiballLocks = m_multiball_locks(name)
        Else
            Dim new_multiball_lock : Set new_multiball_lock = (new GlfMultiballLocks)(name, Me)
            m_multiball_locks.Add name, new_multiball_lock
            Set MultiballLocks = new_multiball_lock
        End If
    End Property

    Public Property Get MultiballsItems() : MultiballsItems = m_multiballs.Items() : End Property
    Public Property Get Multiballs(name)
        If m_multiballs.Exists(name) Then
            Set Multiballs = m_multiballs(name)
        Else
            Dim new_multiball : Set new_multiball = (new GlfMultiballs)(name, Me)
            m_multiballs.Add name, new_multiball
            Set Multiballs = new_multiball
        End If
    End Property

    Public Property Get SequenceShotsItems() : SequenceShotsItems = m_sequence_shots.Items() : End Property
    Public Property Get SequenceShots(name)
        If m_sequence_shots.Exists(name) Then
            Set SequenceShots = m_sequence_shots(name)
        Else
            Dim new_sequence_shot : Set new_sequence_shot = (new GlfSequenceShots)(name, Me)
            m_sequence_shots.Add name, new_sequence_shot
            Set SequenceShots = new_sequence_shot
        End If
    End Property

    Public Property Get StateMachines(name)
        If m_state_machines.Exists(name) Then
            Set StateMachines = m_state_machines(name)
        Else
            Dim new_state_machine : Set new_state_machine = (new GlfStateMachine)(name, Me)
            m_state_machines.Add name, new_state_machine
            Set StateMachines = new_state_machine
        End If
    End Property
    Public Property Get ModeStateMachines(): ModeStateMachines = m_state_machines.Items(): End Property

    Public Property Get ModeShots(): ModeShots = m_shots.Items(): End Property
    Public Property Get Shots(name)
        If m_shots.Exists(name) Then
            Set Shots = m_shots(name)
        Else
            Dim new_shot : Set new_shot = (new GlfShot)(name, Me)
            m_shots.Add name, new_shot
            Set Shots = new_shot
        End If
    End Property

    Public Property Get ShotGroupsItems() : ShotGroupsItems = m_shot_groups.Items() : End Property
    Public Property Get ShotGroups(name)
        If m_shot_groups.Exists(name) Then
            Set ShotGroups = m_shot_groups(name)
        Else
            Dim new_shot_group : Set new_shot_group = (new GlfShotGroup)(name, Me)
            m_shot_groups.Add name, new_shot_group
            Set ShotGroups = new_shot_group
        End If
    End Property

    Public Property Get BallHoldsItems() : BallHoldsItems = m_ballholds.Items() : End Property
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
    End Property

    Public Property Get IsDebug()
        If m_debug = True Then
            IsDebug = 1
        Else
            IsDebug = 0
        End If
    End Property
    Public Property Let Debug(value)
        m_debug = value
        Dim config_item
        For Each config_item in m_ballsaves.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_counters.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_timers.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_multiball_locks.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_multiballs.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_shots.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_shot_groups.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_ballholds.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_sequence_shots.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_state_machines.Items()
            config_item.Debug = value
        Next
        If Not IsNull(m_lightplayer) Then
            m_lightplayer.Debug = value
        End If
        If Not IsNull(m_eventplayer) Then
            m_eventplayer.Debug = value
        End If
        If Not IsNull(m_random_event_player) Then
            m_random_event_player.Debug = value
        End If
        If Not IsNull(m_showplayer) Then
            m_showplayer.Debug = value
        End If
        If Not IsNull(m_segment_display_player) Then
            m_segment_display_player.Debug = value
        End If
        If Not IsNull(m_variableplayer) Then
            m_variableplayer.Debug = value
        End If
        Glf_MonitorModeUpdate Me
    End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_priority = priority
        m_started = False
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
        Set m_shot_profiles = CreateObject("Scripting.Dictionary")
        Set m_sequence_shots = CreateObject("Scripting.Dictionary")
        Set m_state_machines = CreateObject("Scripting.Dictionary")
        
        m_lightplayer = Null
        m_showplayer = Null
        m_segment_display_player = Null
        Set m_eventplayer = (new GlfEventPlayer)(Me)
        Set m_random_event_player = (new GlfRandomEventPlayer)(Me)
        Set m_variableplayer = (new GlfVariablePlayer)(Me)
        Glf_MonitorModeUpdate Me
        AddPinEventListener m_name & "_starting", m_mode & "_starting_end", "ModeEventHandler", -99, Array("started", Me, "")
        AddPinEventListener m_name & "_stopping", m_mode & "_stopping_end", "ModeEventHandler", -99, Array("stopped", Me, "")
        Set Init = Me
	End Function

    Public Sub StartMode()
        Log "Starting"
        m_started=True
        DispatchQueuePinEvent m_name & "_starting", Null
    End Sub

    Public Sub StopMode()
        If m_started = True Then
            m_started = False
            Log "Stopping"
            DispatchQueuePinEvent m_name & "_stopping", Null
        End If
    End Sub

    Public Sub Started()
        DispatchPinEvent m_name & "_started", Null
        Glf_MonitorModeUpdate Me
        Log "Started"
    End Sub

    Public Sub Stopped()
        DispatchPinEvent m_name & "_stopped", Null
        Glf_MonitorModeUpdate Me
        Log "Stopped"
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        dim yaml, child
        yaml = "#config_version=6" & vbCrLf & vbCrLf

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

        yaml = yaml & "  priority: " & m_priority & vbCrLf
        
        If UBound(m_ballsaves.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "ball_saves: " & vbCrLf
            For Each child in m_ballsaves.Keys
                yaml = yaml & m_ballsaves(child).ToYaml
            Next
        End If
        
        If UBound(m_shot_profiles.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "shot_profiles: " & vbCrLf
            For Each child in m_shot_profiles.Keys
                yaml = yaml & m_shot_profiles(child).ToYaml
            Next
        End If
        
        If UBound(m_shots.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "shots: " & vbCrLf
            For Each child in m_shots.Keys
                yaml = yaml & m_shots(child).ToYaml
            Next
        End If
        
        If UBound(m_shot_groups.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "shot_groups: " & vbCrLf
            For Each child in m_shot_groups.Keys
                yaml = yaml & m_shot_groups(child).ToYaml
            Next
        End If
        
        If UBound(m_eventplayer.Events.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "event_player: " & vbCrLf
            yaml = yaml & m_eventplayer.ToYaml()
        End If
        
        If Not IsNull(m_showPlayer) Then
            If UBound(m_showplayer.EventShows)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "show_player: " & vbCrLf
                yaml = yaml & m_showplayer.ToYaml()
            End If
        End If
        
        If Not IsNull(m_lightplayer) Then
            If UBound(m_lightplayer.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "light_player: " & vbCrLf
                For Each child in m_lightplayer.EventNames
                    yaml = yaml & m_lightplayer.ToYaml()
                Next
            End If
        End If

        If Not IsNull(m_segment_display_player) Then
            If UBound(m_segment_display_player.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "segment_display_player: " & vbCrLf
                For Each child in m_segment_display_player.EventNames
                    yaml = yaml & m_segment_display_player.ToYaml()
                Next
            End If
        End If
        
        If UBound(m_ballholds.Keys)>-1 Then
            yaml = yaml & vbCrLf
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
            mode.Started
        Case "stopped"
            mode.Stopped
    End Select
    If IsObject(args(1)) Then
        Set ModeEventHandler = kwargs
    Else
        ModeEventHandler = kwargs
    End If
End Function