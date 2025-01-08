
Class GlfMultiballLocks

    Private m_name
    Private m_lock_device
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_enable_events
    Private m_disable_events
    Private m_balls_to_lock
    Private m_balls_locked
    Private m_balls_to_replace
    Private m_lock_events
    Private m_reset_events
    Private m_enabled
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "locked_balls":
                GetValue = m_balls_locked
        End Select
    End Property
    Public Property Get LockDevice() : LockDevice = m_lock_device : End Property
    Public Property Let LockDevice(value) : m_lock_device = value : End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let BallsToLock(value) : m_balls_to_lock = value : End Property
    Public Property Let LockEvents(value) : m_lock_events = value : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let BallsToReplace(value) : m_balls_to_replace = value : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(name, mode)
        m_name = "multiball_lock_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_lock_events = Array()
        m_reset_events = Array()
        m_lock_device = Empty
        m_balls_to_lock = 0
        m_balls_to_replace = -1
        m_enabled = False
        m_balls_locked = 0
        Set m_base_device = (new GlfBaseModeDevice)(mode, "multiball_lock", Me)
        glf_multiball_locks.Add name, Me
        Set Init = Me
	End Function

    Public Sub Activate()
        If UBound(m_base_device.EnableEvents.Keys()) = -1 Then
            Enable()
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        If Not IsEmpty(m_lock_device) Then
            AddPinEventListener "balldevice_" & m_lock_device & "_ball_enter", m_mode & "_" & name & "_lock", "MultiballLocksHandler", m_priority, Array("lock", me, m_lock_device)
        End If
        Dim evt
        For Each evt in m_lock_events
            AddPinEventListener evt, m_name & "_ball_locked", "MultiballLocksHandler", m_priority, Array("virtual_lock", Me, Null)
        Next
        For Each evt in m_reset_events
            AddPinEventListener evt, m_name & "_reset", "MultiballLocksHandler", m_priority, Array("reset", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        If Not IsEmpty(m_lock_device) Then
            RemovePinEventListener "balldevice_" & m_lock_device & "_ball_enter", m_mode & "_" & name & "_lock"
        End If
        Dim evt
        For Each evt in m_lock_events
            RemovePinEventListener evt, m_name & "_ball_locked"
        Next
        For Each evt in m_reset_events
            RemovePinEventListener evt, m_name & "_reset"
        Next
    End Sub

    Public Function Lock(device, unclaimed_balls)
        
        If unclaimed_balls <= 0 Then
            Lock = unclaimed_balls
            Exit Function
        End If
        
        Dim balls_locked
        If GetPlayerState(m_name & "_balls_locked") = False Then
            balls_locked = 1
        Else
            balls_locked = GetPlayerState(m_name & "_balls_locked") + 1
        End If
        If balls_locked > m_balls_to_lock Then
            Log "Cannot lock balls. Lock is full."
            Lock = unclaimed_balls
            Exit Function
        End If

        SetPlayerState m_name & "_balls_locked", balls_locked
        

        If Not IsNull(device) Then
            
            If glf_ball_devices(device).Balls() > balls_locked Then
                glf_ball_devices(device).Eject()
            Else
                If m_balls_to_replace = -1 Or balls_locked <= m_balls_to_replace Then
                    ' glf_BIP = glf_BIP - 1
                    SetDelay m_name & "_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", Me),Null), 1000
                End If
            End If
        End If

        DispatchPinEvent m_name & "_locked_ball", balls_locked
        
        If balls_locked = m_balls_to_lock Then
            DispatchPinEvent m_name & "_full", balls_locked
        End If

        Lock = unclaimed_balls - 1
    End Function

    Public Sub Reset
        Log "Resetting multiball lock count"
        SetPlayerState m_name & "_balls_locked", 0
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function MultiballLocksHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    Select Case evt
        Case "activate"
            multiball.Activate
        Case "deactivate"
            multiball.Deactivate
        Case "enable"
            multiball.Enable
        Case "disable"
            multiball.Disable
        Case "lock"
            kwargs = multiball.Lock(ownProps(2), kwargs)
        Case "virtual_lock"
            multiball.Lock Null, 1
        Case "reset"
            multiball.Reset
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True And glf_plunger.IncomingBalls = 0  Then
                Glf_ReleaseBall(Null)
                SetDelay multiball.Name&"_auto_launch", "MultiballLocksHandler" , Array(Array("auto_launch", multiball),Null), 500
            Else
                SetDelay multiball.Name&"_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", multiball), Null), 1000
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay multiball.Name&"_auto_launch", "MultiballLocksHandler" , Array(Array("auto_launch", multiball), Null), 500
            End If
    End Select
    If IsObject(args(1)) Then
        Set MultiballLocksHandler = kwargs
    Else
        MultiballLocksHandler = kwargs
    End If
End Function