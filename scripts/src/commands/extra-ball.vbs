
Class GlfExtraBall

    Private m_name
    private m_command_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_award_events
    Private m_max_per_game

    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
    Public Property Get GetValue(value)
        Select Case value
            'Case "":
            '    GetValue = 
        End Select
    End Property

    Public Property Let AwardEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_award_events.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Let MaxPerGame(value) : Set m_max_per_game = CreateGlfInput(value) : End Property

	Public default Function init(name, mode)
        m_name = "extra_ball_" & name
        m_command_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        Set m_award_events = CreateObject("Scripting.Dictionary")
        Set m_max_per_game = CreateGlfInput(0)

        Glf_SetInitialPlayerVar m_name & "_awarded", 0

        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "extra_ball", Me)
        
        Set Init = Me
	End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_award_events.Keys
            AddPinEventListener m_award_events(evt).EventName, m_name & "_" & evt & "_award", "ExtraBallsHandler", m_priority+m_award_events(evt).Priority, Array("award", Me, m_award_events(evt))
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_award_events.Keys
            RemovePinEventListener m_award_events(evt).EventName, m_name & "_" & evt & "_award"
        Next
    End Sub

    Public Sub Award(evt)
        If evt.Evaluate() Then
            If GetPlayerState(m_name & "_awarded") < m_max_per_game.Value() Then
                SetPlayerState "extra_balls", GetPlayerState("extra_balls") + 1
                SetPlayerState m_name & "_awarded", GetPlayerState(m_name & "_awarded") + 1
                DispatchPinEvent m_name & "_awarded", Null
                DispatchPinEvent "extra_ball_awarded", Null
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function ExtraBallsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim extra_ball : Set extra_ball = ownProps(1)
    Select Case evt
        Case "award"
            extra_ball.Award ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set ExtraBallsHandler = kwargs
    Else
        ExtraBallsHandler = kwargs
    End If
End Function