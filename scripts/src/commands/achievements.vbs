Class GlfAchievements

    Private m_name
    Private m_priority
    Private m_complete_events
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
        End Select
    End Property

    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let CompleteEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_complete_events.Add x, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property

    Public default Function Init(name, mode)
        m_name = "achievements_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_complete_events = CreateObject("Scripting.Dictionary")

        Set m_base_device = (new GlfBaseModeDevice)(mode, "achievement", Me)
        glf_achievements.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim key
        For Each key in m_complete_events.Keys
            AddPinEventListener m_complete_events(key).EventName, m_name & "_complete_event_" & key, "AchievementsEventHandler", m_priority+m_complete_events(key).Priority, Array("complete", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim key
        For Each key in m_complete_events.Keys
            RemovePinEventListener m_complete_events(key).EventName, m_name & "_complete_event_" & key
        Next
    End Sub

    Public Sub Complete()
        'TODO: Implement Complete Events
    End Sub

    Private Sub Log(message)
        If m_debug Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function AchievementsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1)
    End If

    Dim evt : evt = ownProps(0)
    Dim achievement : Set achievement = ownProps(1)

    Select Case evt
        Case "complete"
            achievement.Complete
    End Select

    If IsObject(args(1)) Then
        Set AchievementsHandler = kwargs
    Else
        AchievementsHandler = kwargs
    End If
End Function
