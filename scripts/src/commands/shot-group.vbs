Class GlfShotGroup
    Private m_name
    Private m_mode
    Private m_priority
    private m_base_device
    private m_shots

    Public Property Get Name(): Name = m_name: End Property

    Public Property Let Shots(value)
        m_shots = value
        Dim shot_name
        For Each shot_name in m_shots
            AddPinEventListener "shot_" & shot_name & "_hit", m_name & "_" & m_mode & "_hit", "ShotGroupEventHandler", m_priority, Array("hit", Me)
        Next
    End Property

	Public default Function init(name, mode)
        m_name = "shot_group_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot_group", Me)

        Set Init = Me
	End Function

    Public Sub Enable()

    End Sub

    Public Sub Disable()

    End Sub

End Class

Function ShotGroupEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : Set kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Select Case evt
        Case "hit"
            DispatchPinEvent device.Name & "_hit", Null
            DispatchPinEvent device.Name & "_" & kwargs("state") & "_hit", Null
    End Select
    Set ShotGroupEventHandler = kwargs
End Function