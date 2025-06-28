Function CreateGlfStanduptarget(name)
	Dim standuptarget : Set standuptarget = (new GlfStandupTarget)(name)
	Set CreateGlfStanduptarget = standuptarget
End Function

Class GlfStandupTarget

    Private m_name
	Private m_switch
    Private m_use_roth
    Private m_roth_array_index
    
    Private m_debug

    Public Property Get Name()
        Name = Replace(m_name, "standup_target_", "")
    End Property
	Public Property Let Switch(value)
		m_switch = value
	End Property
    Public Property Get Switch()
        Switch = m_switch
    End Property
    Public Property Get UseRothStanduptarget()
        UseRothStanduptarget = m_use_roth
    End Property
    Public Property Let UseRothStanduptarget(value)
        m_use_roth = value
    End Property
    Public Property Get RothSTSwitchID()
        RothSTSwitchID = m_roth_array_index
    End Property
    Public Property Let RothSTSwitchID(value)
        m_roth_array_index = value
    End Property
    
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "standup_target_" & name
		m_switch = Empty
		m_debug = False
        m_use_roth = False
        m_roth_array_index = -1
        glf_standup_targets.Add name, Me
        Set Init = Me
	End Function
 
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class