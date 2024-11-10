Class GlfDropTarget
	Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped
    Private m_reset_events
  
	Public Property Get Primary(): Set Primary = m_primary: End Property
	Public Property Let Primary(input): Set m_primary = input: End Property
  
	Public Property Get Secondary(): Set Secondary = m_secondary: End Property
	Public Property Let Secondary(input): Set m_secondary = input: End Property
  
	Public Property Get Prim(): Set Prim = m_prim: End Property
	Public Property Let Prim(input): Set m_prim = input: End Property
  
	Public Property Get Sw(): Sw = m_sw: End Property
	Public Property Let Sw(input): m_sw = input: End Property
  
	Public Property Get Animate(): Animate = m_animate: End Property
	Public Property Let Animate(input): m_animate = input: End Property
  
	Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
	Public Property Let IsDropped(input): m_isDropped = input: End Property
  
	Public default Function init(primary, secondary, prim, sw, animate, isDropped, reset_events)
	  Set m_primary = primary
	  Set m_secondary = secondary
	  Set m_prim = prim
	  m_sw = sw
	  m_animate = animate
	  m_isDropped = isDropped
      m_reset_events = reset_events
	  If Not IsNull(reset_events) Then
	  	Dim evt
		For Each evt in reset_events
			AddPinEventListener evt, primary.name & "_droptarget_reset", "DropTargetEventHandler", 1000, Array("droptarget_reset", m_sw)
		Next      	
	  End If
	  Set Init = Me
	End Function
End Class

Function DropTargetEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim kwargs : kwargs = args(1)
    Dim evt : evt = ownProps(0)
    Dim switch : switch = ownProps(1)
    Select Case evt
        Case "droptarget_reset"
            DTRaise switch
    End Select
    DropTargetEventHandler = kwargs
End Function