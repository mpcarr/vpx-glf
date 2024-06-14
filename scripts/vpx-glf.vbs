'VPX Game Logic Framework (https://mpcarr.github.io/vpx-gle-framework/)

'
Dim canAddPlayers : canAddPlayers = True
Dim currentPlayer : currentPlayer = Null
Dim glf_PI : glf_PI = 4 * Atn(1)
Dim glf_plunger
Dim gameStarted : gameStarted = False
Dim pinEvents : Set pinEvents = CreateObject("Scripting.Dictionary")
Dim pinEventsOrder : Set pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerEvents : Set playerEvents = CreateObject("Scripting.Dictionary")
Dim playerEventsOrder : Set playerEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")
Dim bcpController : bcpController = Null
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = ""
Dim lightCtrl : Set lightCtrl = new LStateController
Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim glf_BIP : glf_BIP = 0

Dim glf_ballsPerGame : glf_ballsPerGame = 3
Dim glf_troughSize : glf_troughSize = 8

Dim debugLog : Set debugLog = (new DebugLogFile)()
Dim debugEnabled : debugEnabled = True

Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController
    Set bcpController = (new GlfVpxBcpController)(bcpPort	, bcpExeName)
	vpmMapLights alights
	lightCtrl.RegisterLights "VPX"
End Sub

Public Sub Glf_Init()
	If useBCP = True Then
		ConnectToBCPMediaController
	End If
	Dim switch, switchHitSubs
	switchHitSubs = ""
	For Each switch in Glf_Switches
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_Hit() : DispatchPinEvent """ & switch.Name & "_active"", ActiveBall : End Sub" & vbCrLf
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_UnHit() : DispatchPinEvent """ & switch.Name & "_inactive"", ActiveBall : End Sub" & vbCrLf
	Next
	ExecuteGlobal switchHitSubs
End Sub

Sub Glf_Options(ByVal eventId)

	glf_troughSize = Table1.Option("Trough Capacity", 1, 8, 1, 6, 0)
	If gameStarted = False Then
		swTrough1.DestroyBall
		swTrough2.DestroyBall
		swTrough3.DestroyBall
		swTrough4.DestroyBall
		swTrough5.DestroyBall
		swTrough6.DestroyBall
		swTrough7.DestroyBall
		swTrough8.DestroyBall
		If glf_troughSize > 0 Then : Set glf_ball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 1 Then : Set glf_ball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 2 Then : Set glf_ball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 3 Then : Set glf_ball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 4 Then : Set glf_ball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 5 Then : Set glf_ball6 = swTrough6.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 6 Then : Set glf_ball7 = swTrough7.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
		If glf_troughSize > 7 Then : Set glf_ball8 = swTrough8.CreateSizedballWithMass(Ballsize / 2,Ballmass) : End If
	End If
	Dim ballsPerGame : ballsPerGame = Table1.Option("Balls Per Game", 1, 2, 1, 1, 0, Array("3 Balls", "5 Balls"))
	If ballsPerGame = 1 Then
		glf_ballsPerGame = 3
	Else
		glf_ballsPerGame = 5
	End If
End Sub

Public Sub Glf_Exit()
	If Not IsNull(bcpController) Then
		bcpController.Disconnect
		Set bcpController = Nothing
	End If
End Sub

Public Sub Glf_KeyDown(ByVal keycode)
    If gameStarted = True Then
		If keycode = LeftFlipperKey Then
			'DispatchPinEvent GLF_SWITCH_LEFT_FLIPPER_DOWN, Null
		End If
		
		If keycode = RightFlipperKey Then
			'DispatchPinEvent GLF_SWITCH_RIGHT_FLIPPER_DOWN, Null
		End If
		
		If keycode = StartGameKey Then
			If canAddPlayers = True Then
				Glf_AddPlayer()
			End If
		End If
	Else
		If keycode = StartGameKey Then
			Glf_AddPlayer()
			Glf_StartGame()
		End If
	End If
End Sub

Public Sub Glf_KeyUp(ByVal keycode)
	If gameStarted = True Then
		If KeyCode = PlungerKey Then
			Plunger.Fire
		End If
	End If
End Sub

Public Sub Glf_EventTimer_Timer()
	DelayTick
End Sub

'******************************************************
'*****   GLF Pin Events                             ****
'******************************************************

Const GLF_GAME_STARTED = "game_started"
Const GLF_GAME_OVER = "game_ended"
Const GLF_BALL_ENDED = "ball_ended"
Const GLF_NEXT_PLAYER = "next_player"
Const GLF_BALL_DRAIN = "ball_drain"
Const GLF_BALL_STARTED = "ball_started"

'******************************************************
'*****   GLF Player State                          ****
'******************************************************

Const GLF_SCORE = "score"
Const GLF_CURRENT_BALL = "current_ball"
Const GLF_INITIALS = "initials"




'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************

Class GlfVpxBcpController

    Private m_bcpController, m_connected

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_server.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        bcpUpdate.Enabled = True
        If Err Then Debug.print("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

	Public Sub Send(commandMessage)
		If m_connected Then
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
        End If
	End Sub
    
    Public Sub PlaySlide(slide, context, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""" & slide & """: {""action"": ""play"", ""expire"": 0}}, ""context"": """ & context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub SendPlayerVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & GetCurrentPlayerNumber
            '06:34:34.644 : VERBOSE : BCP : Received BCP command: ball_start?player_num=int:1&ball=int:1
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
                retValue = "string:" & value
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
            bcpUpdate.Enabled = False
        End If
    End Sub
End Class

Sub Glf_BcpSendPlayerVar(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim player_var : player_var = kwargs(0)
    Dim value : value = kwargs(1)
    Dim prevValue : prevValue = kwargs(2)
    bcpController.SendPlayerVariable player_var, value, prevValue
End Sub

Sub Glf_BcpAddPlayer(playerNum)
    If useBcp Then
        bcpController.Send("player_added?player_num=int:"&playerNum)
    End If
End Sub

Sub Glf_BcpUpdate_Timer()
    Dim messages : messages = bcpController.GetMessages()
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter
        For Each message in messages
            Select Case message.Command
                case "hello"
                    bcpController.Reset
                case "monitor_start"
                    Dim category : category = message.GetValue("category")
                    If category = "player_vars" Then
                        AddPlayerStateEventListener "score", "bcp_player_var_score", "BcpSendPlayerVar", 1000, Null
                        AddPlayerStateEventListener "current_ball", "bcp_player_var_ball", "BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    Dim eventName : eventName = message.GetValue("event")
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************
Class BallSave

    Private m_name
    Private m_mode
    Private m_active_time
    Private m_grace_period
    Private m_enable_events
    Private m_timer_start_events
    Private m_auto_launch
    Private m_balls_to_save
    Private m_saving_balls
    Private m_enabled
    Private m_timer_started
    Private m_tick
    Private m_in_grace
    Private m_in_hurry_up
    Private m_hurry_up_time
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AutoLaunch(): AutoLaunch = m_auto_launch: End Property
    Public Property Let ActiveTime(value) : m_active_time = value*1000 : End Property
    Public Property Let GracePeriod(value) : m_grace_period = value*1000 : End Property
    Public Property Let HurryUpTime(value) : m_hurry_up_time = value*1000 : End Property
    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let TimerStartEvents(value) : m_timer_start_events = value : End Property
    Public Property Let AutoLaunch(value) : m_auto_launch = value : End Property
    Public Property Let BallsToSave(value) : m_balls_to_save = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "ball_saves_" & name
        m_mode = mode.Name
        m_active_time = 0
	    m_grace_period = 0
        m_hurry_up_time = 0
        m_enable_events = Array()
        m_timer_start_events = Array()
	    m_auto_launch = False
	    m_balls_to_save = 1
        m_enabled = False
        m_timer_started = False
        m_debug = False
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "BallSaveEventHandler", mode.Priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "BallSaveEventHandler", mode.Priority, Array("deactivate", Me)
	  Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "BallSaveEventHandler", 1000, Array("enable", Me)
        Next
        For Each evt in m_timer_start_events
            AddPinEventListener evt, m_name & "_timer_start", "BallSaveEventHandler", 1000, Array("timer_start", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
        For Each evt in m_timer_start_events
            RemovePinEventListener evt, m_name & "_timer_start"
        Next
    End Sub

    Public Sub Enable
        If m_enabled = True Then
            Exit Sub
        End If
        m_enabled = True
        m_saving_balls = m_balls_to_save
        Log "Enabling. Auto launch: "&m_auto_launch&", Balls to save: "&m_balls_to_save&", Active time: "& m_active_time&"ms"
        AddPinEventListener "ball_drain", m_name & "_ball_drain", "BallSaveEventHandler", 1000, Array("drain", Me)
        DispatchPinEvent m_name&"_enabled", Null
        If UBound(m_timer_start_events) = -1 Then
            Log "Timer Starting as no timer start events are set"
            TimerStart()
        End If
    End Sub

    Public Sub Disable
        'Disable ball save
        If m_enabled = False Then
            Exit Sub
        End If
        m_enabled = False
        m_saving_balls = m_balls_to_save
        m_timer_started = False
        Log "Disabling..."
        RemovePinEventListener "ball_drain", m_name & "_ball_drain"
        RemoveDelay "_ball_saves_"&m_name&"_disable"
        RemoveDelay m_name&"_grace_period"
        RemoveDelay m_name&"_hurry_up_time"
            
    End Sub

    Sub Drain(ballsToSave)
        If m_enabled = True And ballsToSave > 0 Then
            If m_saving_balls > 0 Then
                m_saving_balls = m_saving_balls -1
            End If
            Log "Ball(s) drained while active. Requesting new one(s). Auto launch: "& m_auto_launch
            DispatchPinEvent m_name&"_saving_ball", Null
            SetDelay m_name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", Me),Null), 1000
            If m_saving_balls = 0 Then
                Disable()
            End If
        End If
    End Sub

    Public Sub TimerStart
        'Start the timer.
        'This is usually called after the ball was ejected while the ball save may have been enabled earlier.
        If m_timer_started=True Or m_enabled=False Then
            Exit Sub
        End If
        m_timer_started=True
        DispatchPinEvent m_name&"_timer_start", Null
        If m_active_time > 0 Then
            Log "Starting ball save timer: " & m_active_time
            Log "gametime: "& gametime & ". disabled at: " & gametime+m_active_time+m_grace_period
            SetDelay m_name&"_disable", "BallSaveEventHandler" , Array(Array("disable", Me),Null), m_active_time+m_grace_period
            SetDelay m_name&"_grace_period", "BallSaveEventHandler", Array(Array("grace_period", Me),Null), m_active_time
            SetDelay m_name&"_hurry_up_time", "BallSaveEventHandler", Array(Array("hurry_up_time", Me), Null), m_active_time-m_hurry_up_time
        End If
    End Sub

    Public Sub EnterGracePeriod
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Sub EnterHurryUpTime
        DispatchPinEvent m_name & "_hurry_up_time", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallSaveEventHandler(args)
    Dim ownProps, ballsToSave : ownProps = args(0) : ballsToSave = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballSave : Set ballSave = ownProps(1)
    Select Case evt
        Case "activate"
            ballSave.Activate
        Case "deactivate"
            ballSave.Deactivate
        Case "enable"
            ballSave.Enable
        Case "disable"
            ballSave.Disable
        Case "grace_period"
            ballSave.EnterGracePeriod
        Case "hurry_up_time"
            ballSave.EnterHurryUpTime
        Case "drain"
            If ballsToSave > 0 Then
                ballSave.Drain ballsToSave
                ballsToSave = ballsToSave - 1
            End If
        Case "timer_start"
            ballSave.TimerStart
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True Then
                Glf_ReleaseBall(Null)
                If ballSave.AutoLaunch = True Then
                    SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave),Null), 500
                End If
            Else
                SetDelay ballSave.Name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", ballSave), Null), 1000
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave), Null), 500
            End If
    End Select
    BallSaveEventHandler = ballsToSave
End Function

Class Counter

    Private m_name
    Private m_priority
    Private m_mode
    Private m_enable_events
    Private m_count_events
    Private m_count_complete_value
    Private m_disable_on_complete
    Private m_reset_on_complete
    Private m_events_when_complete
    Private m_persist_state
    Private m_debug

    Private m_count

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let CountEvents(value) : m_count_events = value : End Property
    Public Property Let CountCompleteValue(value) : m_count_complete_value = value : End Property
    Public Property Let DisableOnComplete(value) : m_disable_on_complete = value : End Property
    Public Property Let ResetOnComplete(value) : m_reset_on_complete = value : End Property
    Public Property Let EventsWhenComplete(value) : m_events_when_complete = value : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "counter_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_count = -1

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "CounterEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "CounterEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub SetValue(value)
        If value = "" Then
            value = 0
        End If
        m_count = value
        If m_persist_state Then
            SetPlayerState m_name & "_state", m_count
        End If
    End Sub

    Public Sub Activate()
        If m_persist_state And m_count > -1 Then
            If Not IsNull(GetPlayerState(m_name & "_state")) Then
                SetValue GetPlayerState(m_name & "_state")
            Else
                SetValue 0
            End If
        Else
            SetValue 0
        End If
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "CounterEventHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        If Not m_persist_state Then
            SetValue -1
        End If
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_count_events
            AddPinEventListener evt, m_name & "_count", "CounterEventHandler", m_priority, Array("count", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_count_events
            RemovePinEventListener evt, m_name & "_count"
        Next
    End Sub

    Public Sub Count()
        Log "counting: old value: "& m_count & ", new Value: " & m_count+1 & ", target: "& m_count_complete_value
        SetValue m_count + 1
        If m_count = m_count_complete_value Then
            Dim evt
            For Each evt in m_events_when_complete
                DispatchPinEvent evt, Null
            Next
            If m_disable_on_complete Then
                Disable()
            End If
            If m_reset_on_complete Then
                SetValue 0
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function CounterEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim counter : Set counter = ownProps(1)
    Select Case evt
        Case "activate"
            counter.Activate
        Case "deactivate"
            counter.Deactivate
        Case "enable"
            counter.Enable
        Case "count"
            counter.Count
    End Select
    CounterEventHandler = kwargs
End Function


Class EventPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Let Events(value) : Set m_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", "event_player_activate", "EventPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "event_player_deactivate", "EventPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_event_player_play", "EventPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_event_player_play"
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_mode & "_event_player_play", message
        End If
    End Sub
End Class

Function EventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            eventPlayer.Activate
        Case "deactivate"
            eventPlayer.Deactivate
        Case "play"
            dim evtToFire
            For Each evtToFire in ownProps(2)
                DispatchPinEvent evtToFire, Null
            Next
    End Select
    EventPlayerEventHandler = kwargs
End Function

Class LightPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Let Events(value) : Set m_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", "light_player_activate", "LightPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "light_player_deactivate", "LightPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_light_player_play"
        Next
    End Sub

    Public Sub Play(lights)
        Dim light
        For Each light in lights
            If light(1) = "off" Then
                lightCtrl.LightOff light(0)
            Else
                If UBound(light) = 2 Then
                    lightCtrl.LightOn light(0)
                    lightCtrl.FadeLightToColor light(0), light(1), light(2)
                Else
                    lightCtrl.LightOnWithColor light(0), light(1)
                End If
            End If
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_mode & "_light_player", message
        End If
    End Sub
End Class

Function LightPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim LightPlayer : Set LightPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            LightPlayer.Activate
        Case "deactivate"
            LightPlayer.Deactivate
        Case "play"
            LightPlayer.Play ownProps(2)
    End Select
    LightPlayerEventHandler = Null
End Function



Class Mode

    Private m_name
    Private m_start_events
    Private m_stop_events
    private m_priority
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Priority(): Priority = m_priority: End Property

    Public Property Let StartEvents(value)
        m_start_events = value
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_start", "ModeEventHandler", m_priority, Array("start", Me)
        Next
    End Property
    
    Public Property Let StopEvents(value)
        m_stop_events = value
        Dim evt
        For Each evt in m_stop_events
            AddPinEventListener evt, m_name & "_stop", "ModeEventHandler", m_priority, Array("stop", Me)
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_priority = priority
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
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function ModeEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim mode : Set mode = ownProps(1)
    Select Case evt
        Case "start"
            mode.StartMode
        Case "stop"
            mode.StopMode
    End Select
    ModeEventHandler = kwargs
End Function

Class MultiballLocks

    Private m_name
    Private m_priority
    Private m_mode
    Private m_enable_events
    Private m_disable_events
    Private m_balls_to_lock
    Private m_balls_locked
    Private m_lock_events
    Private m_reset_events
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let BallsToLock(value) : m_balls_to_lock = value : End Property
    Public Property Let LockEvents(value) : m_lock_events = value : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "multiball_locks_" & name
        m_mode = mode.Name
        m_lock_events = Array()
        m_reset_events = Array()
        m_balls_to_lock = 0

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "MultiballLocksHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "MultiballLocksHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MultiballLocksHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_lock_events
            AddPinEventListener evt, m_name & "_ball_locked", "MultiballLocksHandler", m_priority, Array("lock", Me)
        Next
        For Each evt in m_reset_events
            AddPinEventListener evt, m_name & "_reset", "MultiballLocksHandler", m_priority, Array("reset", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_lock_events
            RemovePinEventListener evt, m_name & "_ball_locked"
        Next
    End Sub

    Public Sub Lock()
        Dim balls_locked
        If IsNull(GetPlayerState(m_name & "_balls_locked")) Then
            balls_locked = 1
        Else
            balls_locked = GetPlayerState(m_name & "_balls_locked") + 1
        End If
        SetPlayerState m_name & "_balls_locked", balls_locked
        DispatchPinEvent m_name & "_locked_ball", balls_locked
        Log CStr(balls_locked)
        glf_BIP = glf_BIP - 1
        If balls_locked = m_balls_to_lock Then
            DispatchPinEvent m_name & "_full", balls_locked
        Else
            SetDelay m_name & "_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", Me),Null), 1000
        End If
    End Sub

    Public Sub Reset
        SetPlayerState m_name & "_balls_locked", 0
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MultiballLocksHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
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
            multiball.Lock
        Case "reset"
            multiball.Reset
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True Then
                ReleaseBall(Null)
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
    MultiballLocksHandler = kwargs
End Function

Class Multiball

    Private m_name
    Private m_priority
    Private m_mode
    Private m_multiball_locks
    Private m_enable_events
    Private m_disable_events
    Private m_start_events
    Private m_ball_save
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let StartEvents(value) : m_start_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, multiball_locks, mode)
        m_name = "multiball_" & name
        m_mode = mode.Name
        m_start_events = Array()
        m_multiball_locks = multiball_locks
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "MultiballHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "MultiballHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MultiballHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_starting", "MultiballHandler", m_priority, Array("start", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_start_events
            RemovePinEventListener evt, m_name & "_starting"
        Next
    End Sub

    Public Sub StartMultiball()
        glf_BIP = glf_BIP + GetPlayerState(m_multiball_locks & "_balls_locked")
        DispatchPinEvent m_name & "_started", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MultiballHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
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
        Case "start"
            multiball.StartMultiball
    End Select
    MultiballHandler = kwargs
End Function

Class ShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Let Events(value) : Set m_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", "show_player_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "show_player_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_show_player_play", "ShowPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_show_player_play"
        Next
    End Sub

    Public Sub Play(showItem)
        Log "Playing " & showItem.Name
        Dim show_step, stepIdx, lastTime
        stepIdx = 0
        lastTime = 125
        For Each show_step in showItem.Show
            lastTime = lastTime + show_step.Time
            SetDelay m_mode & "_show_player_play_step_" & stepIdx, "ShowPlayerEventHandler", Array(Array("play_step", Me), show_step), lastTime            
            stepIdx = stepIdx + 1
        Next
    End Sub

    Public Sub PlayStep(showStep)
        Dim light
        Log "Playing Step"
        For Each light in showStep.Lights
            If light(1) = "off" Then
                lightCtrl.LightOff light(0)
            Else
                If UBound(light) = 2 Then
                    lightCtrl.LightOn light(0)
                    lightCtrl.FadeLightToColor light(0), light(1), light(2)
                Else
                    lightCtrl.LightOnWithColor light(0), light(1)
                End If
            End If
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_mode & "_show_player", message
        End If
    End Sub
End Class

Function ShowPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ShowPlayer : Set ShowPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            ShowPlayer.Activate
        Case "deactivate"
            ShowPlayer.Deactivate
        Case "play"
            ShowPlayer.Play ownProps(2)
        Case "play_step"
            Dim show_step : Set show_step = args(1)
            ShowPlayer.PlayStep show_step
    End Select
    ShowPlayerEventHandler = Null
End Function

Class ShowPlayerItem

    Private m_name
    Private m_priority
    Private m_mode
    Private m_show
    Private m_speed
    Private m_tokens
    Private m_debug

    Private m_value

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Show(): Show = m_show: End Property

    Public Property Let Speed(value) : m_speed = value : End Property
    Public Property Let Tokens(value) : m_tokens = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode, show)
        m_mode = mode.Name
        m_name = m_mode & "_show_player_" & name
        m_priority = mode.Priority
        m_show = show
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        'Dim evt
        'For Each evt In m_events.Keys()
        '    AddPinEventListener evt, m_mode & "_show_player_play", "ShowPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        'Next
    End Sub

    Public Sub Deactivate()
        'Dim evt
        'For Each evt In m_events.Keys()
        '    RemovePinEventListener evt, m_mode & "_show_player_play"
        'Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Class ShowPlayerLightStep 

    Private m_time
    Private m_lights
    Private m_debug

    Public Property Get Time(): Time = m_time: End Property
    Public Property Get Lights(): Lights = m_lights: End Property

    Public Property Let Time(value) : m_time = value : End Property
    Public Property Let Lights(value) : m_lights = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(time, lights)
        m_time = time
        m_lights = lights
        m_debug = False
        Set Init = Me
	End Function

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class




Class ModeTimer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_start_events
    Private m_stop_events
    Private m_debug

    Private m_value

    Public Property Let StartValue(value) : m_start_value = value : End Property
    Public Property Let EndValue(value) : m_end_value = value : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let StartEvents(value) : m_start_events = value : End Property
    Public Property Let StopEvents(value) : m_stop_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "TimerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "TimerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_start", "TimerEventHandler", m_priority, Array("start", Me)
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events
                AddPinEventListener evt, m_name & "_stop", "TimerEventHandler", m_priority, Array("stop", Me)
            Next
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_start_events
            RemovePinEventListener evt, m_name & "_start"
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events
                RemovePinEventListener evt, m_name & "_stop"
            Next
        End If
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub StartTimer()
        Log "Started"
        DispatchPinEvent m_name & "_started", Null
        m_value = m_start_value
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
    End Sub

    Public Sub StopTimer()
        Log "Stopped"
        DispatchPinEvent m_name & "_stopped", Null
        RemoveDelay m_name & "_tick"
        m_value = m_start_value
    End Sub

    Public Sub Tick()
        Dim newValue
        If m_direction = "down" Then
            newValue = m_value - 1
        Else
            newValue = m_value + 1
        End If
        Log "ticking: old value: "& m_value & ", new Value: " & newValue & ", target: "& m_end_value
        m_value = newValue
        If m_value = m_end_value Then
            DispatchPinEvent m_name & "_complete", Null
        Else
            DispatchPinEvent m_name & "_tick", Null
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    Select Case evt
        Case "activate"
            timer.Activate
        Case "deactivate"
            timer.Deactivate
        Case "start"
            timer.StartTimer
        Case "stop"
            timer.StopTimer
        Case "tick"
            timer.Tick 
    End Select
    TimerEventHandler = kwargs
End Function


Class DelayObject
	Private m_name, m_callback, m_ttl, m_args
  
	Public Property Get Name(): Name = m_name: End Property
	Public Property Let Name(input): m_name = input: End Property
  
	Public Property Get Callback(): Callback = m_callback: End Property
	Public Property Let Callback(input): m_callback = input: End Property
  
	Public Property Get TTL(): TTL = m_ttl: End Property
	Public Property Let TTL(input): m_ttl = input: End Property
  
	Public Property Get Args(): Args = m_args: End Property
	Public Property Let Args(input): m_args = input: End Property
  
	Public default Function init(name, callback, ttl, args)
	  m_name = name
	  m_callback = callback
	  m_ttl = ttl
	  m_args = args

	  Set Init = Me
	End Function
End Class

Dim delayQueue : Set delayQueue = CreateObject("Scripting.Dictionary")
Dim delayQueueMap : Set delayQueueMap = CreateObject("Scripting.Dictionary")
Dim delayCallbacks : Set delayCallbacks = CreateObject("Scripting.Dictionary")

Sub SetDelay(name, callbackFunc, args, delayInMs)
    Dim executionTime
    executionTime = AlignToQuarterSecond(gametime + delayInMs)
    
    If delayQueueMap.Exists(name) Then
        delayQueueMap.Remove name
    End If
    

    If delayQueue.Exists(executionTime) Then
        If delayQueue(executionTime).Exists(name) Then
            delayQueue(executionTime).Remove name
        End If
    Else
        delayQueue.Add executionTime, CreateObject("Scripting.Dictionary")
    End If

    debugLog.WriteToLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc
    delayQueue(executionTime).Add name, (new DelayObject)(name, callbackFunc, executionTime, args)
    delayQueueMap.Add name, executionTime
    
End Sub

Function AlignToQuarterSecond(timeMs)
    AlignToQuarterSecond = Int(timeMs / 125) * 125
End Function

Function RemoveDelay(name)
    If delayQueueMap.Exists(name) Then
        If delayQueue.Exists(delayQueueMap(name)) Then
            delayQueue(delayQueueMap(name)).Remove name
            delayQueueMap.Remove name
            RemoveDelay = True
            debugLog.WriteToLog "Delay", "Removing delay for " & name
            Exit Function
        End If
    End If
    RemoveDelay = False
End Function

Sub DelayTick()
    Dim key, delayObject

    Dim executionTime
    executionTime = AlignToQuarterSecond(gametime)
    If delayQueue.Exists(executionTime) Then
        For Each key In delayQueue(executionTime).Keys()
            Set delayObject = delayQueue(executionTime)(key)
            debugLog.WriteToLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
            GetRef(delayObject.Callback)(delayObject.Args)
        Next
        delayQueue.Remove executionTime
    End If
End Sub
Class BallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeout
    Private m_balls
    Private m_balls_in_device
    Private m_eject_angle
    Private m_eject_pitch
    Private m_eject_strength
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_balls_to_eject
    Private m_ejecting_all
    Private m_ejecting
    Private m_mechcanical_eject
    Private m_eject_targets
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Let DefaultDevice(value)
        m_default_device = value
        If m_default_device = True Then
            Set glf_plunger = Me
        End If
    End Property
	Public Property Get HasBall(): HasBall = (Not IsNull(m_balls(0)) And m_ejecting = False): End Property
    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    
    Public Property Let EjectAngle(value) : m_eject_angle = glf_PI * (0 - 90) / 180 : End Property
    Public Property Let EjectPitch(value) : m_eject_pitch = glf_PI * (0 - 90) / 180 : End Property
    Public Property Let EjectStrength(value) : m_eject_strength = value : End Property
    
    Public Property Let EjectTimeout(value) : m_eject_timeout = value * 1000 : End Property
    Public Property Let EjectAllEvents(value)
        m_eject_all_events = value
        Dim evt
        For Each evt in m_eject_all_events
            AddPinEventListener evt, m_name & "_eject_all", "BallDeviceEventHandler", 1000, Array("ball_eject_all", Me)
        Next
    End Property
    Public Property Let EjectTargets(value)
        m_eject_targets = value
        Dim evt
        For Each evt in m_eject_targets
            AddPinEventListener evt & "_active", m_name & "_eject_target", "BallDeviceEventHandler", 1000, Array("eject_timeout", Me)
        Next
    End Property
    Public Property Let PlayerControlledEjectEvents(value)
        m_player_controlled_eject_event = value
        Dim evt
        For Each evt in m_player_controlled_eject_event
            AddPinEventListener evt, m_name & "_eject_attempt", "BallDeviceEventHandler", 1000, Array("ball_eject", Me)
        Next
    End Property
    Public Property Let BallSwitches(value)
        m_ball_switches = value
        m_balls = Array(Ubound(m_ball_switches))
        Dim x
        For x=0 to UBound(m_ball_switches)
            AddPinEventListener m_ball_switches(x)&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_entering", Me, x)
            AddPinEventListener m_ball_switches(x)&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, x)
        Next
    End Property
    Public Property Let MechcanicalEject(value) : m_mechcanical_eject = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property
        
	Public default Function init(name)
        m_name = "balldevice_" & name
        m_ball_switches = Array()
        m_eject_all_events = Array()
        m_eject_targets = Array()
        m_balls = Array()
        m_debug = False
        m_default_device = False
        m_eject_pitch = 0
        m_eject_angle = 0
        m_eject_strength = 0
        m_ejecting = False
        m_eject_callback = Null
        m_ejecting_all = False
        m_balls_to_eject = 0
        m_balls_in_device = 0
        m_eject_timeout = 0
	    Set Init = Me
	End Function

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_eject_timeout"
        'SoundSaucerLockAtBall ball
        Set m_balls(switch) = ball
        m_balls_in_device = m_balls_in_device + 1
        Log "Ball Entered" 
        Dim unclaimed_balls
        unclaimed_balls = DispatchRelayPinEvent(m_name & "_ball_entered", 1)
        Log "Unclaimed Balls: " & unclaimed_balls
        If (m_default_device = False Or m_ejecting = True) And unclaimed_balls > 0 And Not IsNull(m_balls(0)) Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        m_balls(switch) = Null
        m_balls_in_device = m_balls_in_device - 1
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechcanical_eject = True Then
            SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), ball), m_eject_timeout
        End If
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        m_ejecting = False
        RemoveDelay m_name & "_eject_timeout"
        DispatchPinEvent m_name & "_ball_eject_success", Null
        Log "Ball successfully exited"
        If m_ejecting_all = True Then
            If m_balls_to_eject = 0 Then
                m_ejecting_all = False
                Exit Sub
            End If
            If Not IsNull(m_balls(0)) Then
                m_balls_to_eject = m_balls_to_eject - 1
                Eject()
            Else
                SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 600
            End If
        End If
    End Sub

    Public Sub Eject
        Log "Ejecting."
        SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_balls(0)), m_eject_timeout
        m_ejecting = True
        If m_eject_strength > 0 Then
            m_balls(0).VelX = m_eject_strength * Cos(m_eject_pitch) * Sin(m_eject_angle)
            m_balls(0).VelY = m_eject_strength * Cos(m_eject_pitch) * Cos(m_eject_angle)
            m_balls(0).VelZ = m_eject_strength * Sin(m_eject_pitch)
            Log "VelX: " &  m_balls(0).VelX & ", VelY: " &  m_balls(0).VelY & ", VelZ: " &  m_balls(0).VelZ
        End If

        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(m_balls(0))
        End If
    End Sub

    Public Sub EjectAll
        Log "Ejecting All."
        m_ejecting_all = True
        m_balls_to_eject = m_balls_in_device
        Eject()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallDeviceEventHandler(args)
    Dim ownProps, ball : ownProps = args(0) : Set ball = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballDevice : Set ballDevice = ownProps(1)
    Dim switch
    Select Case evt
        Case "ball_entering"
            switch = ownProps(2)
            SetDelay ballDevice.Name & "_" & switch & "_ball_enter", "BallDeviceEventHandler", Array(Array("ball_enter", ballDevice, switch), ball), 500
        Case "ball_enter"
            switch = ownProps(2)
            ballDevice.BallEnter ball, switch
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_eject_all"
            ballDevice.EjectAll
        Case "ball_exiting"
            switch = ownProps(2)
            If RemoveDelay(ballDevice.Name & "_" & switch & "_ball_enter") = False Then
                ballDevice.BallExiting ball, switch
            End If
        Case "eject_timeout"
            ballDevice.BallExitSuccess ball
    End Select
End Function

Class Diverter

    Private m_name
    Private m_activate_events
    Private m_deactivate_events
    Private m_activation_time
    Private m_enable_events
    Private m_disable_events
    Private m_activation_switches
    Private m_action_cb
    Private m_debug

    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let ActivateEvents(value) : m_activate_events = value : End Property
    Public Property Let DeactivateEvents(value) : m_deactivate_events = value : End Property
    Public Property Let ActivationTime(value) : m_activation_time = value : End Property
    Public Property Let ActivationSwitches(value) : m_activation_switches = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, enable_events, disable_events)
        m_name = "diverter_" & name
        m_enable_events = enable_events
        m_disable_events = disable_events
        m_activate_events = Array()
        m_deactivate_events = Array()
        m_activation_switches = Array()
        m_activation_time = 0
        m_debug = False
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "DiverterEventHandler", 1000, Array("enable", Me)
        Next
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "DiverterEventHandler", 1000, Array("disable", Me)
        Next
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_activate_events
            AddPinEventListener evt, m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
        For Each evt in m_deactivate_events
            AddPinEventListener evt, m_name & "_deactivate", "DiverterEventHandler", 1000, Array("deactivate", Me)
        Next
        For Each evt in m_activation_switches
            AddPinEventListener evt & "_active", m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_activate_events
            RemovePinEventListener evt, m_name & "_activate"
        Next
        For Each evt in m_deactivate_events
            RemovePinEventListener evt, m_name & "_deactivate"
        Next
        For Each evt in m_activation_switches
            RemovePinEventListener evt & "_active", m_name & "_activate"
        Next
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
    End Sub

    Public Sub Activate()
        Log "Activating"
        GetRef(m_action_cb)(1)
        If m_activation_time > 0 Then
            SetDelay m_name & "_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_activation_time
        End If
        DispatchPinEvent m_name & "_activating", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_deactivating", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DiverterEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim diverter : Set diverter = ownProps(1)
    Select Case evt
        Case "enable"
            diverter.Enable
        Case "disable"
            diverter.Disable
        Case "activate"
            diverter.Activate
        Case "deactivate"
            diverter.Deactivate
    End Select
    DiverterEventHandler = kwargs
End Function
Class DropTarget
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


'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub Glf_AddPlayer()
    Select Case UBound(playerState.Keys())
        Case -1:
            playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            Glf_BcpAddPlayer 1
            currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 2", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 2
            End If
        Case 1:
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 3", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 3
            End If     
        Case 2:   
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 4", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 4
            End If  
            canAddPlayers = False
    End Select
End Sub

Function Glf_InitNewPlayer()
    Dim state : Set state = CreateObject("Scripting.Dictionary")
    state.Add GLF_SCORE, 0
    state.Add GLF_INITIALS, ""
    state.Add GLF_CURRENT_BALL, 1
    Set Glf_InitNewPlayer = state
End Function


'****************************
' Setup Player
' Event Listeners:  
    AddPinEventListener GLF_GAME_STARTED,   "start_game_setup",   "Glf_SetupPlayer", 1000, Null
    AddPinEventListener GLF_NEXT_PLAYER,    "next_player_setup",  "Glf_SetupPlayer", 1000, Null
'
'*****************************
Function Glf_SetupPlayer(args)
    EmitAllPlayerEvents()
End Function

'****************************
' StartGame
'
'*****************************
Sub Glf_StartGame()
    gameStarted = True
    If useBcp Then
        bcpController.Send "player_turn_start?player_num=int:1"
        bcpController.Send "ball_start?player_num=int:1&ball=int:1"
        bcpController.PlaySlide "attract", "base", 1000
        bcpController.SendPlayerVariable "number", 1, 0
    End If
    DispatchPinEvent GLF_GAME_STARTED, Null
End Sub


'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener GLF_GAME_STARTED, "start_game_release_ball",   "Glf_ReleaseBall", 1000, True
AddPinEventListener GLF_NEXT_PLAYER, "next_player_release_ball",   "Glf_ReleaseBall", 1000, True
'
'*****************************
Function Glf_ReleaseBall(args)
    If Not IsNull(args) Then
        If args(0) = True Then
            DispatchPinEvent GLF_BALL_STARTED, Null
            If useBcp Then
                bcpController.SendPlayerVariable GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL), GetPlayerState(GLF_CURRENT_BALL)-1
                bcpController.SendPlayerVariable GLF_SCORE, GetPlayerState(GLF_SCORE), GetPlayerState(GLF_SCORE)
            End If
        End If
    End If
    debugLog.WriteToLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    swTrough1.kick 90, 10
    debugLog.WriteToLog "Release Ball", "Just Kicked"
    glf_BIP = glf_BIP + 1
End Function


'****************************
' End Of Ball
' Event Listeners:      
    AddPinEventListener GLF_BALL_DRAIN, "ball_drain", "Glf_Drain", 20, Null
'
'*****************************
Function Glf_Drain(args)
    
    Dim ballsToSave : ballsToSave = args(1) 
    debugLog.WriteToLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
    debugLog.WriteToLog "end_of_ball, balls in play", CStr(glf_BIP)
    If ballsToSave <= 0 Then
        Exit Function
    End If

    If glf_BIP > 0 Then
        Exit Function
    End If
        
    DispatchPinEvent GLF_BALL_ENDED, Null
    SetPlayerState GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = GetCurrentPlayerNumber()
    Select Case currentPlayer
        Case "PLAYER 1":
            If UBound(playerState.Keys()) > 0 Then
                currentPlayer = "PLAYER 2"
            End If
        Case "PLAYER 2":
            If UBound(playerState.Keys()) > 1 Then
                currentPlayer = "PLAYER 3"
            Else
                currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 3":
            If UBound(playerState.Keys()) > 2 Then
                currentPlayer = "PLAYER 4"
            Else
                currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 4":
            currentPlayer = "PLAYER 1"
    End Select
    
    If useBcp Then
        bcpController.SendPlayerVariable "number", GetCurrentPlayerNumber(), previousPlayerNumber
    End If
    If GetPlayerState(GLF_CURRENT_BALL) > glf_ballsPerGame Then
        DispatchPinEvent GLF_GAME_OVER, Null
        gameStarted = False
        currentPlayer = Null
        playerState.RemoveAll()
    Else
        DispatchPinEvent GLF_NEXT_PLAYER, Null
    End If
    
End Function


'***********************************************************************************************************************
' Lights State Controller - 0.9.1
'  
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************


Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqOffset, m_tableSeqSpeed, m_tableSeqDirection, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, m_lightmaps, m_seqOverrideRunners, m_pauseMainLights, m_pausedLights, m_minX, m_minY, m_maxX, m_maxY, m_width, m_height, m_centerX, m_centerY, m_coordsX, m_coordsY, m_angles, m_radii, m_tags

    Private Sub Class_Initialize()
        Set m_lights = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_off = CreateObject("Scripting.Dictionary")
        Set m_seqRunners = CreateObject("Scripting.Dictionary")
        Set m_seqOverrideRunners = CreateObject("Scripting.Dictionary")
        Set m_currentFrameState = CreateObject("Scripting.Dictionary")
        Set m_seqs = CreateObject("Scripting.Dictionary")
        Set m_pulse = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_tags = CreateObject("Scripting.Dictionary")
        m_vpxLightSyncRunning = False
        m_vpxLightSyncCollection = Null
        m_initFrameTime = 0
        m_frameTime = 0
        m_pulseInterval = 26
        m_vpxLightSyncClear = False
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        m_pauseMainLights = False
        Set m_pausedLights = CreateObject("Scripting.Dictionary")
        Set m_lightmaps = CreateObject("Scripting.Dictionary")
        m_minX = 1000000
        m_minY = 1000000
        m_maxX = -1000000
        m_maxY = -1000000
        m_centerX = Round(tableWidth/2)
        m_centerY = Round(tableHeight/2)
    End Sub

    Private Sub AssignStateForFrame(key, state)
        If m_currentFrameState.Exists(key) Then
            m_currentFrameState.Remove key
        End If
        m_currentFrameState.Add key, state
    End Sub

    Public Sub LoadLightShows()
        Dim oFile
        Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-out.txt",2,true)
        For Each oFile In oFSO.GetFolder(cGameName & "_LightShows").Files
            If LCase(oFSO.GetExtensionName(oFile.Name)) = "yaml" And Not Left(oFile.Name,6) = "lights" Then
                Dim textStream : Set textStream = oFSO.OpenTextFile(oFile.Path, 1)
                Dim show : show = textStream.ReadAll
                Dim fileName : fileName = "lSeq" & Replace(oFSO.GetFileName(oFile.Name), "."&oFSO.GetExtensionName(oFile.Name), "")
                Dim lcSeq : lcSeq = "Dim " & fileName & " : Set " & fileName & " = New LCSeq"&vbCrLf
                lcSeq = lcSeq + fileName & ".Name = """&fileName&""""&vbCrLf
                Dim seq : seq = ""
                Dim re : Set re = New RegExp
                With re
                    .Pattern    = "- time:.*?\n"
                    .IgnoreCase = False
                    .Global     = True
                End With
                Dim matches : Set matches = re.execute(show)
                Dim steps : steps = matches.Count
                Dim match, nextMatchIndex, uniqueLights
                Set uniqueLights = CreateObject("Scripting.Dictionary")
                nextMatchIndex = 1
                For Each match in matches
                    Dim lightStep
                    If Not nextMatchIndex < steps Then
                        lightStep = Mid(show, match.FirstIndex, Len(show))
                    Else
                        lightStep = Mid(show, match.FirstIndex, matches(nextMatchIndex).FirstIndex - match.FirstIndex)
                        nextMatchIndex = nextMatchIndex + 1
                    End If

                    Dim re1 : Set re1 = New RegExp
                    With re1
                        .Pattern        = ".*:?: '([A-Fa-f0-9]{6})'"
                        .IgnoreCase     = True
                        .Global         = True
                    End With

                    Dim lightMatches : Set lightMatches = re1.execute(lightStep)
                    If lightMatches.Count > 0 Then
                        Dim lightMatch, lightStr, lightSplit
                        lightStr = "Array("
                        lightSplit = 0
                        For Each lightMatch in lightMatches
                            Dim sParts : sParts = Split(lightMatch.Value, ":")
                            Dim lightName : lightName = Trim(sParts(0))
                            Dim color : color = Trim(Replace(sParts(1),"'", ""))
                            If color = "000000" Then
                                lightStr = lightStr + """"&lightName&"|0|000000"","
                            Else
                                lightStr = lightStr + """"&lightName&"|100|"&color&""","
                            End If

                            If Len(lightStr)+20 > 2000 And lightSplit = 0 Then                           
                                lightSplit = Len(lightStr)
                            End If

                            uniqueLights(lightname) = 0
                        Next
                        lightStr = Left(lightStr, Len(lightStr) - 1)
                        lightStr = lightStr & ")"
                        
                        If lightSplit > 0 Then
                            lightStr = Left(lightStr, lightSplit) & " _ " & vbCrLF & Right(lightStr, Len(lightStr)-lightSplit)
                        End If

                        seq = seq + lightStr & ", _"&vbCrLf
                    Else
                        seq = seq + "Array(), _"&vbCrLf
                    End If

                    
                    Set re1 = Nothing
                Next
                
                lcSeq = lcSeq + filename & ".Sequence = Array( " & Left(seq, Len(seq) - 5) & ")"&vbCrLf
                'lcSeq = lcSeq + seq & vbCrLf
                lcSeq = lcSeq + fileName & ".UpdateInterval = 20"&vbCrLf
                lcSeq = lcSeq + fileName & ".Color = Null"&vbCrLf
                lcSeq = lcSeq + fileName & ".Repeat = False"&vbCrLf

                'MsgBox(lcSeq)
                objFileToWrite.WriteLine(lcSeq)
                ExecuteGlobal lcSeq
                Set re = Nothing

                textStream.Close
            End if
        Next
        'Clean up
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Set oFile = Nothing
        Set oFSO = Nothing
    End Sub

    Public Sub CompileLights(collection, name)
        Dim light
        Dim lights : lights = "light:" & vbCrLf
        For Each light in collection
            lights = lights + light.name & ":"&vbCrLf
            lights = lights + "   x: "& light.x/tablewidth & vbCrLf
            lights = lights + "   y: "& light.y/tableheight & vbCrLf
        Next
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-"&name&".yaml",2,true)
        objFileToWrite.WriteLine(lights)
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Debug.print("Lights YAML File saved to: " & cGameName & "LightShows/lights-"&name&".yaml")
    End Sub

    Dim leds
    Dim ledGrid()
    Dim lightsToLeds

    Sub PrintLEDs
        Dim light
        Dim lights : lights = ""
    
        Dim row,col,value,i
        For row = LBound(ledGrid, 1) To UBound(ledGrid, 1)
            For col = LBound(ledGrid, 2) To UBound(ledGrid, 2)
                ' Access the array element and do something with it
                value = ledGrid(row, col)
                lights = lights + cstr(value) & vbTab
            Next
            lights = lights + vbCrLf
        Next

        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/led-grid.txt",2,true)
        objFileToWrite.WriteLine(lights)
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Debug.print("Lights File saved to: " & cGameName & "LightShows/led-grid.txt")


        lights = ""
        For i = 0 To UBound(leds)
            value = leds(i)
            If IsArray(value) Then
                lights = lights + "Index: " & cstr(value(0)) & ". X: " & cstr(value(1)) & ". Y:" & cstr(value(2)) & ". Angle:" & cstr(value(3)) & ". Radius:" & cstr(value(4)) & ". CoordsX:" & cstr(value(5)) & ". CoordsY:" & cstr(value(6)) & ". Angle256:" & cstr(value(7)) &". Radii256:" & cstr(value(8)) &","
            End If
            lights = lights + vbCrLf
            'lights = lights + cstr(value) & ","
            
        Next

        
        Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/coordsX.txt",2,true)
        objFileToWrite.WriteLine(lights)
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Debug.print("Lights File saved to: " & cGameName & "LightShows/coordsX.txt")


    End Sub

    Public Sub RegisterLights(mode)

        Dim idx,tmp,vpxLight,lcItem
    
            vpmMapLights aLights
            Dim colCount : colCount = Round(tablewidth/20)
            Dim rowCount : rowCount = Round(tableheight/20)
                
            dim ledIdx : ledIdx = 0
            redim leds(UBound(Lights)-1)
            redim lightsToLeds(UBound(Lights)-1)
            ReDim ledGrid(rowCount,colCount)
            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
                debug.print("TRYING TO REGISTER IDX: " & idx)
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
                    debug.print("TEMP LIGHT NAME for idx:" & idx & ", light: " & vpxLight.name)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
                If Not IsNull(vpxLight) Then
                    Debug.print("Registering Light: "& vpxLight.name)


                    Dim r : r = Round(vpxLight.y/20)
                    Dim c : c = Round(vpxLight.x/20)
                    If r < rowCount And c < colCount And r >= 0 And c >= 0 Then
                        If Not ledGrid(r,c) = "" Then
                            MsgBox(vpxLight.name & " is too close to another light")
                        End If
                        leds(ledIdx) = Array(ledIdx, c, r, 0,0,0,0,0,0) 'index, row, col, angle, radius, x256, y256, angle256, radius256
                        lightsToLeds(idx) = ledIdx
                        ledGrid(r,c) = ledIdx
                        ledIdx = ledIdx + 1
                        If (c < m_minX) Then : m_minX = c
                        if (c > m_maxX) Then : m_maxX = c
                
                        if (r < m_minY) Then : m_minY = r
                        if (r > m_maxY) Then : m_maxY = r
                    End If
                    Dim e, lmStr: lmStr = "lmArr = Array("    
                    For Each e in GetElements()
                        If InStr(LCase(e.Name), LCase("_" & vpxLight.Name & "_")) Or InStr(LCase(e.Name), LCase("_" & vpxLight.UserValue & "_")) Then
                            Debug.Print(e.Name)
                            lmStr = lmStr & e.Name & ","
                        End If
                    Next
                    lmStr = lmStr & "Null)"
                    lmStr = Replace(lmStr, ",Null)", ")")
                    ExecuteGlobal "Dim lmArr : "&lmStr
                    m_lightmaps.Add vpxLight.Name, lmArr
                    Debug.print("Registering Light: "& vpxLight.name) 
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next
            'ReDim Preserve leds(ledIdx)
            m_width = m_maxX - m_minX + 1
            m_height = m_maxY - m_minY + 1
            m_centerX = m_width / 2
            m_centerY = m_height / 2
            GenerateLedMapCode()
    End Sub

    Private Sub GenerateLedMapCode()

        Dim minX256, minY256, minAngle, minAngle256, minRadius, minRadius256
        Dim maxX256, maxY256, maxAngle, maxAngle256, maxRadius, maxRadius256
        Dim i, led
        minX256 = 1000000
        minY256 = 1000000
        minAngle = 1000000
        minAngle256 = 1000000
        minRadius = 1000000
        minRadius256 = 1000000

        maxX256 = -1000000
        maxY256 = -1000000
        maxAngle = -1000000
        maxAngle256 = -1000000
        maxRadius = -1000000
        maxRadius256 = -1000000

        For i = 0 To UBound(leds)
            led = leds(i)
            If IsArray(led) Then
                
                Dim x : x = led(1)
                Dim y : y = led(2)
            
                Dim radius : radius = Sqr((x - m_centerX) ^ 2 + (y - m_centerY) ^ 2)
                Dim radians: radians = Atn2(m_centerY - y, m_centerX - x)
                Dim angle
                angle = radians * (180 / 3.141592653589793)
                Do While angle < 0
                    angle = angle + 360
                Loop
                Do While angle > 360
                    angle = angle - 360
                Loop
            
                If angle < minAngle Then
                    minAngle = angle
                End If
                If angle > maxAngle Then
                    maxAngle = angle
                End If
            
                If radius < minRadius Then
                    minRadius = radius
                End If
                If radius > maxRadius Then
                    maxRadius = radius
                End If
            
                led(3) = angle
                led(4) = radius
                leds(i) = led
            End If
        Next

        For i = 0 To UBound(leds)
            led = leds(i)
            If IsArray(led) Then
                Dim x256 : x256 = MapNumber(led(1), m_minX, m_maxX, 0, 255)
                Dim y256 : y256 = MapNumber(led(2), m_minY, m_maxY, 0, 255)
                Dim angle256 : angle256 = MapNumber(led(3), 0, 360, 0, 255)
                Dim radius256 : radius256 = MapNumber(led(4), 0, maxRadius, 0, 255)
            
                led(5) = Round(x256)
                led(6) = Round(y256)
                led(7) = Round(angle256)
                led(8) = Round(radius256)
            
                If x256 < minX256 Then minX256 = x256
                If x256 > maxX256 Then maxX256 = x256
            
                If y256 < minY256 Then minY256 = y256
                If y256 > maxY256 Then maxY256 = y256
            
                If angle256 < minAngle256 Then minAngle256 = angle256
                If angle256 > maxAngle256 Then maxAngle256 = angle256
            
                If radius256 < minRadius256 Then minRadius256 = radius256
                If radius256 > maxRadius256 Then maxRadius256 = radius256
                leds(i) = led
            End If
        Next

        reDim m_coordsX(UBound(leds)-1)
        reDim m_coordsY(UBound(leds)-1)
        reDim m_angles(UBound(leds)-1)
        reDim m_radii(UBound(leds)-1)
        
        For i = 0 To UBound(leds)
            led = leds(i)
            If IsArray(led) Then
                m_coordsX(i)    =  leds(i)(5) 'x256
                m_coordsY(i)    =  leds(i)(6) 'y256
                m_angles(i)     =  leds(i)(7) 'angle256
                m_radii(i)      =  leds(i)(8) 'radius256
            End If
        Next

    End Sub

    Private Function MapNumber(l, inMin, inMax, outMin, outMax)
        If (inMax - inMin + outMin) = 0 Then
            MapNumber = 0
        Else
            MapNumber = ((l - inMin) * (outMax - outMin)) / (inMax - inMin) + outMin
        End If
    End Function

    Private Function ReverseArray(arr)
        Dim i, upperBound
        upperBound = UBound(arr)

        ' Create a new array of the same size
        Dim reversedArr()
        ReDim reversedArr(upperBound)

        ' Fill the new array with elements in reverse order
        For i = 0 To upperBound
            reversedArr(i) = arr(upperBound - i)
        Next

        ReverseArray = reversedArr
    End Function

    Private Function Atn2(dy, dx)
        If dx > 0 Then
            Atn2 = Atn(dy / dx)
        ElseIf dx < 0 Then
            If dy = 0 Then 
                Atn2 = pi
            Else
                Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
            end if
        ElseIf dx = 0 Then
            if dy = 0 Then
                Atn2 = 0
            else
                Atn2 = Sgn(dy) * pi / 2
            end if
        End If
    End Function

    Private Function ColtoArray(aDict)	'converts a collection to an indexed array. Indexes will come out random probably.
        redim a(999)
        dim count : count = 0
        dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
        redim preserve a(count-1) : ColtoArray = a
    End Function

    Function IncrementUInt8(x, increment)
        If x + increment > 255 Then
            IncrementUInt8 = x + increment - 256
        Else
            IncrementUInt8 = x + increment
        End If
    End Function

    Public Sub AddLight(light, idx)
        If m_lights.Exists(light.name) Then
            Exit Sub
        End If
        Dim lcItem : Set lcItem = new LCItem
        lcItem.Init idx, light.BlinkInterval, Array(light.color, light.colorFull), light.name, light.x, light.y
        m_lights.Add light.Name, lcItem
        m_seqRunners.Add "lSeqRunner" & CStr(light.name), new LCSeqRunner
    End Sub

    Public Sub AddLightTags(light, tags)
        If m_lights.Exists(light.name) Then
            Dim tagArray, tag, lightName
            lightName = light.name
            tagArray = Split(tags, ",")
            
            For Each tag In tagArray
                tag = Trim(tag) ' Remove any leading or trailing spaces
                If Not m_tags.Exists(tag) Then
                    Set m_tags(tag) = CreateObject("Scripting.Dictionary")
                End If
                If Not m_tags(tag).Exists(lightName) Then
                    m_tags(tag).Add lightName, True
                End If
            Next
        End If
    End Sub

    Public Function GetLightsForTag(tag)
        Dim lightsArray, lightName
        If m_tags.Exists(tag) Then
            ReDim lightsArray(m_tags(tag).Count - 1)
            Dim i : i = 0
            For Each lightName In m_tags(tag)
                lightsArray(i) = lightName
                i = i + 1
            Next
            GetLightsForTag = lightsArray
        Else
            GetLightsForTag = Array()
        End If
    End Function

    Public Sub LightState(light, state)
        m_lightOff(light.name)
        If state = 1 Then
            m_lightOn(light.name)
        ElseIF state = 2 Then
            Blink(light)
        End If
    End Sub

    Public Sub LightOn(light)
        m_LightOn(light.name)
    End Sub

    Public Sub LightOnWithColor(light, color)
        m_LightOnWithColor light.name, color
    End Sub

    Public Sub FadeLightToColor(light, color, fadeSpeed)
        If m_lights.Exists(light.name) Then
            dim lightColor, steps
            steps = Round(fadeSpeed/20)
            If steps < 10 Then
                steps = 10
            End If
            lightColor = m_lights(light.name).Color
            Dim seq : Set seq = new LCSeq
            seq.Name = light.name & "Fade"
            seq.Sequence = FadeRGB(light.name, lightColor(0), color, steps)
            seq.Color = Null
            seq.UpdateInterval = 20
            seq.Repeat = False
            m_lights(light.name).Color = color
            m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem light.name & "Fade", seq, 1, 20, Null
            If color = RGB(0,0,0) Then
                m_lightOff(light.name)
            End If
        End If
    End Sub

    Public Sub FlickerOn(light)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            m_lightOn(name)

            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70), 0, m_pulseInterval, 1, null)
        End If
    End Sub  
    
    Public Sub LightColor(light, color)

        If m_lights.Exists(light.name) Then
            m_lights(light.name).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Color = color
            End If

        End If
    End Sub

    Public Function GetLightColor(light)
        If m_lights.Exists(light.name) Then
            Dim colors : colors = m_lights(light.name).Color
            GetLightColor = colors(0)
        Else
            GetLightColor = RGB(0,0,0)
        End If
    End Function

    Private Sub m_LightOn(name)
        
        If m_lights.Exists(name) Then
            
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If
            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Private Sub m_LightOnWithColor(name, color)
        If m_lights.Exists(name) Then
            m_lights(name).Color = color
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Public Sub LightOff(light)
        m_lightOff(light.name)
    End Sub

    Private Sub m_lightOff(name)
        If m_lights.Exists(name) Then
            If m_on.Exists(name) Then 
                m_on.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_off.Exists(name) Then 
                Exit Sub
            End If
            LightColor m_lights(name), m_lights(name).BaseColor
            m_off.Add name, m_lights(name)
        End If
    End Sub

    Public Sub UpdateBlinkInterval(light, interval)
        If m_lights.Exists(light.name) Then
            light.BlinkInterval = interval
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs.Item(light.name & "Blink").UpdateInterval = interval
            End If
        End If
    End Sub

    Public Sub Pulse(light, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            'Array(100,94,32,13,6,3,0)
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount, null)
        End If
    End Sub

    Public Sub PulseWithColor(light, color, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            'Array(100,94,32,13,6,3,0)
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount,  Array(color,null))
        End If
    End Sub

    Public Sub PulseWithProfile(light, profile, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), profile, 0, m_pulseInterval, repeatCount, null)
        End If
    End Sub       

    Public Sub PulseWithState(pulse)
        
        If m_lights.Exists(pulse.Light) Then
            If m_off.Exists(pulse.Light) Then 
                m_off.Remove(pulse.Light)
            End If
            If m_pulse.Exists(pulse.Light) Then 
                Exit Sub
            End If
            m_pulse.Add name, pulse
        End If
    End Sub

    Public Sub LightLevel(light, lvl)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Level = lvl

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Sequence = m_buildBlinkSeq(light.name, light.BlinkPattern)
            End If
        End If
    End Sub

    Public Sub AddShot(name, light, color)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(name & light.name) Then
                m_seqs(name & light.name).Color = color
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddSequenceItem m_seqs(name & light.name)
            Else
                Dim stateOn : stateOn = light.name&"|100"
                Dim stateOff : stateOff = light.name&"|0"
                Dim seq : Set seq = new LCSeq
                seq.Name = name
                seq.Sequence = Array(stateOn, stateOff,stateOn, stateOff)
                seq.Color = color
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem name, seq, -1, light.BlinkInterval, Null
                m_seqs.Add name & light.name, seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub RemoveShot(name, light)
        If m_lights.Exists(light.name) And m_seqs.Exists(name & light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveItem m_seqs(name & light.name)
            If IsNUll(m_seqRunners("lSeqRunner"&CStr(light.name)).CurrentItem) Then
            LightOff(light)
            End If
        End If
    End Sub

    Public Sub RemoveAllShots()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub RemoveShotsFromLight(light)
        If m_lights.Exists(light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveAll   
            m_lightOff(light.name)  
        End If
    End Sub

    Public Sub Blink(light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").ResetInterval
                m_seqs(light.name & "Blink").CurrentIdx = 0
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddSequenceItem m_seqs(light.name & "Blink")
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = light.name & "Blink"
                seq.Sequence = m_buildBlinkSeq(light.name, light.BlinkPattern)
                seq.Color = Null
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem light.name & "Blink", seq, -1, light.BlinkInterval, Null
                m_seqs.Add light.name & "Blink", seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub AddLightToBlinkGroup(group, light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(group & "BlinkGroup") Then

                Dim i, pattern, buff : buff = Array()
                pattern = m_seqs(group & "BlinkGroup").Pattern
                ReDim buff(Len(pattern)-1)
                For i = 0 To Len(pattern)-1
                    Dim lightInSeq, ii, p, buff2
                    buff2 = Array()
                    If Mid(pattern, i+1, 1) = 1 Then
                        p=1
                    Else
                        p=0
                    End If
                    ReDim buff2(UBound(m_seqs(group & "BlinkGroup").LightsInSeq)+1)
                    ii=0
                    For Each lightInSeq in m_seqs(group & "BlinkGroup").LightsInSeq
                        If p = 1 Then
                            buff2(ii) = lightInSeq & "|100"
                        Else
                            buff2(ii) = lightInSeq & "|0"
                        End If
                        ii = ii + 1
                    Next
                    If p = 1 Then
                        buff2(ii) = light.name & "|100"
                    Else
                        buff2(ii) = light.name & "|0"
                    End If
                    buff(i) = buff2
                Next
                m_seqs(group & "BlinkGroup").Sequence = buff
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = group & "BlinkGroup"
                seq.Sequence = Array(Array(light.name & "|100"), Array(light.name & "|0"))
                seq.Color = Null
                seq.Pattern = "10"
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True
                m_seqRunners.Add "lSeqRunner" & group & "BlinkGroup", new LCSeqRunner
                m_seqs.Add group & "BlinkGroup", seq
            End If
        End If
    End Sub

    Public Sub RemoveLightFromBlinkGroup(group, light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(group & "BlinkGroup") Then

                Dim i, pattern, buff : buff = Array()
                pattern = m_seqs(group & "BlinkGroup").Pattern
                ReDim buff(Len(pattern)-1)
                For i = 0 To Len(pattern)-1
                    Dim lightInSeq, ii, p, buff2
                    buff2 = Array()
                    If Mid(pattern, i+1, 1) = 1 Then
                        p=1
                    Else
                        p=0
                    End If
                    ReDim buff2(UBound(m_seqs(group & "BlinkGroup").LightsInSeq)-1)
                    ii=0
                    For Each lightInSeq in m_seqs(group & "BlinkGroup").LightsInSeq
                        If Not lightInSeq = light.name Then
                            If p = 1 Then
                                buff2(ii) = lightInSeq & "|100"
                            Else
                                buff2(ii) = lightInSeq & "|0"
                            End If
                            ii = ii + 1
                        End If
                    Next
                    buff(i) = buff2
                Next
                AssignStateForFrame light.name, (new FrameState)(0, Null, m_lights(light.name).Idx)
                m_seqs(group & "BlinkGroup").Sequence = buff
            End If
        End If
    End Sub

    Public Sub UpdateBlinkGroupPattern(group, pattern)
        If m_seqs.Exists(group & "BlinkGroup") Then

            Dim i, buff : buff = Array()
            m_seqs(group & "BlinkGroup").Pattern = pattern
            ReDim buff(Len(pattern)-1)
            For i = 0 To Len(pattern)-1
                Dim lightInSeq, ii, p, buff2
                buff2 = Array()
                If Mid(pattern, i+1, 1) = 1 Then
                    p=1
                Else
                    p=0
                End If
                ReDim buff2(UBound(m_seqs(group & "BlinkGroup").LightsInSeq))
                ii=0
                For Each lightInSeq in m_seqs(group & "BlinkGroup").LightsInSeq
                    If p = 1 Then
                        buff2(ii) = lightInSeq & "|100"
                    Else
                        buff2(ii) = lightInSeq & "|0"
                    End If
                    ii = ii + 1
                Next
                buff(i) = buff2
            Next
            m_seqs(group & "BlinkGroup").Sequence = buff
        End If
    End Sub

    Public Sub UpdateBlinkGroupInterval(group, interval)
        If m_seqs.Exists(group & "BlinkGroup") Then
            m_seqs(group & "BlinkGroup").UpdateInterval = interval
        End If 
    End Sub
    
    Public Sub StartBlinkGroup(group)
        If m_seqs.Exists(group & "BlinkGroup") Then
            If Not m_seqRunners.Exists("lSeqRunner" & group & "BlinkGroup") Then
                CreateSeqRunner "lSeqRunner" & group & "BlinkGroup"
            End If
            m_seqRunners("lSeqRunner" & group & "BlinkGroup").AddSequenceItem m_seqs(group & "BlinkGroup")
        End If
    End Sub

    Public Sub StopBlinkGroup(group)
        If m_seqs.Exists(group & "BlinkGroup") Then
            RemoveLightSeq "lSeqRunner" & group & "BlinkGroup", m_seqs(group & "BlinkGroup").Name
        End If
    End Sub

    Public Function GetLightState(light)
        GetLightState = 0
        If(m_lights.Exists(light.name)) Then
            If m_on.Exists(light.name) Then
                GetLightState = 1
            Else
                If m_seqs.Exists(light.name & "Blink") Then
                    GetLightState = 2
                End If
            End If
        End If
    End Function

    Public Function IsShotLit(name, light)
        IsShotLit = False
        If(m_lights.Exists(light.name)) Then
            If m_seqRunners("lSeqRunner"&CStr(light.name)).HasSeq(name) Then
                IsShotLit = True
            End If
        End If
    End Function

    Public Sub CreateSeqRunner(name)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqRunners.Add name, seqRunner
    End Sub

    Private Sub CreateOverrideSeqRunner(name)
        If m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqOverrideRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(runner, key, sequence, loops, speed, tokens)
        If Not m_seqRunners.Exists(runner) Then
            CreateSeqRunner runner
        End If

        m_seqRunners(runner).AddItem key, sequence, loops, speed, tokens
    End Sub

    Public Sub RemoveLightSeq(runner, key)
        If Not m_seqRunners.Exists(runner) Then
            Exit Sub
        End If

        Dim light
        For Each light in m_seqRunners(runner).ItemByKey(key).LightsInSeq
            If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            End If
        Next

        m_seqRunners(lcSeqRunner).RemoveItem lcSeq
    End Sub

    Public Sub RemoveAllLightSeq(lcSeqRunner)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If
        Dim lcSeqKey, light, seqs, lcSeq
        Set seqs = m_seqRunners(lcSeqRunner).Items()
        For Each lcSeqKey in seqs.Keys()
            Set lcSeq = seqs(lcSeqKey)
            For Each light in lcSeq.LightsInSeq
                If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
                End If
            Next
        Next

        m_seqRunners(lcSeqRunner).RemoveAll
    End Sub

    Public Sub AddTableLightSeq(name, lcSeq)
        CreateOverrideSeqRunner(name)

        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
        m_seqOverrideRunners(name).AddSequenceItem lcSeq
    End Sub

    Public Sub RemoveTableLightSeq(name, lcSeq)
        If Not m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        m_seqOverrideRunners(name).RemoveItem lcSeq
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
    End Sub

    Public Sub RemoveAllTableLightSeqs()
        Dim light, runner
        For Each runner in m_seqOverrideRunners.Keys()
            m_seqOverrideRunners(runner).RemoveAll()
        Next
        For Each light in m_lights.Keys()
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub SyncLightMapColors()
        dim light,lm
        For Each light in m_lights.Keys()
            If m_lightmaps.Exists(light) Then
                For Each lm in m_lightmaps(light)
                    dim color : color = m_lights(light).Color
                    If not IsNull(lm) Then
                        lm.Color = color(0)
                    End If
                Next
            End If
        Next
    End Sub

    Public Sub SyncWithVpxLights(lightSeq)
        m_vpxLightSyncCollection = ColToArray(eval(lightSeq.collection))
        m_vpxLightSyncRunning = True
        m_tableSeqSpeed = Null
        m_tableSeqOffset = 0
        m_tableSeqDirection = Null
    End Sub

    Public Sub StopSyncWithVpxLights()
        m_vpxLightSyncRunning = False
        m_vpxLightSyncClear = True
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        m_tableSeqSpeed = Null
        m_tableSeqOffset = 0
        m_tableSeqDirection = Null
    End Sub

    Public Sub SetVpxSyncLightColor(color)
        m_tableSeqColor = color
    End Sub
    
    Public Sub SetVpxSyncLightsPalette(palette, direction, speed)
        m_tableSeqColor = palette
        Select Case direction:
            Case "BottomToTop": 
                m_tableSeqDirection = m_coordsY
                m_tableSeqColor = ReverseArray(palette)
            Case "TopToBottom": 
                m_tableSeqDirection = m_coordsY
            Case "RightToLeft": 
                m_tableSeqDirection = m_coordsX
            Case "LeftToRight": 
                m_tableSeqDirection = m_coordsX
                m_tableSeqColor = ReverseArray(palette)       
            Case "RadialOut": 
                m_tableSeqDirection = m_radii      
            Case "RadialIn": 
                m_tableSeqDirection = m_radii
                m_tableSeqColor = ReverseArray(palette) 
            Case "Clockwise": 
                m_tableSeqDirection = m_angles
            Case "AntiClockwise": 
                m_tableSeqDirection = m_angles
                m_tableSeqColor = ReverseArray(palette) 
        End Select  

        m_tableSeqSpeed = speed
    End Sub

    Public Sub SetTableSequenceFade(fadeUp, fadeDown)
        m_tableSeqFadeUp = fadeUp
        m_tableSeqFadeDown = fadeDown
    End Sub

    Public Function GetLightIdx(light)
        dim syncLight : syncLight = Null
        If m_lights.Exists(light.name) Then
            'found a light
            Set syncLight = m_lights(light.name)
        End If
        If Not IsNull(syncLight) Then
            'Found a light to sync.
            GetLightIdx = lightsToLeds(syncLight.Idx)
        Else
            GetLightIdx = Null
        End If
        
    End Function

    Private Function m_buildBlinkSeq(lightName, pattern)
        Dim i, buff : buff = Array()
        ReDim buff(Len(pattern)-1)
        For i = 0 To Len(pattern)-1
            
            If Mid(pattern, i+1, 1) = 1 Then
                buff(i) = lightName & "|100"
            Else
                buff(i) = lightName & "|0"
            End If
        Next
        m_buildBlinkSeq=buff
    End Function

    Private Function GetTmpLight(idx)
        If IsArray(Lights(idx) ) Then	'if array
            Set GetTmpLight = Lights(idx)(0)
        Else
            Set GetTmpLight = Lights(idx)
        End If
    End Function

    Public Sub ResetLights()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            m_lightOff(light) 
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
        RemoveAllTableLightSeqs()
        Dim k
        For Each k in m_seqRunners.Keys()
            Dim lsRunner: Set lsRunner = m_seqRunners(k)
            lsRunner.RemoveAll
        Next

    End Sub

    Public Sub PauseMainLights
        If m_pauseMainLights = False Then
            m_pauseMainLights = True
            m_pausedLights.RemoveAll
            Dim pon
            Set pon = CreateObject("Scripting.Dictionary")
            Dim poff : Set poff = CreateObject("Scripting.Dictionary")
            Dim ppulse : Set ppulse = CreateObject("Scripting.Dictionary")
            Dim pseqs : Set pseqs = CreateObject("Scripting.Dictionary")
            Dim lightProps : Set lightProps = CreateObject("Scripting.Dictionary")
            'Add State in
            Dim light, item
            For Each item in m_on.Keys()
                pon.add item, m_on(Item)
            Next
            For Each item in m_off.Keys()
                poff.add item, m_off(Item)
            Next
            For Each item in m_pulse.Keys()
                ppulse.add item, m_pulse(Item)
            Next
            For Each item in m_seqRunners.Keys()
                dim tmpSeq : Set tmpSeq = new LCSeqRunner
                dim seqItem
                For Each seqItem in m_seqRunners(Item).Items.Items()
                    tmpSeq.AddSequenceItem seqItem
                Next
                tmpSeq.CurrentItemIdx = m_seqRunners(Item).CurrentItemIdx
                pseqs.add item, tmpSeq
            Next
            
            Dim savedProps(1,3)
            
            For Each light in m_lights.Keys()
                    
                savedProps(0,0) = m_lights(light).Color
                savedProps(0,1) = m_lights(light).Level
                If m_seqs.Exists(light & "Blink") Then
                    savedProps(0,2) = m_seqs.Item(light & "Blink").UpdateInterval
                Else
                    savedProps(0,2) = Empty
                End If
                lightProps.add light, savedProps
            Next
            m_pausedLights.Add "on", pon
            m_pausedLights.Add "off", poff
            m_pausedLights.Add "pulse", ppulse
            m_pausedLights.Add "runners", pseqs
            m_pausedLights.Add "lightProps", lightProps
            m_on.RemoveAll
            m_off.RemoveAll
            m_pulse.RemoveAll
            For Each item in m_seqRunners.Items()
                item.removeAll
            Next			
        End If
    End Sub

    Public Sub ResumeMainLights
        If m_pauseMainLights = True Then
            m_pauseMainLights = False
            m_on.RemoveAll
            m_off.RemoveAll
            m_pulse.RemoveAll
            Dim light, item
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
            For Each item in m_seqRunners.Items()
                item.removeAll
            Next
            'Add State in
            For Each item in m_pausedLights("on").Keys()
                m_on.add item, m_pausedLights("on")(Item)
            Next
            For Each item in m_pausedLights("off").Keys()
                m_off.add item, m_pausedLights("off")(Item)
            Next			
            For Each item in m_pausedLights("pulse").Keys()
                m_pulse.add item, m_pausedLights("pulse")(Item)
            Next
            For Each item in m_pausedLights("runners").Keys()
                
                
                Set m_seqRunners(Item) = m_pausedLights("runners")(Item)
            Next
            For Each item in m_pausedLights("lightProps").Keys()
                LightColor Eval(Item), m_pausedLights("lightProps")(Item)(0,0)
                LightLevel Eval(Item), m_pausedLights("lightProps")(Item)(0,1)
                If Not IsEmpty(m_pausedLights("lightProps")(Item)(0,2)) Then
                    UpdateBlinkInterval Eval(Item), m_pausedLights("lightProps")(Item)(0,2)
                End If
            Next			
            m_pausedLights.RemoveAll
        End If
    End Sub

    Public Sub Update()

        m_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
        Dim x
        Dim lk
        dim color
        dim idx
        Dim lightKey
        Dim lcItem
        Dim tmpLight
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                RunLightSeq m_seqOverrideRunners(seqOverride)
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
        
            If HasKeys(m_on) Then   
                For Each lightKey in m_on.Keys()
                    Set lcItem = m_on(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, m_on(lightKey).Color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_off) Then
                For Each lightKey in m_off.Keys()
                    Set lcItem = m_off(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(0, Null, lcItem.Idx)
                Next
            End If

        
            If HasKeys(m_seqRunners) Then
                Dim k
                For Each k in m_seqRunners.Keys()
                    Dim lsRunner: Set lsRunner = m_seqRunners(k)
                    If Not IsNull(lsRunner.CurrentItem) Then
                            RunLightSeq lsRunner
                    End If
                Next
            End If

            If HasKeys(m_pulse) Then   
                For Each lightKey in m_pulse.Keys()
                    Dim pulseColor : pulseColor = m_pulse(lightKey).Color
                    If IsNull(pulseColor) Then
                        AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).Light.Color, m_pulse(lightKey).light.Idx)
                    Else
                        AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).Color, m_pulse(lightKey).light.Idx)
                    End If						
                    
                    Dim pulseUpdateInt : pulseUpdateInt = m_pulse(lightKey).interval - m_frameTime
                    Dim pulseIdx : pulseIdx = m_pulse(lightKey).idx
                    If pulseUpdateInt <= 0 Then
                        pulseUpdateInt = m_pulseInterval
                        pulseIdx = pulseIdx + 1
                    End If
                    
                    Dim pulses : pulses = m_pulse(lightKey).pulses
                    Dim pulseCount : pulseCount = m_pulse(lightKey).Cnt
                    
                    
                    If pulseIdx > UBound(m_pulse(lightKey).pulses) Then
                        m_pulse.Remove lightKey    
                        If pulseCount > 0 Then
                            pulseCount = pulseCount - 1
                            pulseIdx = 0
                            m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount, pulseColor)
                        End If
                    Else
                        m_pulse.Remove lightKey
                        m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount, pulseColor)
                    End If
                Next
            End If

            If m_vpxLightSyncRunning = True Then
                Dim lx
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lx in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncLight : syncLight = Null
                        If m_lights.Exists(lx.name) Then
                            'found a light
                            Set syncLight = m_lights(lx.name)
                        End If
                        If Not IsNull(syncLight) Then
                            'Found a light to sync.
                            
                            Dim lightState

                            If IsNull(m_tableSeqColor) Then
                                color = syncLight.Color
                            Else
                                If Not IsArray(m_tableSeqColor) Then
                                    color = Array(m_tableSeqColor, Null)
                                Else
                                    
                                    If Not IsNull(m_tableSeqSpeed) And Not m_tableSeqSpeed = 0 Then

                                        Dim colorPalleteIdx : colorPalleteIdx = IncrementUInt8(m_tableSeqDirection(lightsToLeds(syncLight.Idx)),m_tableSeqOffset)
                                        If gametime mod m_tableSeqSpeed = 0 Then
                                            m_tableSeqOffset = m_tableSeqOffset + 1
                                            If m_tableSeqOffset > 255 Then
                                                m_tableSeqOffset = 0
                                            End If	
                                        End If
                                        If colorPalleteIdx < 0 Then 
                                            colorPalleteIdx = 0
                                        End If
                                        color = Array(m_TableSeqColor(Round(colorPalleteIdx)), Null)
                                        'color = syncLight.Color
                                    Else
                                        color = Array(m_TableSeqColor(m_tableSeqDirection(lightsToLeds(syncLight.Idx))), Null)
                                    End If
                                    
                                End If
                            End If
                        
                    
                            AssignStateForFrame syncLight.name, (new FrameState)(lx.GetInPlayState*100,color, syncLight.Idx)                     
                        End If
                    Next
                End If
            End If

            If m_vpxLightSyncClear = True Then  
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lk in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncClearLight : syncClearLight = Null
                        If m_lights.Exists(lk.name) Then
                            'found a light
                            Set syncClearLight = m_lights(lk.name)
                        End If
                        If Not IsNull(syncClearLight) Then
                            AssignStateForFrame syncClearLight.name, (new FrameState)(0, Null, syncClearLight.idx) 
                        End If
                    Next
                End If
            
                m_vpxLightSyncClear = False
            End If
        End If
        

        If HasKeys(m_currentFrameState) Then
            
            Dim frameStateKey
            For Each frameStateKey in m_currentFrameState.Keys()
                idx = m_currentFrameState(frameStateKey).idx
                
                Dim newColor : newColor = m_currentFrameState(frameStateKey).colors
                Dim bUpdate

                If Not IsNull(newColor) Then
                    'Check current color is the new color coming in, if not, set the new color.
                    
                    Set tmpLight = GetTmpLight(idx)

                    Dim c, cf
                    c = newColor(0)
                    cf= newColor(1)

                    If Not IsNull(c) Then
                        If Not CStr(tmpLight.Color) = CStr(c) Then
                            bUpdate = True
                        End If
                    End If

                    If Not IsNull(cf) Then
                        If Not CStr(tmpLight.ColorFull) = CStr(cf) Then
                            bUpdate = True
                        End If
                    End If
                End If

               
                Dim lm
                If IsArray(Lights(idx)) Then
                    For Each x in Lights(idx)
                        If bUpdate Then 
                            If Not IsNull(c) Then
                                x.color = c
                            End If
                            If Not IsNull(cf) Then
                                x.colorFull = cf
                            End If
                            If m_lightmaps.Exists(x.Name) Then
                                For Each lm in m_lightmaps(x.Name)
                                    If Not IsNull(lm) Then
                                        lm.Color = c
                                    End If
                                Next
                            End If
                        End If
                        x.State = m_currentFrameState(frameStateKey).level/100
                    Next
                Else
                    If bUpdate Then    
                        If Not IsNull(c) Then
                            Lights(idx).color = c
                        End If
                        If Not IsNull(cf) Then
                            Lights(idx).colorFull = cf
                        End If
                        If m_lightmaps.Exists(Lights(idx).Name) Then
                            For Each lm in m_lightmaps(Lights(idx).Name)
                                If Not IsNull(lm) Then
                                    lm.Color = c
                                End If
                            Next
                        End If
                    End If
                    Lights(idx).State = m_currentFrameState(frameStateKey).level/100
                End If
            Next
        End If
        m_currentFrameState.RemoveAll
        m_off.RemoveAll

    End Sub

    Private Function HexToInt(hex)
        HexToInt = CInt("&H" & hex)
    End Function

    Function RGBToHex(r, g, b)
        RGBToHex = Right("0" & Hex(r), 2) & _
            Right("0" & Hex(g), 2) & _
            Right("0" & Hex(b), 2)
    End Function

    Function FadeRGB(light, color1, color2, steps)

    
        Dim r1, g1, b1, r2, g2, b2
        Dim i
        Dim r, g, b
        color1 = clng(color1)
        color2 = clng(color2)
        ' Extract RGB values from the color integers
        r1 = color1 Mod 256
        g1 = (color1 \ 256) Mod 256
        b1 = (color1 \ (256 * 256)) Mod 256

        r2 = color2 Mod 256
        g2 = (color2 \ 256) Mod 256
        b2 = (color2 \ (256 * 256)) Mod 256

        ' Resize the output array
        ReDim outputArray(steps - 1)

        ' Generate the fade
        For i = 0 To steps - 1
            ' Calculate RGB values for this step
            r = r1 + (r2 - r1) * i / (steps - 1)
            g = g1 + (g2 - g1) * i / (steps - 1)
            b = b1 + (b2 - b1) * i / (steps - 1)

            ' Convert RGB to hex and add to output
            outputArray(i) = light & "|100|" & RGBToHex(CInt(r), CInt(g), CInt(b))
        Next
        FadeRGB = outputArray
    End Function

    Public Function CreateColorPalette(startColor, endColor, steps)
        Dim colors()
        ReDim colors(steps)
        
        Dim startRed, startGreen, startBlue, endRed, endGreen, endBlue
        startRed = HexToInt(Left(startColor, 2))
        startGreen = HexToInt(Mid(startColor, 3, 2))
        startBlue = HexToInt(Right(startColor, 2))
        endRed = HexToInt(Left(endColor, 2))
        endGreen = HexToInt(Mid(endColor, 3, 2))
        endBlue = HexToInt(Right(endColor, 2))
        
        Dim redDiff, greenDiff, blueDiff
        redDiff = endRed - startRed
        greenDiff = endGreen - startGreen
        blueDiff = endBlue - startBlue
        
        Dim i
        For i = 0 To steps
            Dim red, green, blue
            red = startRed + (redDiff * (i / steps))
            green = startGreen + (greenDiff * (i / steps))
            blue = startBlue + (blueDiff * (i / steps))
            colors(i) = RGB(red,green,blue)'IntToHex(red, 2) & IntToHex(green, 2) & IntToHex(blue, 2)
        Next
        
        CreateColorPalette = colors
    End Function

    Function CreateColorPaletteWithStops(startColor, endColor, stopPositions, stopColors)

        Dim colors(255)

        Dim fStop : fStop = CreateColorPalette(startColor, stopColors(0), stopPositions(0))
        Dim i, istep
        For i = 0 to stopPositions(0)
            colors(i) = fStop(i)
        Next
        For i = 1 to Ubound(stopColors)
            Dim stopStep : stopStep = CreateColorPalette(stopColors(i-1), stopColors(i), stopPositions(i))
            Dim ii
        ' MsgBox(stopPositions(i) - stopPositions(i-1))
            istep = 0
            For ii = stopPositions(i-1)+1 to stopPositions(i)
            '  MsgBox(ii)
            colors(ii) = stopStep(iStep)
            iStep = iStep + 1
            Next
        Next
        ' MsgBox("Here")
        Dim eStop : eStop = CreateColorPalette(stopColors(UBound(stopColors)), endColor, 255-stopPositions(UBound(stopPositions)))
        'MsgBox(UBound(eStop))
        iStep = 0
        For i = 255-(255-stopPositions(UBound(stopPositions))) to 254
            colors(i) = eStop(iStep)
            iStep = iStep + 1
        Next

        CreateColorPaletteWithStops = colors
    End Function

    Private Function HasKeys(o)
        If Ubound(o.Keys())>-1 Then
            HasKeys = True
        Else
            HasKeys = False
        End If
    End Function

    Private Sub RunLightSeq(seqRunner)

        Dim lcSeq: Set lcSeq = seqRunner.CurrentItem
        dim lsName, isSeqEnd
        If UBound(lcSeq.Sequence)<lcSeq.CurrentIdx Then
            isSeqEnd = True
        Else
            isSeqEnd = False
        End If

        dim lightInSeq
        For each lightInSeq in lcSeq.LightsInSeq
        
            If isSeqEnd Then

                

            'Needs a guard here for something, but i've forgotten. 
            'I remember: Only reset the light if there isn't frame data for the light. 
            'e.g. a previous seq has affected the light, we don't want to clear that here on this frame
                If m_lights.Exists(lightInSeq) = True AND NOT m_currentFrameState.Exists(lightInSeq) Then
                    AssignStateForFrame lightInSeq, (new FrameState)(0, Null, m_lights(lightInSeq).Idx)
                End If
            Else
                


                If m_currentFrameState.Exists(lightInSeq) Then

                    
                    'already frame data for this light.
                    'replace with the last known state from this seq
                    If Not IsNull(lcSeq.LastLightState(lightInSeq)) Then
                        AssignStateForFrame lightInSeq, lcSeq.LastLightState(lightInSeq)
                    End If
                End If

            End If
        Next

        If isSeqEnd Then
            lcSeq.CurrentIdx = 0
            seqRunner.NextItem()
        End If

        If Not IsNull(seqRunner.CurrentItem) Then
            Dim framesRemaining, seq, color
            Set lcSeq = seqRunner.CurrentItem
            seq = lcSeq.Sequence
            

            Dim name
            Dim ls, x
            If IsArray(seq(lcSeq.CurrentIdx)) Then
                For x = 0 To UBound(seq(lcSeq.CurrentIdx))
                    lsName = Split(seq(lcSeq.CurrentIdx)(x),"|")
                    name = lsName(0)
                    If m_lights.Exists(name) Then
                        Set ls = m_lights(name)
                        
                        color = lcSeq.Color

                        If IsNull(color) Then
                            color = ls.Color
                        End If
                        
                        If Ubound(lsName) = 2 Then
                            If lsName(2) = "" Then
                                AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                            Else
                                AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                            End If
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        End If
                        lcSeq.SetLastLightState name, m_currentFrameState(name) 
                    End If
                Next       
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)
                    
                    color = lcSeq.Color
                    If IsNull(color) Then
                        color = ls.Color
                    End If
                    If Ubound(lsName) = 2 Then
                        If lsName(2) = "" Then
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                        End If
                    Else
                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                    End If
                    lcSeq.SetLastLightState name, m_currentFrameState(name) 
                End If
            End If

            framesRemaining = lcSeq.Update(m_frameTime)
            If framesRemaining < 0 Then
                lcSeq.ResetInterval()
                lcSeq.NextFrame()
            End If
            
        End If
    End Sub

End Class

Class FrameState
    Private m_level, m_colors, m_idx

    Public Property Get Level(): Level = m_level: End Property
    Public Property Let Level(input): m_level = input: End Property

    Public Property Get Colors(): Colors = m_colors: End Property
    Public Property Let Colors(input): m_colors = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public default function init(level, colors, idx)
        m_level = level
        m_colors = colors
        m_idx = idx 

        Set Init = Me
    End Function

    Public Function ColorAt(idx)
        ColorAt = m_colors(idx) 
    End Function
End Class

Class PulseState
    Private m_light, m_pulses, m_idx, m_interval, m_cnt, m_color

    Public Property Get Light(): Set Light = m_light: End Property
    Public Property Let Light(input): Set m_light = input: End Property

    Public Property Get Pulses(): Pulses = m_pulses: End Property
    Public Property Let Pulses(input): m_pulses = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public Property Get Interval(): Interval = m_interval: End Property
    Public Property Let Interval(input): m_interval = input: End Property

    Public Property Get Cnt(): Cnt = m_cnt: End Property
    Public Property Let Cnt(input): m_cnt = input: End Property

    Public Property Get Color(): Color = m_color: End Property
    Public Property Let Color(input): m_color = input: End Property		

    Public default function init(light, pulses, idx, interval, cnt, color)
        Set m_light = light
        m_pulses = pulses
        'debug.Print(Join(Pulses))
        m_idx = idx 
        m_interval = interval
        m_cnt = cnt
        m_color = color

        Set Init = Me
    End Function

    Public Function PulseAt(idx)
        PulseAt = m_pulses(idx) 
    End Function
End Class

Class LCItem
    
    Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y, m_baseColor

        Public Property Get Idx()
            Idx=m_Idx
        End Property

        Public Property Get Color()
            Color=m_color
        End Property

        Public Property Get BaseColor()
            BaseColor=m_baseColor
        End Property

        Public Property Let Color(input)
            If IsNull(input) Then
                m_Color = Null
            Else
                If Not IsArray(input) Then
                    input = Array(input, null)
                End If
                m_Color = input
            End If
        End Property

        Public Property Let Level(input)
            m_level = input
        End Property

        Public Property Get Level()
            Level=m_level
        End Property

        Public Property Get Name()
            Name=m_name
        End Property

        Public Property Get X()
            X=m_x
        End Property

        Public Property Get Y()
            Y=m_y
        End Property

        Public Property Get Row()
            Row=Round(m_x/20)
        End Property

        Public Property Get Col()
            Col=Round(m_y/20)
        End Property

        Public Sub Init(idx, intervalMs, color, name, x, y)
            m_Idx = idx
            If Not IsArray(color) Then
                m_color = Array(color, null)
            Else
                m_color = color
            End If
            m_baseColor = m_color
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
        End Sub

End Class

Class LCSeq
    
    Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates, m_palette, m_pattern, m_tokens, m_loops

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
        m_lastLightStates.RemoveAll()
        m_currentIdx = input
    End Property

    Public Property Get Tokens()
        Tokens=m_tokens
    End Property

    Public Property Let Tokens(input)
        Set m_tokens = input
    End Property    

    Public Property Get LightsInSeq()
        LightsInSeq=m_lightsInSeq.Keys()
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property
    
    Public Property Let Sequence(input)
        Dim item, light, lightItem, token
        m_lightsInSeq.RemoveAll
        Dim i, x
    
        ' Iterate through the input array
        For i = 0 To UBound(input)
            item = input(i)
            If IsArray(item) Then
                For x = 0 To UBound(item)
                    light = item(x)
                    lightItem = Split(light, "|")
                    token = IsToken(lightItem(0))
                    If Not IsNull(token) Then
                        lightItem(0) = m_tokens(token)
                    End If
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If
                    light = Join(lightItem, "|")
                    item(x) = light
                    Debug.Print(light)
                Next
                ' Update the input array with modified item array
                input(i) = item
            Else
                lightItem = Split(item, "|")
                token = IsToken(lightItem(0))
                If Not IsNull(token) Then
                    lightItem(0) = m_tokens(token)
                End If
                If Not m_lightsInSeq.Exists(lightItem(0)) Then
                    m_lightsInSeq.Add lightItem(0), True
                End If
                light = Join(lightItem, "|")
                ' Update the input array with modified item string
                input(i) = light
            End If
        Next
        ' Assign the modified input array to m_sequence
        m_sequence = input
    End Property

    Private Function IsToken(mainString)
        ' Check if the string contains an opening parenthesis and ends with a closing parenthesis
        If InStr(mainString, "(") > 0 And Right(mainString, 1) = ")" Then
            ' Extract the substring within the parentheses
            Dim startPos, subString
            startPos = InStr(mainString, "(")
            subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
            IsToken = subString
        Else
            IsToken = Null
        End If
    End Function
    

    Public Property Get LastLightState(light)
        If m_lastLightStates.Exists(light) Then
            dim c : Set c = m_lastLightStates(light)
            Set LastLightState = c
        Else
            LastLightState = Null
        End If
    End Property

    Public Property Let LastLightState(light, input)
        If m_lastLightStates.Exists(light) Then
            m_lastLightStates.Remove light
        End If
        If input.level > 0 Then
            m_lastLightStates.Add light, input
        End If
    End Property

    Public Sub SetLastLightState(light, input)	
        If m_lastLightStates.Exists(light) Then	
            m_lastLightStates.Remove light	
        End If	
        If input.level > 0 Then	
                m_lastLightStates.Add light, input	
        End If	
    End Sub

    Public Property Get Color()
        Color=m_color
    End Property
    
    Public Property Let Color(input)
        If IsNull(input) Then
            m_Color = Null
        Else
            If Not IsArray(input) Then
                input = Array(input, null)
            End If
            m_Color = input
        End If
    End Property

    Public Property Get Palette()
        Palette=m_palette
    End Property
    
    Public Property Let Palette(input)
        If IsNull(input) Then
            m_palette = Null
        Else
            If Not IsArray(input) Then
                m_palette = Null
            Else
                m_palette = input
            End If
        End If
    End Property

    Public Property Get Name()
        Name=m_name
    End Property
    
    Public Property Let Name(input)
        m_name = input
    End Property        

    Public Property Get UpdateInterval()
        UpdateInterval=m_updateInterval
    End Property

    Public Property Let UpdateInterval(input)
        m_updateInterval = input
        'm_Frames = input
    End Property

    Public Property Get Repeat()
        Repeat=m_repeat
    End Property

    Public Property Let Repeat(input)
        m_repeat = input
    End Property

    Public Property Get Loops()
        Loops=m_loops
    End Property

    Public Property Let Loops(input)
        m_loops = input
    End Property

    Public Property Get Pattern()
        Pattern=m_pattern
    End Property

    Public Property Let Pattern(input)
        m_pattern = input
    End Property    

    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Array(Null, Null)
        m_updateInterval = 180
        m_repeat = False
        m_loops = 1
        m_Frames = 180
        m_pattern = Null
        Set m_lightsInSeq = CreateObject("Scripting.Dictionary")
        Set m_lastLightStates = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get Update(framesPassed)
        m_Frames = m_Frames - framesPassed
        Update = m_Frames
    End Property

    Public Sub NextFrame()
        m_currentIdx = m_currentIdx + 1
    End Sub

    Public Sub ResetInterval()
        m_Frames = m_updateInterval
        Exit Sub
    End Sub

End Class

Class LCSeqRunner
    
    Private m_name, m_items, m_currentItemIdx

    Public Property Get Name()
        Name=m_name
    End Property
    
    Public Property Let Name(input)
        m_name = input
    End Property

    Public Property Get Items()
        Set Items = m_items
    End Property

    Public Property Get ItemByKey(key)
        Set Items = m_items(key)
    End Property

    Public Property Get CurrentItemIdx()
        CurrentItemIdx = m_currentItemIdx
    End Property

    Public Property Let CurrentItemIdx(input)
        m_currentItemIdx = input
    End Property

    Public Property Get CurrentItem()
        Dim items: items = m_items.Items()
        If m_currentItemIdx > UBound(items) Then
            m_currentItemIdx = 0
        End If
        If UBound(items) = -1 Then       
            CurrentItem  = Null
        Else
            Set CurrentItem = items(m_currentItemIdx)                
        End If
    End Property

    Private Sub Class_Initialize()    
        Set m_items = CreateObject("Scripting.Dictionary")
        m_currentItemIdx = 0
    End Sub

    Public Sub AddItem(key, sequence, loops, speed, tokens)
        If Not IsNull(sequence) Then

            Dim lSeq : Set lSeq = New LCSeq
            lSeq.Name = key
            lSeq.Tokens = tokens
            lSeq.Sequence = sequence
            lSeq.UpdateInterval = speed
            lSeq.Loops = loops

            If Not m_items.Exists(key) Then
                m_items.Add key, lSeq
            End If
        End If
    End Sub

    Public Sub AddSequenceItem(sequence)
        If Not IsNull(sequence) Then
            Dim loops
            If sequence.Repeat = True Then
                sequence.Loops = -1
            Else
                sequence.Loops = 1
            End If
            If Not m_items.Exists(sequence.Name) Then
                m_items.Add sequence.Name, sequence
            End If
        End If
    End Sub

    Public Sub RemoveAll()
        Dim item
        For Each item in m_items.Keys()
            m_items(item).ResetInterval
            m_items(item).CurrentIdx = 0
            m_items.Remove item
        Next
    End Sub

    Public Sub RemoveItem(key)
        If m_items.Exists(key) Then
                m_items(key).ResetInterval
                m_items(key).CurrentIdx = 0
                m_items.Remove key
        End If
    End Sub

    Public Sub NextItem()
        Dim items: items = m_items.Items
        Dim keys : keys = m_items.Keys
        If items(m_currentItemIdx).Loops = 1 Then
            RemoveItem(keys(m_currentItemIdx))
        Else
            If items(m_currentItemIdx).Loops > 1 Then
                items(keys(m_currentItemIdx)).Loops = items(keys(m_currentItemIdx)).Loops - 1
            End If
            m_currentItemIdx = m_currentItemIdx + 1
        End If
        
        If m_currentItemIdx > UBound(m_items.Items) Then   
            m_currentItemIdx = 0
        End If
    End Sub

    Public Function HasSeq(key)
        If m_items.Exists(key) Then
            HasSeq = True
        Else
            HasSeq = False
        End If
    End Function

End Class

'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Class DebugLogFile
	Private Filename
	Private TxtFileStream

	Public default Function init()
        Filename = cGameName + "_" & GetTimeStamp & "_debug_log.txt"
	  Set Init = Me
	End Function
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
		GetTimeStamp = _
		LZ(Year(CurrTime),   4) & "-" _
		 & LZ(Month(CurrTime),  2) & "-" _
		 & LZ(Day(CurrTime),	2) & "_" _
		 & LZ(Hour(CurrTime),   2) & "" _
		 & LZ(Minute(CurrTime), 2) & "" _
		 & LZ(Second(CurrTime), 2) & "" _
		 & LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message)
		If debugEnabled = True Then
			Dim FormattedMsg, Timestamp
			
			Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, 8, True)
			Timestamp = GetTimeStamp
			FormattedMsg = GetTimeStamp + ": " + label + ": " + message
			TxtFileStream.WriteLine FormattedMsg
			TxtFileStream.Close
			Debug.print label & ": " & message
		End If
	End Sub
End Class

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************


Dim glf_lastPinEvent : glf_lastPinEvent = Null

Sub DispatchPinEvent(e, kwargs)
    If Not pinEvents.Exists(e) Then
        debugLog.WriteToLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = pinEvents(e)
    debugLog.WriteToLog "DispatchPinEvent", e
    For Each k In pinEventsOrder(e)
        'debugLog.WriteToLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs))
    Next
End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not pinEvents.Exists(e) Then
        'debugLog.WriteToLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = pinEvents(e)
    'debugLog.WriteToLog "DispatchReplayPinEvent", e
    For Each k In pinEventsOrder(e)
        'debugLog.WriteToLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        kwargs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs))
    Next
    DispatchRelayPinEvent = kwargs
End Function

Sub AddPinEventListener(evt, key, callbackName, priority, args)
    Dim i, inserted, tempArray
    If Not pinEvents.Exists(evt) Then
        pinEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not pinEvents(evt).Exists(key) Then
        pinEvents(evt).Add key, Array(callbackName, priority, args)
        SortPinEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePinEventListener(evt, key)
    If pinEvents.Exists(evt) Then
        If pinEvents(evt).Exists(key) Then
            pinEvents(evt).Remove key
            SortPinEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub SortPinEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the pinEventsOrder to maintain order based on priority
    If Not pinEventsOrder.Exists(evt) Then
        ' If the event does not exist in pinEventsOrder, just add it directly if we're adding
        If isAdding Then
            pinEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = pinEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the pinEventsOrder with the newly ordered or modified list
        pinEventsOrder(evt) = tempArray
    End If
End Sub
Function GetPlayerState(key)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If playerState(currentPlayer).Exists(key)  Then
        GetPlayerState = playerState(currentPlayer)(key)
    Else
        GetPlayerState = Null
    End If
End Function

Function GetPlayerScore(player)
    dim p
    Select Case player
        Case 1:
            p = "PLAYER 1"
        Case 2:
            p = "PLAYER 2"
        Case 3:
            p = "PLAYER 3"
        Case 4:
            p = "PLAYER 4"
    End Select

    If playerState.Exists(p) Then
        GetPlayerScore = playerState(p)(SCORE)
    Else
        GetPlayerScore = 0
    End If
End Function

Function GetCurrentPlayerNumber()
    Select Case currentPlayer
        Case "PLAYER 1":
            GetCurrentPlayerNumber = 1
        Case "PLAYER 2":
            GetCurrentPlayerNumber = 2
        Case "PLAYER 3":
            GetCurrentPlayerNumber = 3
        Case "PLAYER 4":
            GetCurrentPlayerNumber = 4
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If IsArray(value) Then
        If Join(GetPlayerState(key)) = Join(value) Then
            Exit Function
        End If
    Else
        If GetPlayerState(key) = value Then
            Exit Function
        End If
    End If   
    Dim prevValue
    If playerState(currentPlayer).Exists(key) Then
        prevValue = playerState(currentPlayer)(key)
        playerState(currentPlayer).Remove key
    End If
    playerState(currentPlayer).Add key, value
    
    If playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue
    End If
    
    SetPlayerState = Null
End Function

Sub FirePlayerEventHandlers(evt, value, prevValue)
    If Not playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = playerEvents(evt)
    For Each k In playerEventsOrder(evt)
        GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), Array(evt,value,prevValue)))
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, callbackName, priority, args)
    If Not playerEvents.Exists(evt) Then
        playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not playerEvents(evt).Exists(key) Then
        playerEvents(evt).Add key, Array(callbackName, priority, args)
        SortPlayerEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePlayerStateEventListener(evt, key)
    If playerEvents.Exists(evt) Then
        If playerEvents(evt).Exists(key) Then
            playerEvents(evt).Remove key
            SortPlayerEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub SortPlayerEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the playerEventsOrder to maintain order based on priority
    If Not playerEventsOrder.Exists(evt) Then
        ' If the event does not exist in playerEventsOrder, just add it directly if we're adding
        If isAdding Then
            playerEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = playerEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the playerEventsOrder with the newly ordered or modified list
        playerEventsOrder(evt) = tempArray
    End If
End Sub

Sub EmitAllPlayerEvents()
    Dim key
    For Each key in playerState(currentPlayer).Keys()
        FirePlayerEventHandlers key, GetPlayerState(key), GetPlayerState(key)
    Next
End Sub

'*******************************************
'  Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this, 
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to 
'	 the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 100 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


Dim ballInReleasePostion : ballInReleasePostion = False
'TROUGH 
Sub swTrough1_Hit
	ballInReleasePostion = True
	UpdateTrough
End Sub
Sub swTrough1_UnHit
	ballInReleasePostion = False
	UpdateTrough
End Sub
Sub swTrough2_Hit
	UpdateTrough
End Sub
Sub swTrough2_UnHit
	UpdateTrough
End Sub
Sub swTrough3_Hit
	UpdateTrough
End Sub
Sub swTrough3_UnHit
	UpdateTrough
End Sub
Sub swTrough4_Hit
	UpdateTrough
End Sub
Sub swTrough4_UnHit
	UpdateTrough
End Sub
Sub swTrough5_Hit
	UpdateTrough
End Sub
Sub swTrough5_UnHit
	UpdateTrough
End Sub
Sub swTrough6_Hit
	UpdateTrough
End Sub
Sub swTrough6_UnHit
	UpdateTrough
End Sub
Sub swTrough7_Hit
	UpdateTrough
End Sub
Sub swTrough7_UnHit
	UpdateTrough
End Sub
Sub swTrough8_Hit
	UpdateTrough
    If gameStarted = True Then
        glf_BIP = glf_BIP - 1
        DispatchRelayPinEvent GLF_BALL_DRAIN, 1
    End If
End Sub
Sub swTrough8_UnHit
	UpdateTrough
End Sub

Sub UpdateTrough
	SetDelay "update_trough", "UpdateTroughDebounced", Null, 100
End Sub

Sub UpdateTroughDebounced(args)
	If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
    If swTrough5.BallCntOver = 0 Then swTrough6.kick 57, 10
    If swTrough6.BallCntOver = 0 Then swTrough7.kick 57, 10
    If swTrough7.BallCntOver = 0 Then swTrough8.kick 57, 10
End Sub
