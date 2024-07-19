'VPX Game Logic Framework (https://mpcarr.github.io/vpx-gle-framework/)

'
Dim glf_currentPlayer : glf_currentPlayer = Null
Dim glf_canAddPlayers : glf_canAddPlayers = True
Dim glf_PI : glf_PI = 4 * Atn(1)
Dim glf_plunger
Dim glf_gameStarted : glf_gameStarted = False
Dim glf_pinEvents : Set glf_pinEvents = CreateObject("Scripting.Dictionary")
Dim glf_pinEventsOrder : Set glf_pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim glf_playerEvents : Set glf_playerEvents = CreateObject("Scripting.Dictionary")
Dim Glf_EventBlocks : Set Glf_EventBlocks = CreateObject("Scripting.Dictionary")
Dim Glf_ShotProfiles : Set Glf_ShotProfiles = CreateObject("Scripting.Dictionary")
Dim Glf_ShowStartQueue : Set Glf_ShowStartQueue = CreateObject("Scripting.Dictionary")
Dim glf_playerEventsOrder : Set glf_playerEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")
Dim glf_Modes : Set glf_Modes = CreateObject("Scripting.Dictionary")
Dim bcpController : bcpController = Null
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = ""
Dim lightCtrl : Set lightCtrl = new LStateController
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0

Dim glf_ballsPerGame : glf_ballsPerGame = 3
Dim glf_troughSize : glf_troughSize = tnob

Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False

lightCtrl.RegisterLights Glf_Lights
lightCtrl.Debug = False
Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController
    Set bcpController = (new GlfVpxBcpController)(bcpPort	, bcpExeName)
End Sub

Public Sub Glf_Init()
	If useBCP = True Then
		Glf_ConnectToBCPMediaController
	End If

	swTrough1.DestroyBall
	swTrough2.DestroyBall
	swTrough3.DestroyBall
	swTrough4.DestroyBall
	swTrough5.DestroyBall
	swTrough6.DestroyBall
	swTrough7.DestroyBall
	swTrough8.DestroyBall
	If glf_troughSize > 0 Then : Set glf_ball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1) : End If
	If glf_troughSize > 1 Then : Set glf_ball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2) : End If
	If glf_troughSize > 2 Then : Set glf_ball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3) : End If
	If glf_troughSize > 3 Then : Set glf_ball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4) : End If
	If glf_troughSize > 4 Then : Set glf_ball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5) : End If
	If glf_troughSize > 5 Then : Set glf_ball6 = swTrough6.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6) : End If
	If glf_troughSize > 6 Then : Set glf_ball7 = swTrough7.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7) : End If
	If glf_troughSize > 7 Then : Set glf_ball8 = swTrough8.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8) : End If
	
	Dim switch, switchHitSubs
	switchHitSubs = ""
	For Each switch in Glf_Switches
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_Hit() : DispatchPinEvent """ & switch.Name & "_active"", ActiveBall : End Sub" & vbCrLf
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_UnHit() : DispatchPinEvent """ & switch.Name & "_inactive"", ActiveBall : End Sub" & vbCrLf
	Next
	ExecuteGlobal switchHitSubs
End Sub

Sub Glf_Options(ByVal eventId)
	Dim ballsPerGame : ballsPerGame = Table1.Option("Balls Per Game", 1, 2, 1, 1, 0, Array("3 Balls", "5 Balls"))
	If ballsPerGame = 1 Then
		glf_ballsPerGame = 3
	Else
		glf_ballsPerGame = 5
	End If

	Dim glfDebug : glfDebug = Table1.Option("Glf Debug Log", 0, 1, 1, 1, 0, Array("Off", "On"))
	If glfDebug = 1 Then
		glf_debugEnabled = True
	Else
		glf_debugEnabled = False
	End If

	Dim glfuseBCP : glfuseBCP = Table1.Option("Glf Backbox Control Protocol", 0, 1, 1, 1, 0, Array("Off", "On"))
	If glfuseBCP = 1 Then
		useBCP = True
		If Not IsNull(bcpController) Then
			bcpController.Disconnect
			bcpController = Null
		End If
		Glf_ConnectToBCPMediaController
	Else
		useBCP = False
		If Not IsNull(bcpController) Then
			bcpController.Disconnect
			bcpController = Null
		End If
	End If
End Sub

Public Sub Glf_Exit()
	If Not IsNull(bcpController) Then
		bcpController.Disconnect
		bcpController = Null
	End If
End Sub

Public Sub Glf_KeyDown(ByVal keycode)
    If glf_gameStarted = True Then
		If keycode = LeftFlipperKey Then
			DispatchPinEvent "s_left_flipper_active", Null
		End If
		
		If keycode = RightFlipperKey Then
			DispatchPinEvent "s_right_flipper_active", Null
		End If
		
		If keycode = StartGameKey Then
			If glf_canAddPlayers = True Then
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
	If glf_gameStarted = True Then
		If KeyCode = PlungerKey Then
			Plunger.Fire
		End If

		If keycode = LeftFlipperKey Then
			DispatchPinEvent "s_left_flipper_inactive", Null
		End If
		
		If keycode = RightFlipperKey Then
			DispatchPinEvent "s_right_flipper_inactive", Null
		End If
	End If
End Sub

Public Sub Glf_GameTimer_Timer()
	
End Sub

Public Sub Glf_EventTimer_Timer()
	DelayTick
End Sub

Public Sub Glf_BCPUpdateTimer_Timer()
	Glf_BcpUpdate
End Sub

Public Sub Glf_LampTimer_Timer()
	lightCtrl.Update
End Sub

Public Function Glf_ParseInput(value, isTime)
	Dim templateCode : templateCode = ""
	Dim tmp: tmp = value
    Select Case VarType(value)
        Case 8 ' vbString
			tmp = Glf_ReplaceCurrentPlayerAttributes(tmp)
			If InStr(tmp, " if ") Then
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & Glf_ConvertIf(tmp, "Glf_" & glf_FuncCount) & vbCrLf
				templateCode = templateCode & "End Function"
			Else
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
				templateCode = templateCode & "End Function"
			End IF
        Case Else
			templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
			If isTime Then
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp * 1000 & vbCrLf
			Else
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
			End If
			templateCode = templateCode & "End Function"
    End Select
	'msgbox templateCode
	ExecuteGlobal templateCode
	Dim funcRef : funcRef = "Glf_" & glf_FuncCount
	glf_FuncCount = glf_FuncCount + 1
	Glf_ParseInput = Array(funcRef, value)
End Function

Public Function Glf_ParseEventInput(value)
	Dim templateCode : templateCode = ""
	Dim condition : condition = Glf_IsCondition(value)
	If IsNull(condition) Then
		Glf_ParseEventInput = Array(value, value, Null)
	Else
		dim conditionReplaced : conditionReplaced = Glf_ReplaceCurrentPlayerAttributes(condition)
		templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
		templateCode = templateCode & vbTab & Glf_ConvertCondition(conditionReplaced, "Glf_" & glf_FuncCount) & vbCrLf
		templateCode = templateCode & "End Function"
		'msgbox templateCode
		ExecuteGlobal templateCode
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		glf_FuncCount = glf_FuncCount + 1
		Glf_ParseEventInput = Array(Replace(value, "{"&condition&"}", funcRef) ,Replace(value, "{"&condition&"}", ""), funcRef)
	End If
End Function

Function Glf_ReplaceCurrentPlayerAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "current_player\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    replacement = "GetPlayerState(""$1"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceCurrentPlayerAttributes = outputString
End Function

Function Glf_ConvertIf(value, retName)
    Dim parts, condition, truePart, falsePart
    parts = Split(value, " if ")
    truePart = Trim(parts(0))
    Dim conditionAndFalsePart
    conditionAndFalsePart = Split(parts(1), " else ")
    condition = Trim(conditionAndFalsePart(0))
    falsePart = Trim(conditionAndFalsePart(1))
    Dim vbscriptIfStatement
    vbscriptIfStatement = "If " & condition & " Then" & vbCrLf & _
                          "    "&retName&" = " & truePart & vbCrLf & _
                          "Else" & vbCrLf & _
                          "    "&retName&" = " & falsePart & vbCrLf & _
                          "End If"
	Glf_ConvertIf = vbscriptIfStatement
End Function

Function Glf_ConvertCondition(value, retName)
	value = Replace(value, "==", "=")
	value = Replace(value, "!=", "<>")
	Glf_ConvertCondition = "    "&retName&" = " & value
End Function



Function GlfShotProfiles(name)
	If Glf_ShotProfiles.Exists(name) Then
		Set GlfShotProfiles = Glf_ShotProfiles(name)
	Else
		Dim new_shotprofile : Set new_shotprofile = (new GlfShotProfile)(name)
		Glf_ShotProfiles.Add name, new_shotprofile
		Set GlfShotProfiles = new_shotprofile
	End If
End Function

Function CreateGlfMode(name, priority)
	If Not glf_Modes.Exists(name) Then 
		Dim mode : Set mode = (new Mode)(name, priority)
		glf_Modes.Add name, mode
		Set CreateGlfMode = mode
	End If
End Function

Function GlfModes(name)
	If glf_Modes.Exists(name) Then 
		Set GlfModes = glf_Modes(name)
	Else
		GlfModes = Null
	End If
End Function

Function GlfKwargs()
	Set GlfKwargs = CreateObject("Scripting.Dictionary")
End Function

Function Glf_ConvertShow(show, tokens)

	Dim showStep, light, lightsCount, x,tagLight, tagLights, lightParts, token, stepIdx
	Dim newShow
	ReDim newShow(UBound(show))
	stepIdx = 0
	For Each showStep in show
		lightsCount = 0 
		For Each light in showStep
			lightParts = Split(light, "|")
			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
					tagLights = lightCtrl.GetLightsForTag(lightParts(0))
					lightsCount = UBound(tagLights)+1
				Else
					If IsNull(token) Then
						lightsCount = lightsCount + 1
					Else
						'resolve token lights
						If IsNull(lightCtrl.GetLightIdx(tokens(token))) Then
							'token is a tag
							tagLights = lightCtrl.GetLightsForTag(tokens(token))
							lightsCount = UBound(tagLights)+1
						Else
							lightsCount = lightsCount + 1
						End If
					End If
				End If
			End If
		Next
	
		Dim seqArray
		ReDim seqArray(lightsCount-1)
		x=0
		For Each light in showStep
			lightParts = Split(light, "|")
			Dim lightColor : lightColor = ""
			If Ubound(lightParts) = 2 Then 
				If IsNull(Glf_IsToken(lightParts(2))) Then
					lightColor = lightParts(2)
				Else
					lightColor = tokens(Glf_IsToken(lightParts(2)))
				End If
			End If

			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
					tagLights = lightCtrl.GetLightsForTag(lightParts(0))
					For Each tagLight in tagLights
						If UBound(lightParts) >=1 Then
							seqArray(x) = tagLight & "|"&lightParts(1)&"|"&lightColor
						Else
							seqArray(x) = tagLight & "|"&lightParts(1)
						End If
						x=x+1
					Next
				Else
					If IsNull(token) Then
						If UBound(lightParts) >= 1 Then
							seqArray(x) = lightParts(0) & "|"&lightParts(1)&"|"&lightColor
						Else
							seqArray(x) = lightParts(0) & "|"&lightParts(1)
						End If
						x=x+1
					Else
						'resolve token lights
						If IsNull(lightCtrl.GetLightIdx(tokens(token))) Then
							'token is a tag
							tagLights = lightCtrl.GetLightsForTag(tokens(token))
							For Each tagLight in tagLights
								If UBound(lightParts) >=1 Then
									seqArray(x) = tagLight & "|"&lightParts(1)&"|"&lightColor
								Else
									seqArray(x) = tagLight & "|"&lightParts(1)
								End If
								x=x+1
							Next
						Else
							If UBound(lightParts) >= 1 Then
								seqArray(x) = tokens(token) & "|"&lightParts(1)&"|"&lightColor
							Else
								seqArray(x) = tokens(token) & "|"&lightParts(1)
							End If
							x=x+1
						End If
					End If
				End If
			End If
		Next
		glf_debugLog.WriteToLog "Convert Show", Join(seqArray)
		newShow(stepIdx) = seqArray
		stepIdx = stepIdx + 1
	Next
	Glf_ConvertShow = newShow
End Function

Private Function Glf_IsToken(mainString)
	' Check if the string contains an opening parenthesis and ends with a closing parenthesis
	If InStr(mainString, "(") > 0 And Right(mainString, 1) = ")" Then
		' Extract the substring within the parentheses
		Dim startPos, subString
		startPos = InStr(mainString, "(")
		subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
		Glf_IsToken = subString
	Else
		Glf_IsToken = Null
	End If
End Function

Private Function Glf_IsCondition(mainString)
	' Check if the string contains an opening { and ends with a closing }
	If InStr(mainString, "{") > 0 And Right(mainString, 1) = "}" Then
		Dim startPos, subString
		startPos = InStr(mainString, "{")
		subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
		Glf_IsCondition = subString
	Else
		Glf_IsCondition = Null
	End If
End Function

Function Glf_RotateArray(arr, direction)
    Dim n, rotatedArray, i
    ReDim rotatedArray(UBound(arr))
 
    If LCase(direction) = "l" Then
        For i = 0 To UBound(arr) - 1
            rotatedArray(i) = arr(i + 1)
        Next
        rotatedArray(UBound(arr)) = arr(0)
    ElseIf LCase(direction) = "r" Then
        For i = UBound(arr) To 1 Step -1
            rotatedArray(i) = arr(i - 1)
        Next
        rotatedArray(0) = arr(UBound(arr))
    Else
        ' Invalid direction
        Glf_RotateArray = arr
        Exit Function
    End If
    
    ' Return the rotated array
    Glf_RotateArray = rotatedArray
End Function

Function Glf_CopyArray(arr)
    Dim newArr, i
    ReDim newArr(UBound(arr))
    For i = 0 To UBound(arr)
        newArr(i) = arr(i)
    Next
    Glf_CopyArray = newArr
End Function

Function Glf_IsInArray(value, arr)
    Dim i
    Glf_IsInArray = False

    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            Glf_IsInArray = True
            Exit Function
        End If
    Next
End Function

'******************************************************
'*****   GLF Shows 		                           ****
'******************************************************

Dim glf_ShowOn : glf_ShowOn = Array(Array("(lights)|100"))
Dim glf_ShowOff : glf_ShowOff = Array(Array("(lights)|0"))
Dim glf_ShowFlash : glf_ShowFlash = Array(Array("(lights)|100"), Array("(lights)|0"))
Dim glf_ShowFlashColor : glf_ShowFlashColor = Array(Array("(lights)|100|(color)"), Array("(lights)|0|(color)"))

With GlfShotProfiles("default")
	With .States("on")
			.Show = glf_ShowFlash
	End With
	With .States("off")
			.Show = glf_ShowOff
	End With
End With

With GlfShotProfiles("flash_color")
	With .States("on")
			.Show = glf_ShowFlashColor
	End With
	With .States("off")
			.Show = glf_ShowOff
	End With
End With


'******************************************************
'*****   GLF Pin Events                            ****
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
        Glf_BCPUpdateTimer.Enabled = True
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
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & Getglf_currentPlayerNumber
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
            Glf_BCPUpdateTimer.Enabled = False
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

Sub Glf_BcpUpdate()
    If IsNull(bcpController) Then
        Exit Sub
    End If
    Dim messages : messages = bcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
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
    Public Property Let ActiveTime(value) : m_active_time = Glf_ParseInput(value, True) : End Property
    Public Property Let GracePeriod(value) : m_grace_period = Glf_ParseInput(value, True) : End Property
    Public Property Let HurryUpTime(value) : m_hurry_up_time = Glf_ParseInput(value, True) : End Property
    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let TimerStartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_timer_start_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let AutoLaunch(value) : m_auto_launch = value : End Property
    Public Property Let BallsToSave(value) : m_balls_to_save = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "ball_saves_" & name
        m_mode = mode.Name
        m_active_time = Null
	    m_grace_period = Null
        m_hurry_up_time = Null
        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_timer_start_events = CreateObject("Scripting.Dictionary")
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
        For Each evt in m_enable_events.Keys
            AddPinEventListener m_enable_events(evt).EventName, m_name & "_enable", "BallSaveEventHandler", 1000, Array("enable", Me, evt)
        Next
        For Each evt in m_timer_start_events.Keys
            AddPinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start", "BallSaveEventHandler", 1000, Array("timer_start", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events.Keys
            RemovePinEventListener m_enable_events(evt).EventName, m_name & "_enable"
        Next
        For Each evt in m_timer_start_events.Keys
            RemovePinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start"
        Next
    End Sub

    Public Sub Enable(evt)
        If m_enabled = True Then
            Exit Sub
        End If
        If Not IsNull(m_enable_events(evt).Condition) Then
            If GetRef(m_enable_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        m_enabled = True
        m_saving_balls = m_balls_to_save
        Log "Enabling. Auto launch: "&m_auto_launch&", Balls to save: "&m_balls_to_save
        AddPinEventListener "ball_drain", m_name & "_ball_drain", "BallSaveEventHandler", 1000, Array("drain", Me)
        DispatchPinEvent m_name&"_enabled", Null
        If UBound(m_timer_start_events.Keys) = -1 Then
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
        If Not IsNull(m_active_time) Then
            Dim active_time : active_time = GetRef(m_active_time(0))()
            Dim grace_period, hurry_up_time
            If Not IsNull(m_grace_period) Then
                grace_period = GetRef(m_grace_period(0))()
            Else
                grace_period = 0
            End If
            If Not IsNull(m_hurry_up_time) Then
                hurry_up_time = GetRef(m_hurry_up_time(0))()
            Else
                hurry_up_time = 0
            End If
            Log "Starting ball save timer: " & active_time
            Log "gametime: "& gametime & ". disabled at: " & gametime+active_time+grace_period
            SetDelay m_name&"_disable", "BallSaveEventHandler" , Array(Array("disable", Me),Null), active_time+grace_period
            SetDelay m_name&"_grace_period", "BallSaveEventHandler", Array(Array("grace_period", Me),Null), active_time
            SetDelay m_name&"_hurry_up_time", "BallSaveEventHandler", Array(Array("hurry_up_time", Me), Null), active_time-hurry_up_time
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
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = Replace(m_name, "ball_saves_", "") & ":" & vbCrLf
        yaml = yaml & "    active_time: " & m_active_time(1) & "s" & vbCrLf
        yaml = yaml & "    grace_period: " & m_grace_period(1) & "s" & vbCrLf
        yaml = yaml & "    hurry_up_time: " & m_hurry_up_time(1) & "s" & vbCrLf
        yaml = yaml & "    enable_events: "
        Dim evt,x : x = 0
        For Each evt in m_enable_events.Keys
            yaml = yaml & m_enable_events(evt).Raw
            If x <> UBound(m_enable_events.Keys) Then
                yaml = yaml & ", "
            End If
            x = x +1
        Next
        yaml = yaml & vbCrLf
        yaml = yaml & "    timer_start_events: "
        x=0
        For Each evt in m_timer_start_events.Keys
            yaml = yaml & m_timer_start_events(evt).Raw
            If x <> UBound(m_timer_start_events.Keys) Then
                yaml = yaml & ", "
            End If
            x = x +1
        Next
        yaml = yaml & vbCrLf
        yaml = yaml & "    auto_launch: " & m_auto_launch & vbCrLf
        yaml = yaml & "    balls_to_save: " & m_balls_to_save & vbCrLf
        yaml = yaml & vbCrLf
        ToYaml = yaml
    End Function
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
            ballSave.Enable ownProps(2)
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

Class GlfCounter

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

    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CountEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_count_events.Add newEvent.Name, newEvent
        Next
    End Property
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
        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_count_events = CreateObject("Scripting.Dictionary")

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
        For Each evt in m_enable_events.Keys
            AddPinEventListener m_enable_events(evt).EventName, m_name & "_enable", "CounterEventHandler", m_priority, Array("enable", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        If Not m_persist_state Then
            SetValue -1
        End If
        Dim evt
        For Each evt in m_enable_events.Keys
            RemovePinEventListener m_enable_events(evt).EventName, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_count_events.Keys
            AddPinEventListener m_count_events(evt).EventName, m_name & "_count", "CounterEventHandler", m_priority, Array("count", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_count_events.Keys
            RemovePinEventListener m_count_events(evt).EventName, m_name & "_count"
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
            glf_debugLog.WriteToLog m_name, message
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


Class GlfEventPlayer

    Private m_priority
    Private m_mode
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "event_player" : End Property
    Public Property Get Events() : Set Events = m_events : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Name, newEvent
        m_eventValues.Add newEvent.Name, value
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_event_player_play", "EventPlayerEventHandler", m_priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        If Not IsNull(m_events(evt).Condition) Then
            If GetRef(m_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            DispatchPinEvent evtValue, Null
        Next
    End Sub

    Public Function ToYaml
        Dim yaml
        Dim evt
        For Each evt In m_events.Keys()
            yaml = yaml & "  evt: " & Join(m_events(evt), ",") & vbCrLf
        Next
    End Function

End Class

Function EventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            eventPlayer.FireEvent ownProps(2)
    End Select
    EventPlayerEventHandler = kwargs
End Function

Class GlfLightPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug
    Private m_name
    Private m_value

    Public Property Let Events(value)
        Set m_events = value
        Dim evt
        For Each evt in m_events
            lightCtrl.CreateSeqRunner m_name & "_" & evt, m_priority
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_name = "light_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "light_player_activate", "LightPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "light_player_deactivate", "LightPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_light_player_play"
            PlayOff evt, m_events(evt)
        Next
    End Sub

    Public Sub Add(evt, lights)
        
        Dim light, lightsCount, x,tagLight, tagLights, lightParts
        lightsCount = 0
        For Each light in lights
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            Else
                If IsNull(lightCtrl.GetLightIdx(lightParts)) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts)
                    Log "Tag Lights: " & Join(tagLights)
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            End If
        Next
        Log "Adding " & lightsCount & " lights for event: " & evt 
        Dim seqArray
        ReDim seqArray(lightsCount-1)
        x=0
        For Each light in lights
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        If UBound(lightParts) >=1 Then
                            seqArray(x) = tagLight & "|100|"&lightParts(1)
                        Else
                            seqArray(x) = tagLight & "|100"
                        End If
                        x=x+1
                    Next
                Else
                    If UBound(lightParts) >= 1 Then
                        seqArray(x) = lightParts(0) & "|100|"&lightParts(1)
                    Else
                        seqArray(x) = lightParts(0) & "|100"
                    End If
                    x=x+1
                End If
            Else
                If IsNull(lightCtrl.GetLightIdx(lightParts)) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts)
                    For Each tagLight in tagLights
                        seqArray(x) = tagLight & "|100"
                        x=x+1
                    Next
                Else
                    seqArray(x) = lightParts & "|100"
                    x=x+1
                End If
            End If
        Next
        Log "Light List: " & Join(seqArray)
        m_events.Add evt, Array(lights, Array(seqArray))
    End Sub

    Public Sub Play(evt, lights)
        Dim light, lightParts
        lightCtrl.AddLightSeq m_name & "_" & evt, evt, lights(1), -1, 1, Null, m_priority, 0
        For Each light in lights(0)
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    Dim tagLight, tagLights
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        ProcessLight tagLight, lightParts, evt
                    Next
                Else
                    ProcessLight lightParts(0), lightParts, evt
                End If
            End If
        Next
    End Sub

    Public Sub PlayOff(evt, lights)
        Dim light
        lightCtrl.RemoveLightSeq m_name & "_" & evt, evt
    End Sub

    Private Sub ProcessLight(light_name, light_props, evt)
        If UBound(light_props) = 2 Then
            lightCtrl.FadeLightToColor light_name, light_props(1), light_props(2), m_name & "_" & evt & "_" & light_name, m_priority
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_light_player", message
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
            LightPlayer.Play ownProps(3), ownProps(2)
    End Select
    LightPlayerEventHandler = Null
End Function


Class GlfBaseModeDevice

    Private m_mode
    Private m_priority
    Private m_enable_events
    Private m_disable_events
    Private m_device
    Private m_parent
    Private m_debug

    Public Property Get EnableEvents(): Set EnableEvents = m_enable_events: End Property
    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Get DisableEvents(): Set DisableEvents = m_disable_events: End Property
    Public Property Let DisableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get Mode(): Set Mode = m_mode: End Property

    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode, device, parent)
        Set m_mode = mode
        m_priority = mode.Priority
        m_device = device
        Set m_parent = parent
        m_debug = mode.Debug

        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_disable_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode.Name & "_starting", m_device & "_" & m_parent.Name & "_activate", "BaseModeDeviceEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode.Name & "_stopping", m_device & "_" & m_parent.Name & "_deactivate", "BaseModeDeviceEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_enable_events.Keys()
            AddPinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable", "BaseModeDeviceEventHandler", m_priority, Array("enable", m_parent, evt)
        Next
        For Each evt In m_disable_events.Keys()
            AddPinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable", "BaseModeDeviceEventHandler", m_priority, Array("disable", m_parent, evt)
        Next
        m_parent.Activate
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        Dim evt
        For Each evt In m_enable_events.Keys()
            RemovePinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable"
        Next
        For Each evt In m_disable_events.Keys()
            RemovePinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable"
        Next
        m_parent.Deactivate
    End Sub

    Public Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode.Name & m_device & "_" & m_parent.Name & "_play", message
        End If
    End Sub
End Class


Function BaseModeDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Select Case evt
        Case "activate"
            device.Activate
        Case "deactivate"
            device.Deactivate
        Case "enable"
            device.Enable
        Case "disable"
            device.Disable
    End Select
    BaseModeDeviceEventHandler = kwargs
End Function

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
            AddPinEventListener evt, m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me)
        Next
        AddPinEventListener "ball_ended", m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me)
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_priority = priority
        Set m_ballsaves = CreateObject("Scripting.Dictionary")
        Set m_counters = CreateObject("Scripting.Dictionary")
        Set m_timers = CreateObject("Scripting.Dictionary")
        Set m_multiball_locks = CreateObject("Scripting.Dictionary")
        Set m_multiballs = CreateObject("Scripting.Dictionary")
        Set m_shots = CreateObject("Scripting.Dictionary")
        Set m_shot_groups = CreateObject("Scripting.Dictionary")
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
        yaml = "mode:" & vbCrLf
        yaml = yaml & "  start_events: " & Join(m_start_events, ",") & vbCrLf
        yaml = yaml & "  stop_events: " & Join(m_stop_events, ",") & vbCrLf
        yaml = yaml & "  priority: " & m_priority & vbCrLf
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
        ToYaml = yaml
        Log yaml
    End Function
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
        Case "started"
            DispatchPinEvent mode.Name & "_started", Null
        Case "stopped"
            DispatchPinEvent mode.Name & "_stopped", Null
    End Select
    ModeEventHandler = kwargs
End Function

Class GlfMultiballLocks

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
            glf_debugLog.WriteToLog m_name, message
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

Class GlfMultiball

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
            glf_debugLog.WriteToLog m_name, message
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
Class GlfShotGroup
    Private m_name
    Private m_mode
    Private m_priority
    private m_base_device
    private m_shots
    private m_common_state
    Private m_enable_rotation_events
    Private m_disable_rotation_events
    Private m_restart_events
    Private m_reset_events
    Private m_rotate_events
    Private m_rotate_left_events
    Private m_rotate_right_events
    Private rotation_enabled
    Private m_temp_shots
    Private m_rotation_pattern
    Private m_rotation_enabled
 
    Public Property Get Name(): Name = m_name: End Property
    Public Property Get CommonState()
        Dim state : state = m_base_device.Mode.Shots(m_shots(0)).State
        Dim shot
        For Each shot in m_shots
            If state <> m_base_device.Mode.Shots(shot).State Then
                CommonState = Empty
                Exit Property
            End If
        Next
        CommonState = state
    End Property
 
    Public Property Let Shots(value)
        m_shots = value
        Dim shot_name
        For Each shot_name in m_shots
            AddPinEventListener "shot_" & shot_name & "_hit", m_name & "_" & m_mode & "_hit", "ShotGroupEventHandler", m_priority, Array("hit", Me)
        Next
        m_rotation_pattern = Glf_CopyArray(Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).RotationPattern)
    End Property
 
    Public Property Let EnableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_rotation_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let DisableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_rotation_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RotateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RotateLeftEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_left_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RotateRightEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_right_events.Add newEvent.Name, newEvent
        Next
    End Property
 
	Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_common_state = Empty
        m_rotation_enabled = True
        m_rotation_pattern = Empty
 
        Set m_enable_rotation_events = CreateObject("Scripting.Dictionary")
        Set m_disable_rotation_events = CreateObject("Scripting.Dictionary")
        Set m_restart_events = CreateObject("Scripting.Dictionary")
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_left_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_right_events = CreateObject("Scripting.Dictionary")
        Set m_temp_shots = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot_group", Me)
 
        Set Init = Me
	End Function
 
    Public Sub Activate
        For Each evt in m_enable_rotation_events.Keys()
            m_rotation_enabled = False
            AddPinEventListener m_enable_rotation_events(evt).EventName, m_name & "_enable_rotation", "ShotGroupEventHandler", m_priority, Array("enable_rotation", Me, evt)
        Next
        For Each evt in m_disable_rotation_events.Keys()
            AddPinEventListener m_disable_rotation_events(evt).EventName, m_name & "_disable_rotation", "ShotGroupEventHandler", m_priority, Array("disable_rotation", Me, evt)
        Next
        For Each evt in m_restart_events.Keys()
            AddPinEventListener m_restart_events(evt).EventName, m_name & "_restart", "ShotGroupEventHandler", m_priority, Array("restart", Me, evt)
        Next
        For Each evt in m_reset_events.Keys()
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_reset", "ShotGroupEventHandler", m_priority, Array("reset", Me, evt)
        Next
        For Each evt in m_rotate_events.Keys()
            AddPinEventListener m_rotate_events(evt).EventName, m_name & "_rotate", "ShotGroupEventHandler", m_priority, Array("rotate", Me, evt)
        Next
        For Each evt in m_rotate_left_events.Keys()
            AddPinEventListener m_rotate_left_events(evt).EventName, m_name & "_rotate_left", "ShotGroupEventHandler", m_priority, Array("rotate_left", Me, evt)
        Next
        For Each evt in m_rotate_right_events.Keys()
            AddPinEventListener m_rotate_right_events(evt).EventName, m_name & "_rotate_right", "ShotGroupEventHandler", m_priority, Array("rotate_right", Me, evt)
        Next
    End Sub
 
    Public Sub Deactivate
        m_rotation_enabled = True
        For Each evt in m_enable_rotation_events.Keys()
            RemovePinEventListener m_enable_rotation_events(evt).EventName, m_name & "_enable_rotation"
        Next
        For Each evt in m_disable_rotation_events.Keys()
            RemovePinEventListener m_disable_rotation_events(evt).EventName, m_name & "_disable_rotation"
        Next
        For Each evt in m_restart_events.Keys()
            RemovePinEventListener m_restart_events(evt).EventName, m_name & "_restart"
        Next
        For Each evt in m_reset_events.Keys()
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_reset"
        Next
        For Each evt in m_rotate_events.Keys()
            RemovePinEventListener m_rotate_events(evt).EventName, m_name & "_rotate"
        Next
        For Each evt in m_rotate_left_events.Keys()
            RemovePinEventListener m_rotate_left_events(evt).EventName, m_name & "_rotate_left"
        Next
        For Each evt in m_rotate_right_events.Keys()
            RemovePinEventListener m_rotate_right_events(evt).EventName, m_name & "_rotate_right"
        Next
    End Sub
 
    Private Function CheckForComplete()
        Dim state : state = CommonState()
        If state = m_common_state Then
            Exit Function
        End If
 
        m_common_state = state
 
        If state = Empty Then
            Exit Function
        End If
 
        Dim state_name : state_name = Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).StateName(m_state)
 
        Log "Shot group is complete with state: " & state_name
        Dim kwargs : Set kwargs = GlfKwargs()
		With kwargs
            .Add "state", state_name
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        DispatchPinEvent m_name & "_" & state_name & "_complete", Null
 
    End Function
 
    Public Sub Enable()
        Dim shot
        m_base_device.Log "Enabling"
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Enable()
        Next
        Dim evt
    End Sub
 
    Public Sub Disable()
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Disable()
        Next
    End Sub
 
    Public Sub EnableRotation
        m_base_device.Log "Enabling Rotation"
        m_rotation_enabled = True
    End Sub
 
    Public Sub DisableRotation
        m_base_device.Log "Disabling Rotation"
        m_rotation_enabled = False
    End Sub
 
    Public Sub Restart
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Restart()
        Next
    End Sub
 
    Public Sub Reset
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Reset()
        Next
    End Sub
 
    Public Sub Rotate(direction)
 
        If m_rotation_enabled = False Then
            Exit Sub
        End If
        Dim shots_to_rotate : shots_to_rotate = Array()
 
        m_temp_shots.RemoveAll
        Dim shot
        For Each shot in m_shots
            If m_base_device.Mode.Shots(shot).CanRotate Then
                m_temp_shots.Add shot, m_base_device.Mode.Shots(shot)
            End If
        Next
 
        Dim shot_states, x
        x=0
        ReDim shot_states(UBound(m_temp_shots.Keys))
        For Each shot in m_temp_shots.Keys
            shot_states(x) = m_temp_shots(shot).State
            x=x+1
        Next 
 
        If direction = Empty Then
            direction = m_rotation_pattern(0)
            Glf_RotateArray m_rotation_pattern, "l"
        End If
 
        shot_states = Glf_RotateArray(shot_states, direction)
        x=0
        For Each shot in m_temp_shots.Keys
            m_base_device.Log "Rotating Shot:" & shot
            m_temp_shots(shot).Jump shot_states(x), True, False
            x=x+1
        Next 
 
    End Sub
 
    Public Function ToYaml
        Dim yaml
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml & "    shots: " & Join(m_shots, ",") & vbCrLf
 
        If UBound(m_enable_rotation_events.Keys) > -1 Then
            yaml = yaml & "    enable_rotation_events: "
            x=0
            For Each key in m_enable_rotation_events.keys
                yaml = yaml & m_enable_rotation_events(key).Raw
                If x <> UBound(m_enable_rotation_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_disable_rotation_events.Keys) > -1 Then
            yaml = yaml & "    disable_rotation_events: "
            x=0
            For Each key in m_disable_rotation_events.keys
                yaml = yaml & m_disable_rotation_events(key).Raw
                If x <> UBound(m_disable_rotation_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_restart_events.Keys) > -1 Then
            yaml = yaml & "    restart_events: "
            x=0
            For Each key in m_restart_events.keys
                yaml = yaml & m_restart_events(key).Raw
                If x <> UBound(m_restart_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_reset_events.Keys) > -1 Then
            yaml = yaml & "    reset_events: "
            x=0
            For Each key in m_reset_events.keys
                yaml = yaml & m_reset_events(key).Raw
                If x <> UBound(m_reset_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_events.Keys) > -1 Then
            yaml = yaml & "    rotate_events: "
            x=0
            For Each key in m_rotate_events.keys
                yaml = yaml & m_rotate_events(key).Raw
                If x <> UBound(m_rotate_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_left_events.Keys) > -1 Then
            yaml = yaml & "    rotate_left_events: "
            x=0
            For Each key in m_rotate_left_events.keys
                yaml = yaml & m_rotate_left_events(key).Raw
                If x <> UBound(m_rotate_left_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_right_events.Keys) > -1 Then
            yaml = yaml & "    rotate_right_events: "
            x=0
            For Each key in m_rotate_right_events.keys
                yaml = yaml & m_rotate_right_events(key).Raw
                If x <> UBound(m_rotate_right_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_base_device.EnableEvents.Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents.keys
                yaml = yaml & m_base_device.EnableEvents(key).Raw
                If x <> UBound(m_base_device.EnableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_base_device.DisableEvents.Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents.keys
                yaml = yaml & m_base_device.DisableEvents(key).Raw
                If x <> UBound(m_base_device.DisableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        yaml = yaml & "    priority: " & m_priority & vbCrLf
        yaml = yaml & "    rotation_enabled: " & m_rotation_enabled & vbCrLf
 
        ToYaml = yaml
        End Function
End Class
 
Function ShotGroupEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Select Case evt
        Case "hit"
            DispatchPinEvent device.Name & "_hit", Null
            DispatchPinEvent device.Name & "_" & kwargs("state") & "_hit", Null
        Case "enable_rotation"
            device.EnableRotation
        Case "disable_rotation"
            device.DisableRotation
        Case "restart"
            device.Restart
        Case "reset"
            device.Reset
        Case "rotate"
            device.Rotate Empty
        Case "rotate_left"
            device.Rotate "l"
        Case "rotate_right"
            device.Rotate "r"
    End Select
    If IsObject(args(1)) Then
        Set ShotGroupEventHandler = kwargs
    Else
        ShotGroupEventHandler = kwargs
    End If
 
End Function

Class GlfShotProfile

    Private m_name
    Private m_advance_on_hit
    Private m_block
    Private m_loop
    Private m_rotation_pattern
    Private m_states
    Private m_states_not_to_rotate
    Private m_states_to_rotate

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AdvanceOnHit(): AdvanceOnHit = m_advance_on_hit: End Property
    Public Property Get Block(): Block = m_block: End Property
    Public Property Let Block(input): m_block = input: End Property
    Public Property Get ProfileLoop(): ProfileLoop = m_loop: End Property
    Public Property Get RotationPattern(): RotationPattern = m_rotation_pattern: End Property
    Public Property Get States(name)
        If m_states.Exists(name) Then
            Set States = m_states(name)
        Else
            Dim new_state : Set new_state = (new GlfShowPlayerItem)()
            m_states.Add name, new_state
            Set States = new_state
        End If
    End Property
    Public Property Get StateForIndex(index)
        Dim stateItems : stateItems = m_states.Items()
        If UBound(stateItems) >= index Then
            Set StateForIndex = stateItems(index)
        Else
            StateForIndex = Null
        End If
    End Property
    Public Property Get StateName(index)
        Dim stateKeys : stateKeys = m_states.Keys()
        If UBound(stateKeys) >= index Then
            StateName = stateKeys(index)
        Else
            StateName = Empty
        End If
    End Property
    Public Property Get StatesCount()
        StatesCount = UBound(m_states.Keys())
    End Property

    Public Property Get StateNamesToRotate(): StateNamesToRotate = m_states_to_rotate: End Property
    Public Property Let StateNamesToRotate(input): m_states_to_rotate = input: End Property
    Public Property Get StateNamesNotToRotate(): StateNamesNotToRotate = m_states_not_to_rotate: End Property
    Public Property Let StateNamesNotToRotate(input): m_states_not_to_rotate = input: End Property
    
	Public default Function init(name)
        m_name = "shotprofile_" & name
        m_advance_on_hit = True
        m_block = False
        m_loop = False
        m_rotation_pattern = Array("r")
        m_states_to_rotate = Array()
        m_states_not_to_rotate = Array()
        Set m_states = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

End Class

Class GlfShot

    Private m_name
    Private m_mode
    Private m_priority
    Private m_base_device
    Private m_profile
    Private m_control_events
    Private m_advance_events
    Private m_reset_events
    Private m_restart_events
    Private m_switches
    Private m_tokens
    Private m_hit_events
    Private m_start_enabled
    Private m_show_cache
    Private m_state
    Private m_enabled

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Profile(): Profile = m_profile: End Property
    Public Property Get ShotKey(): ShotKey = m_name & "_" & m_profile: End Property
    Public Property Get State(): State = m_state: End Property
    Public Property Get Tokens() : Set Tokens = m_tokens : End Property
    Public Property Get CanRotate()
        If Glf_IsInArray(Glf_ShotProfiles(m_profile).StateName(m_state), Glf_ShotProfiles(m_profile).StateNamesNotToRotate) Then
            CanRotate = False
        Else
            CanRotate = True
        End If
    End Property
    
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Get ControlEvents(name)
        If m_control_events.Exists(name) Then
            Set ControlEvents = m_control_events(name)
        Else
            Dim newEvent : Set newEvent = (new GlfShotControlEvent)()
            m_control_events.Add name, newEvent
            Set ControlEvents = newEvent
        End If
    End Property
    Public Property Let AdvanceEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_advance_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Name, newEvent
        Next
    End Property   
    Public Property Let Profile(value) : m_profile = value : End Property
    Public Property Let Switch(value) : m_switches = Array(value) : End Property
    Public Property Let Switches(value) : m_switches = value : End Property
    Public Property Let StartEnabled(value) : m_start_enabled = value : End Property
    Public Property Let HitEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_hit_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "shot_" & name
        m_mode = mode.Name
        m_priority = mode.Priority

        m_enabled = False

        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot", Me)

        m_profile = "default"
        m_state = -1
        m_switches = Array()
        m_start_enabled = Empty
        Set m_hit_events = CreateObject("Scripting.Dictionary")
        Set m_tokens = CreateObject("Scripting.Dictionary")
        Set m_show_cache = CreateObject("Scripting.Dictionary")
        Set m_advance_events = CreateObject("Scripting.Dictionary")
        Set m_control_events = CreateObject("Scripting.Dictionary")
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_restart_events = CreateObject("Scripting.Dictionary")

        Set Init = Me
	End Function

    Public Sub Activate()
        m_state = 0
        If m_start_enabled = True Then
            Enable()
        Else
            If IsEmpty(m_start_enabled) And UBound(m_base_device.EnableEvents.Keys) = -1 Then
                Enable()
            End If
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
        m_state = -1
    End Sub

    Public Sub Enable()
        m_base_device.Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_switches
            AddPinEventListener evt & "_active", m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me)
        Next
        For Each evt in m_hit_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me)
        Next
        For Each evt in m_advance_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_advance", "ShotEventHandler", m_priority, Array("advance", Me)
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events
                AddPinEventListener cEvt, m_mode & "_" & m_name & "_control", "ShotEventHandler", m_priority, Array("control", Me, m_control_events(evt))
            Next
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_reset", "ShotEventHandler", m_priority, Array("reset", Me)
        Next
        For Each evt in m_restart_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_restart", "ShotEventHandler", m_priority, Array("restart", Me)
        Next
        'Play the show for the active state
        PlayShowForState(m_state)
    End Sub

    Public Sub Disable()
        m_base_device.Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_switches
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        For Each evt in m_hit_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        For Each evt in m_advance_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_advance"
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events
                RemovePinEventListener cEvt, m_mode & "_" & m_name & "_control"
            Next
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_reset"
        Next
        For Each evt in m_restart_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_restart"
        Next
        Dim x
        For x=0 to Glf_ShotProfiles(m_profile).StatesCount()
            StopShowForState(x)
        Next
    End Sub

    Private Sub StopShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        m_base_device.Log "Removing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        lightCtrl.RemoveLightSeq m_mode & "_" & m_name, profileState.Key
    End Sub

    Private Sub PlayShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        m_base_device.Log "Playing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        If IsObject(profileState) Then
            If IsArray(profileState.Show) Then
                If m_show_cache.Exists(CStr(m_state) & "_" & profileState.Key) Then
                    lightCtrl.AddLightSeq m_mode & "_" & m_name, profileState.Key, m_show_cache(CStr(m_state) & "_" & profileState.Key), profileState.Loops, profileState.Speed, Null, m_priority, profileState.SyncMS
                Else
                    Dim key
                    Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
                    Dim profileStateTokens : Set profileStateTokens = profileState.Tokens
                    For Each key In profileStateTokens.Keys
                        mergedTokens.Add key, profileStateTokens(key)
                    Next
                    Dim shotTokens : Set shotTokens = m_tokens
                    For Each key In shotTokens.Keys
                        If mergedTokens.Exists(key) Then
                            mergedTokens(key) = shotTokens(key)
                        Else
                            mergedTokens.Add key, shotTokens(key)
                        End If
                    Next
                    Dim show : show = Glf_ConvertShow(profileState.Show, mergedTokens)
                    m_show_cache.Add CStr(m_state) & "_" & profileState.Key, show
                    lightCtrl.AddLightSeq m_mode & "_" & m_name, profileState.Key, show, profileState.Loops, profileState.Speed, Null, m_priority, profileState.SyncMS
                End If
            End If
        End If
    End Sub

    Public Sub Hit(evt)
        If m_enabled = False Then
            Exit Sub
        End If

        Dim profile : Set profile = Glf_ShotProfiles(m_profile)
        Dim old_state : old_state = m_state
        m_base_device.Log "Hit! Profile: "&m_profile&", State: " & profile.StateName(m_state)

        Dim advancing
        If profile.AdvanceOnHit Then
            m_base_device.Log "Advancing shot because advance_on_hit is True."
            advancing = Advance(False)
        Else
            m_base_device.Log "Not advancing shot"
            advancing = False
        End If

    
        If profile.Block Then
            Glf_EventBlocks(evt).Add Name, True
        Else
            Glf_EventBlocks(evt).Add ShotKey, True
        End If
        Dim kwargs : Set kwargs = GlfKwargs()
		With kwargs
            .Add "profile", m_profile
            .Add "state", profile.StateName(old_state)
            .Add "advancing", advancing
        End With

        DispatchPinEvent m_name & "_hit", kwargs
        DispatchPinEvent m_name & "_" & m_profile & "_hit", kwargs
        DispatchPinEvent m_name & "_" & m_profile & "_" & profile.StateName(old_state) & "_hit", kwargs
        DispatchPinEvent m_name & "_" & profile.StateName(old_state) & "_hit", kwargs
        
    End Sub

    Public Function Advance(force)

        If m_enabled = False And force = False Then
            Advance = False
            Exit Function
        End If
        Dim profile : Set profile = Glf_ShotProfiles(m_profile)

        m_base_device.Log "Advancing 1 step. Profile: "&m_profile&", Current State: " & profile.StateName(m_state)

        If profile.StatesCount() = m_state Then
            If profile.ProfileLoop Then
                StopShowForState(m_state)
                m_state = 0
                PlayShowForState(m_state)
            Else
                Advance = False
                Exit Function
            End If
        Else
            StopShowForState(m_state)
            m_state = m_state + 1
            PlayShowForState(m_state)
        End If

        Advance = True
        
    End Function

    Public Sub Reset()
        Jump 0, True, False
    End Sub

    Public Sub Jump(state, force, force_show)
        m_base_device.Log "Received jump request. State: " & state & ", Force: "& force

        If Not m_enabled And Not force Then
            m_base_device.Log "Profile is disabled and force is False. Not jumping"
            Exit Sub
        End If
        If state = m_state And Not force_show Then
            m_base_device.Log "Shot is already in the jump destination state"
            Exit Sub
        End If
        m_base_device.Log "Jumping to profile state " & state

        StopShowForState(m_state)
        m_state = state
        PlayShowForState(m_state)
    End Sub

    Public Sub Restart()
        Reset()
        Enable()
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & Replace(m_name, "shot_", "") & ":" & vbCrLf
        If UBound(m_switches) = 0 Then
            yaml = yaml & "    switch: " & m_switches(0) & vbCrLf
        Else
            yaml = yaml & "    switches: " & Join(m_switches, ",") & vbCrLf
        End If
        yaml = yaml & "    show_tokens: " & vbCrLf
        dim key
        For Each key in m_tokens.keys
            If IsArray(m_tokens(key)) Then
                yaml = yaml & "      " & key & ": " & Join(m_tokens(key), ",") & vbCrLf
            Else  
                yaml = yaml & "      " & key & ": " & m_tokens(key) & vbCrLf
            End If
        Next

        If UBound(m_base_device.EnableEvents.Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents.keys
                yaml = yaml & m_base_device.EnableEvents(key).Raw
                If x <> UBound(m_base_device.EnableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_base_device.DisableEvents.Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents.keys
                yaml = yaml & m_base_device.DisableEvents(key).Raw
                If x <> UBound(m_base_device.DisableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_advance_events.Keys) > -1 Then
            yaml = yaml & "    advance_events: "
            x=0
            For Each key in m_advance_events.keys
                yaml = yaml & m_advance_events(key).Raw
                If x <> UBound(m_advance_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_hit_events.Keys) > -1 Then
            yaml = yaml & "    hit_events: "
            x=0
            For Each key in m_hit_events.keys
                yaml = yaml & m_hit_events(key).Raw
                If x <> UBound(m_hit_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        yaml = yaml & "    profile: " & m_profile & vbCrLf
        If Not IsEmpty(m_start_enabled) Then
            yaml = yaml & "    start_enabled: " & m_start_enabled & vbCrLf
        End If
        If UBound(m_restart_events.Keys) > -1 Then
            yaml = yaml & "    restart_events: "
            x=0
            For Each key in m_restart_events.keys
                yaml = yaml & m_restart_events(key).Raw
                If x <> UBound(m_restart_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_reset_events.Keys) > -1 Then
            yaml = yaml & "    reset_events: "
            x=0
            For Each key in m_reset_events.keys
                yaml = yaml & m_reset_events(key).Raw
                If x <> UBound(m_reset_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_control_events.Keys) > -1 Then
            yaml = yaml & "    control_events: " & vbCrLf
            For Each key in m_control_events.keys
                yaml = yaml & "      - events: "
                Dim cEvt
                x=0
                For Each cEvt in m_control_events(key).Events
                    yaml = yaml & cEvt
                    If x <> UBound(m_control_events(key).Events) Then
                        yaml = yaml & ", "
                    End If
                    x = x + 1
                Next
                yaml = yaml & vbCrLf
                yaml = yaml & "      state: " & m_control_events(key).State & vbCrLf
            Next
        End If
        
        ToYaml = yaml
    End Function
End Class

Function ShotEventHandler(args)
    Dim ownProps, kwargs, e
    ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    e = args(2)
    Dim evt : evt = ownProps(0)
    Dim shot : Set shot = ownProps(1)
    Select Case evt
        Case "activate"
            shot.Activate
        Case "deactivate"
            shot.Deactivate
        Case "enable"
            shot.Enable
        Case "hit"
            If Not Glf_EventBlocks(e).Exists(shot.Name) And Not Glf_EventBlocks(e).Exists(shot.ShotKey) Then
                shot.Hit e
            End If
        Case "advance"
            shot.Advance False
        Case "control"
            shot.Jump ownProps(2).State, ownProps(2).Force, ownProps(2).ForceShow
        Case "reset"
            shot.Reset
        Case "restart"
            shot.Restart
            
    End Select
    If IsObject(args(1)) Then
        Set ShotEventHandler = kwargs
    Else
        ShotEventHandler = kwargs
    End If
End Function

Class GlfShotControlEvent
	Private m_events, m_state, m_force, m_force_show
  
	Public Property Get Events(): Events = m_events: End Property
    Public Property Let Events(input): m_events = input: End Property

    Public Property Get State(): State = m_state End Property
    Public Property Let State(input): m_state = input End Property

    Public Property Get Force(): Force = m_force: End Property
	Public Property Let Force(input): m_force = input: End Property
  
	Public Property Get ForceShow(): ForceShow = m_force_show: End Property
	Public Property Let ForceShow(input): m_force_show = input: End Property   

	Public default Function init()
        m_events = Array()
        m_state = 0
        m_force = True
        m_force_show = False
	    Set Init = Me
	End Function

End Class

Class GlfShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_show_cache
    Private m_debug
    Private m_name
    Private m_value

    Public Property Get Events(name)
        If m_events.Exists(name) Then
            Set Events = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfShowPlayerItem)()
            m_events.Add name, new_event
            Set Events = new_event
        End If
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = True
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_show_cache = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "show_player_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "show_player_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            If IsObject(m_events(evt)) Then
                AddPinEventListener Replace(evt, "_" & m_events(evt).Key, "") , m_mode & "_" & m_events(evt).Key & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
            Else
                AddPinEventListener Replace(evt, "_" & m_events(evt), "") , m_mode & "_" & m_events(evt) & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
            End If
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            If IsObject(m_events(evt)) Then
                RemovePinEventListener Replace(evt, "_" & m_events(evt).Key, ""), m_mode & "_" & m_events(evt).Key & "_show_player_play"
            Else
                RemovePinEventListener Replace(evt, "_" & m_events(evt), ""), m_mode & "_" & m_events(evt) & "_show_player_play"
            End If
            PlayOff m_events(evt).Key
        Next
    End Sub

    Public Sub Play(evt, show)
        
        If show.Action = "stop" Then
            PlayOff show.Key
        Else
            If m_show_cache.Exists(show.Key) Then
                lightCtrl.AddLightSeq m_name & "_" & show.Key, show.Key, m_show_cache(show.Key), show.Loops, show.Speed, Null, m_priority, 0
            Else
                Dim cachedShow : cachedShow = Glf_ConvertShow(show.Show, show.Tokens)
                m_show_cache.Add show.Key, cachedShow
                lightCtrl.AddLightSeq m_name & "_" & show.Key, show.Key, cachedShow, show.Loops, show.Speed, Null, m_priority, 0
            End If
        End If
    End Sub

    'Public Sub StopShow(evt, key)
    '    m_events.Add evt & "_" & key & ".stop", key & ".stop"
    'End Sub

    Public Sub PlayOff(key)
        lightCtrl.RemoveLightSeq m_name & "_" & key, key
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_show_player", message
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
            ShowPlayer.Play ownProps(3), ownProps(2)
    End Select
    ShowPlayerEventHandler = Null
End Function

Class GlfShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Show(): Show = m_show: End Property
	Public Property Let Show(input): m_show = input: End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property        

    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property        
  
	Public default Function init()
        m_action = "play"
        m_key = ""
        m_loops = -1
        m_speed = 1
        m_syncms = 0
        Set m_tokens = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class

Class GlfStateMachine
    Private m_name
    Private m_priority
    private m_base_device
    Private m_states
    Private m_transistions
 
    Private m_current_state
 
    Public Property Get Name(): Name = m_name: End Property
 
    Public Property Get States(): States = m_states: End Property
    Public Property Let States(value)
        m_states = value
    End Property
 
    Public Property Get Transistions(): Transistions = m_transistions: End Property
    Public Property Let Transistions(value)
        m_transistions = value
    End Property
 
    Public Property Get CurrentState(): CurrentState = m_current_state: End Property
    Public Property Let CurrentState(value)
        m_current_state = value
    End Property
 
    Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
 
        Set m_states = CreateObject("Scripting.Dictionary")
        Set m_transistions = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "state_machine", Me)
 
        Set Init = Me
    End Function
 
End Class
 
Class GlfStateMachineState
	Private m_name, m_label, m_show_when_active, m_show_tokens, m_events_when_started, m_events_when_stopped
 
	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property
 
    Public Property Get Label(): Label = m_label: End Property
    Public Property Let Label(input): m_label = input: End Property
 
    Public Property Get ShowWhenActive(): ShowWhenActive = m_show_when_active: End Property
    Public Property Let ShowWhenActive(input): m_show_when_active = input: End Property
 
    Public Property Get ShowTokens(): Set ShowTokens = m_show_tokens: End Property
 
    Public Property Get EventsWhenStarted(): EventsWhenStarted = m_events_when_started: End Property
    Public Property Let EventsWhenStarted(input): m_events_when_started = input: End Property
 
    Public Property Get EventsWhenStopped(): EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(input): m_events_when_stopped = input: End Property
 
	Public default Function init(name)
        m_name = name
        m_label = Empty
        m_show_when_active = Empty
        Set m_show_tokens = CreateObject("Scripting.Dictionary")
        Set m_events_when_started = Array()
        Set m_events_when_stopped = Array()
	    Set Init = Me
	End Function
 
End Class
 
Class GlfStateMachineTranistion
	Private m_source, m_target, m_events, m_events_when_transistioning
 
    Public Property Get Source(): Source = m_source: End Property
    Public Property Let Source(input): m_source = input: End Property
 
    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input): m_target = input: End Property
 
    Public Property Get Events(): Events = m_events: End Property
    Public Property Let Events(input): m_events = input: End Property
 
    Public Property Get EventsWhenTransitioning(): EventsWhenTransitioning = m_events_when_transitioning: End Property
    Public Property Let EventsWhenTransitioning(input): m_events_when_transitioning = input: End Property
 
	Public default Function init(name)
        m_source = name
        m_target = Empty
        m_events = Array()
        m_events_when_transistioning = Array()
	    Set Init = Me
	End Function
 
End Class


Class GlfTimer

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

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get StartValue() : StartValue = m_start_value : End Property
    Public Property Get EndValue() : EndValue = m_end_value : End Property
    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Get StartEvents() : StartEvents = m_start_events : End Property
    Public Property Get StopEvents() : StopEvents = m_stop_events : End Property
    
    Public Property Let StartValue(value) : m_start_value = value : End Property
    Public Property Let EndValue(value) : m_end_value = value : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "TimerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "TimerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_start_events.Keys
            AddPinEventListener m_start_events(evt).EventName, m_name & "_start", "TimerEventHandler", m_priority, Array("start", Me, evt)
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events.Keys
                AddPinEventListener m_stop_events(evt).EventName, m_name & "_stop", "TimerEventHandler", m_priority, Array("stop", Me, evt)
            Next
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_start_events.Keys
            RemovePinEventListener m_start_events(evt).EventName, m_name & "_start"
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events.Keys
                RemovePinEventListener m_stop_events(evt).EventName, m_name & "_stop"
            Next
        End If
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub StartTimer(evt)
        If Not IsNull(m_start_events(evt).Condition) Then
            If GetRef(m_start_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        Log "Started"
        DispatchPinEvent m_name & "_started", Null
        m_value = m_start_value
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
    End Sub

    Public Sub StopTimer(evt)
        If Not IsNull(m_stop_events(evt).Condition) Then
            If GetRef(m_stop_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
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
            glf_debugLog.WriteToLog m_name, message
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
            timer.StartTimer ownProps(2)
        Case "stop"
            timer.StopTimer ownProps(2)
        Case "tick"
            timer.Tick 
    End Select
    TimerEventHandler = kwargs
End Function
Class GlfVariablePlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Get Events(name)
        Dim newEvent : Set newEvent = (new GlfVariablePlayerEvent)(name)
        m_events.Add newEvent.BaseEvent.Name, newEvent
        Set Events = newEvent
    End Property
   
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "variable_player_activate", "VariablePlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "variable_player_deactivate", "VariablePlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_variable_player_play", "VariablePlayerEventHandler", m_priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_variable_player_play"
        Next
    End Sub

    Public Sub Play(evt)
        Log "Playing: " & evt
        If Not IsNull(m_events(evt).BaseEvent.Condition) Then
            If GetRef(m_events(evt).BaseEvent.Condition)() = False Then
                Exit Sub
            End If
        End If
        Dim vKey, v
        For Each vKey in m_events(evt).Variables.Keys
            Log "Setting Variable " & vKey
            Set v = m_events(evt).Variable(vKey)
            Dim varValue : varValue = v.VariableValue
            Select Case v.Action
                Case "add"
                    SetPlayerState vKey, GetPlayerState(vKey) + varValue
                Case "set"
                    SetPlayerState vKey, varValue
        End Select
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_variable_player_play", message
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

End Class

Class GlfVariablePlayerItem
	Private m_block, m_show, m_flaot, m_int, m_string, m_player, m_action, m_type
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Block(): Block = m_block End Property
    Public Property Let Block(input): m_block = input End Property

	Public Property Let Float(input): m_float = Glf_ParseInput(input, False): m_type = "float" : End Property
  
	Public Property Let Int(input): m_int = Glf_ParseInput(input, False): m_type = "int" : End Property
  
	Public Property Let String(input): m_string = input: m_type = "string" : End Property

    Public Property Get VariableType(): VariableType = m_type: End Property
    Public Property Get VariableValue()
        Select Case m_type
            Case "float"
                VariableValue = GetRef(m_float(0))()
            Case "int"
                VariableValue = GetRef(m_int(0))()
            Case "string"
                VariableValue = m_string
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
        Case "activate"
            variablePlayer.Activate
        Case "deactivate"
            variablePlayer.Deactivate
        Case "play"
            variablePlayer.Play ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set VariablePlayerEventHandler = kwargs
    Else
        VariablePlayerEventHandler = kwargs
    End If
    
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
    executionTime = AlignToNearest5th(gametime + delayInMs)
    
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

    glf_debugLog.WriteToLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc & ", ExecutionTime: " & executionTime
    delayQueue(executionTime).Add name, (new DelayObject)(name, callbackFunc, executionTime, args)
    delayQueueMap.Add name, executionTime
    
End Sub

Function AlignToNearest5th(timeMs)
    AlignToNearest5th = Int(timeMs / 200) * 200
End Function

Function RemoveDelay(name)
    If delayQueueMap.Exists(name) Then
        If delayQueue.Exists(delayQueueMap(name)) Then
            If delayQueue(delayQueueMap(name)).Exists(name) Then
                delayQueue(delayQueueMap(name)).Remove name
            End If
            delayQueueMap.Remove name
            RemoveDelay = True
            glf_debugLog.WriteToLog "Delay", "Removing delay for " & name
            Exit Function
        End If
    End If
    RemoveDelay = False
End Function

Sub DelayTick()
    Dim key, delayObject

    Dim executionTime
    executionTime = AlignToNearest5th(gametime)
    If delayQueue.Exists(executionTime) Then
        For Each key In delayQueue(executionTime).Keys()
            Set delayObject = delayQueue(executionTime)(key)
            glf_debugLog.WriteToLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
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
	Public Property Get HasBall()
        HasBall = (Not IsNull(m_balls(0)) And m_ejecting = False)
    End Property
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
            m_balls(x) = Null
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
        m_mechcanical_eject = False
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
        If m_mechcanical_eject = True And m_eject_timeout > 0 Then
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
            If Not IsNull(m_balls(0)) Then
                m_balls(0).VelX = m_eject_strength * Cos(m_eject_pitch) * Sin(m_eject_angle)
                m_balls(0).VelY = m_eject_strength * Cos(m_eject_pitch) * Cos(m_eject_angle)
                m_balls(0).VelZ = m_eject_strength * Sin(m_eject_pitch)
                Log "VelX: " &  m_balls(0).VelX & ", VelY: " &  m_balls(0).VelY & ", VelZ: " &  m_balls(0).VelZ
            End If
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
            glf_debugLog.WriteToLog m_name, message
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
            glf_debugLog.WriteToLog m_name, message
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

Class GlfEvent
	Private m_raw, m_name, m_event, m_condition
  
    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventName() : EventName = m_event : End Property
    Public Property Get Condition() : Condition = m_condition : End Property
    Public Property Get Raw() : Raw = m_raw : End Property

	Public default Function init(evt)
        m_raw = evt
        Dim parsedEvent : parsedEvent = Glf_ParseEventInput(evt)
        m_name = parsedEvent(0)
        m_event = parsedEvent(1)
        m_condition = parsedEvent(2)
	    Set Init = Me
	End Function

End Class


'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub Glf_AddPlayer()
    Select Case UBound(playerState.Keys())
        Case -1:
            playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            Glf_BcpAddPlayer 1
            glf_currentPlayer = "PLAYER 1"
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
            glf_canAddPlayers = False
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
    EmitAllglf_playerEvents()
End Function

'****************************
' StartGame
'
'*****************************
Sub Glf_StartGame()
    glf_gameStarted = True
    If useBcp Then
        bcpController.Send "player_turn_start?player_num=int:1"
        bcpController.Send "ball_start?player_num=int:1&ball=int:1"
        bcpController.PlaySlide "base", "base", 1000
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
    glf_debugLog.WriteToLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    swTrough1.kick 90, 10
    glf_debugLog.WriteToLog "Release Ball", "Just Kicked"
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
    glf_debugLog.WriteToLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
    glf_debugLog.WriteToLog "end_of_ball, balls in play", CStr(glf_BIP)
    If ballsToSave <= 0 Then
        Exit Function
    End If

    If glf_BIP > 0 Then
        Exit Function
    End If
        
    DispatchPinEvent GLF_BALL_ENDED, Null
    SetPlayerState GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            If UBound(playerState.Keys()) > 0 Then
                glf_currentPlayer = "PLAYER 2"
            End If
        Case "PLAYER 2":
            If UBound(playerState.Keys()) > 1 Then
                glf_currentPlayer = "PLAYER 3"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 3":
            If UBound(playerState.Keys()) > 2 Then
                glf_currentPlayer = "PLAYER 4"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 4":
            glf_currentPlayer = "PLAYER 1"
    End Select
    
    If useBcp Then
        bcpController.SendPlayerVariable "number", Getglf_currentPlayerNumber(), previousPlayerNumber
    End If
    If GetPlayerState(GLF_CURRENT_BALL) > glf_ballsPerGame Then
        DispatchPinEvent GLF_GAME_OVER, Null
        glf_gameStarted = False
        glf_currentPlayer = Null
        playerState.RemoveAll()
    Else
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000 
    End If
    
End Function

Public Function EndOfBallNextPlayer(args)
    DispatchPinEvent GLF_NEXT_PLAYER, Null
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

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqOffset, m_tableSeqSpeed, m_tableSeqDirection, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, m_lightmaps, m_seqOverrideRunners, m_pauseMainLights, m_pausedLights, m_minX, m_minY, m_maxX, m_maxY, m_width, m_height, m_centerX, m_centerY, m_coordsX, m_coordsY, m_angles, m_radii, m_tags, m_debug, m_syncMs, m_syncMsCounter

    Private Lights(260)

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
        Set m_syncMs = CreateObject("Scripting.Dictionary")
        Set m_syncMsCounter = CreateObject("Scripting.Dictionary")
        m_minX = 1000000
        m_minY = 1000000
        m_maxX = -1000000
        m_maxY = -1000000
        m_centerX = Round(tableWidth/2)
        m_centerY = Round(tableHeight/2)
        m_debug = False
    End Sub

    Public Property Let Debug(value) : m_debug = value : End Property

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
        Log  "Lights YAML File saved to: " & cGameName & "LightShows/lights-"&name&".yaml"
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
        Log "Lights File saved to: " & cGameName & "LightShows/led-grid.txt"

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
        Log "Lights File saved to: " & cGameName & "LightShows/coordsX.txt"


    End Sub

    Public Sub RegisterLights(aLights)


        Dim obj, str, ii
        For Each obj In aLights
            idx = obj.TimerInterval
            If IsArray(Lights(idx)) Then
                str = "Lights(" & idx & ") = Array("
                For Each ii In Lights(idx) : str = str & ii.Name & "," : Next
                ExecuteGlobal str & obj.Name & ")"
            ElseIf IsObject(Lights(idx)) Then
                Lights(idx) = Array(Lights(idx),obj)
            Else
                Set Lights(idx) = obj
            End If
        Next

        Dim idx,tmp,vpxLight,lcItem
    
            Dim colCount : colCount = Round(tablewidth/20)
            Dim rowCount : rowCount = Round(tableheight/20)
            Log "Registering Lights" 
            dim ledIdx : ledIdx = 0
            redim leds(UBound(Lights)-1)
            redim lightsToLeds(UBound(Lights)-1)
            ReDim ledGrid(rowCount,colCount)
            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
                If Not IsNull(vpxLight) Then
                    Log "Registering Light: "& vpxLight.name


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
                        On Error Resume Next
                        If InStr(LCase(e.Name), LCase("_" & vpxLight.Name & "_")) Or InStr(LCase(e.Name), LCase("_" & vpxLight.UserValue & "_")) Then
                            lmStr = lmStr & e.Name & ","
                        End If
                        If Err Then Log "Error: " & Err
                    Next
                    lmStr = lmStr & "Null)"
                    lmStr = Replace(lmStr, ",Null)", ")")
                    ExecuteGlobal "Dim lmArr : "&lmStr
                    m_lightmaps.Add vpxLight.Name, lmArr
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    CreateSeqRunner "lSeqRunner" & CStr(vpxLight.name), 1000
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
        CreateSeqRunner "lSeqRunner" & CStr(light.name), 1000
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
        If m_tags.Exists(tag) Then
            GetLightsForTag = m_tags(tag).Keys()
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
        If vartype(light) = 8 Then
            m_LightOn(light)
        Else
            m_LightOn(light.name)
        End If
    End Sub

    Public Sub LightOnWithColor(light, color)
        If vartype(light) = 8 Then
            m_LightOnWithColor light, color
        Else
            m_LightOnWithColor light.name, color
        End If
    End Sub

    Public Sub FadeLightToColor(light, color, fadeSpeed, runner, priority)
        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If

        If vartype(color) = 8 Then
            color = RGB( HexToInt(Left(color, 2)), HexToInt(Mid(color, 3, 2)), HexToInt(Right(color, 2)) )
        End If

        If m_lights.Exists(lightName) Then
            dim lightColor, steps
            steps = Round(fadeSpeed/20)
            If steps < 10 Then
                steps = 10
            End If
            lightColor = m_lights(lightName).Color
            Dim seq : seq  = FadeRGB(lightName, lightColor(0), color, steps)
            m_lights(lightName).Color = color
            CreateSeqRunner runner, priority
            m_seqRunners(runner).AddItem lightName & "Fade", seq, 1, 10, Null,0
            If color = RGB(0,0,0) Then
                m_lightOff(lightName)
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

        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If

        If m_lights.Exists(lightName) Then
            m_lights(lightName).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(lightName & "Blink") Then
                m_seqs(lightName & "Blink").Color = color
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
        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If
        m_lightOff(lightName)
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

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem name, seq, -1, 200/light.BlinkInterval, Null,0
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

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem light.name & "Blink", seq.Sequence, -1, 200/light.BlinkInterval, Null,0
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
                CreateSeqRunner "lSeqRunner" & group & "BlinkGroup", 1000
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
                CreateSeqRunner "lSeqRunner" & group & "BlinkGroup", 1000
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

    Public Sub CreateSeqRunner(name, priority)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        seqRunner.Priority = priority
        m_seqRunners.Add name, seqRunner
        Dim runnerKeys : runnerKeys = m_seqRunners.Keys()
        Dim itemsArray()
        ReDim itemsArray(m_seqRunners.Count - 1)
        Dim i, key
        i=0
        For Each key in runnerKeys
            Set itemsArray(i) = m_seqRunners(key)
            i=i+1
        Next
        
        Dim j, temp
        For i = 0 To UBound(itemsArray) - 1
            For j = i + 1 To UBound(itemsArray)
                If itemsArray(i).Priority > itemsArray(j).Priority Then
                    Set temp = itemsArray(i)
                    Set itemsArray(i) = itemsArray(j)
                    Set itemsArray(j) = temp
                End If
            Next
        Next
        m_seqRunners.RemoveAll
        For i = 0 To UBound(itemsArray)
            m_seqRunners.Add itemsArray(i).Name, itemsArray(i)
        Next
    End Sub

    Private Sub CreateOverrideSeqRunner(name)
        If m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqOverrideRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(runner, key, sequence, loops, speed, tokens, priority, syncMs)
        If Not m_seqRunners.Exists(runner) Then
            CreateSeqRunner runner, priority
        End If

        If syncMs = 0 Then
            If m_seqRunners(runner).Items().Exists(key) Then
                RemoveLightSeq runner, key
            End If
            m_seqRunners(runner).AddItem key, sequence, loops, speed, tokens,0
        Else
            If Not m_seqRunners.Exists(syncMs) Then
                CreateSeqRunner syncMs, 10
                m_seqRunners(syncMs).AddItem syncMs,Array("lXX|100", "lXX|0"), -1, 400/syncMs, Null, syncMs
            End If

            If Not m_syncMs.Exists(syncMs) Then
                m_syncMs.Add syncMs, CreateObject("Scripting.Dictionary")
            End If
            If m_syncMs(syncMs).Exists(runner & "_" & key) Then
                m_syncMs(syncMs).Remove runner & "_" & key
            End If
            m_syncMs(syncMs).Add runner & "_" & key, Array(runner, key, sequence, loops, speed, tokens)
        End If
    End Sub

    Public Sub RemoveLightSeq(runner, key)
        If Not m_seqRunners.Exists(runner) Then
            Exit Sub
        End If

        If m_seqRunners(runner).Items().Exists(key) Then
            Dim light
            For Each light in m_seqRunners(runner).ItemByKey(key).LightsInSeq
                If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
                End If
            Next
            m_seqRunners(runner).RemoveItem key
        End If

        Dim syncMs
        For Each syncMs in m_syncMs.Keys
            If m_syncMs(syncMs).Exists(runner & "_" & key) Then
                m_syncMs(syncMs).Remove runner & "_" & key
            End If
        Next
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
            If m_lights(light).IsRGB Then
                dim color : color = m_lights(light).Color
                If m_lightmaps.Exists(light) Then
                    For Each lm in m_lightmaps(light)
                        If not IsNull(lm) Then
                            lm.Color = color(0)
                        End If
                    Next
                End If
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
        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If
        dim syncLight : syncLight = Null
        
        If m_lights.Exists(lightName) Then
            'found a light
            Set syncLight = m_lights(lightName)
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

        'm_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
        m_frameTime = 20
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
                    If lcItem.IsRGB Then
                        color = m_on(lightKey).Color
                    Else
                        color = Null
                    End If
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, color, m_on(lightKey).Idx)
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
                    Dim pulseColor
                    If m_pulse(lightKey).IsRGB Then
                        pulseColor = m_pulse(lightKey).Color
                    Else
                        pulseColor = Null
                    End If
                    If IsNull(pulseColor) Then
                        AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), Null, m_pulse(lightKey).light.Idx)
                    Else
                        AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), pulseColor, m_pulse(lightKey).light.Idx)
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

                            If syncLight.IsRGB Then
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
                            Else
                                color = Null
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
                    LightColor m_lights(lightInSeq), m_lights(lightInSeq).BaseColor
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
            'Check if any seqs need starting via syncms
            If lcSeq.SyncMs > 0 Then
                If m_syncMs.Exists(lcSeq.SyncMs) Then
                    Dim seqToStart
                    For Each seqToStart in m_syncMs(lcSeq.SyncMs).Items
                        m_seqRunners(seqToStart(0)).AddItem seqToStart(1), seqToStart(2), seqToStart(3), seqToStart(4), seqToStart(5), 0
                        RunLightSeq(m_seqRunners(seqToStart(0)))
                    Next
                    m_syncMs(lcSeq.SyncMs).RemoveAll
                End If
            End If
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
                        
                        If ls.IsRGB Then
                            color = lcSeq.Color

                            If IsNull(color) Then
                                color = ls.Color
                            End If
                        Else
                            Log "NOT RGB:" & ls.Idx
                            color = Null
                            If Ubound(lsName) = 2 Then
                                If lsName(2) <> "" Then
                                    color = Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), Null)        
                                End If
                            End If
                        End If

                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        
                        lcSeq.SetLastLightState name, m_currentFrameState(name) 
                    End If
                Next       
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)
                    
                    If ls.IsRGB Then
                        color = lcSeq.Color

                        If IsNull(color) Then
                            color = ls.Color
                        End If
                    Else
                        color = Null
                        If Ubound(lsName) = 2 Then
                            If lsName(2) <> "" Then
                                color = Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), Null)        
                            End If
                        End If
                    End If

                    AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
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

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog "Light Controller", message
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
    
    Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y, m_baseColor, m_isRGB

        Public Property Get Idx()
            Idx=m_Idx
        End Property

        Public Property Get IsRGB()
            IsRGB = m_isRGB
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
            If CLng(color) = vbWhite Then
                m_isRGB = True
            Else
                m_isRGB = False
            End If
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
        End Sub

End Class

Class LCSeq
    
    Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates, m_palette, m_pattern, m_tokens, m_loops, m_syncMs

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
        If Not IsNull(input) Then
            Set m_tokens = input
        End If
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
                    If UBound(lightItem) = 2 Then
                        token = IsToken(lightItem(2))
                        If Not IsNull(token) Then
                            lightItem(2) = m_tokens(token)
                        End If
                    End If
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If
                    light = Join(lightItem, "|")
                    item(x) = light
                Next
                ' Update the input array with modified item array
                input(i) = item
            Else
                lightItem = Split(item, "|")
                token = IsToken(lightItem(0))
                If Not IsNull(token) Then
                    lightItem(0) = m_tokens(token)
                End If
                If UBound(lightItem) = 2 Then
                    token = IsToken(lightItem(2))
                    If Not IsNull(token) Then
                        lightItem(2) = m_tokens(token)
                    End If
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

    Public Property Get SyncMs()
        SyncMs=m_syncMs
    End Property
    
    Public Property Let SyncMs(input)
        m_syncMs = input
    End Property
    
    Public Property Get Frames()
        Frames=m_Frames
    End Property
    
    Public Property Let Frames(input)
        m_Frames = input
    End Property

    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Null
        m_updateInterval = 200
        m_repeat = False
        m_loops = 1
        m_Frames = 200
        m_pattern = Null
        m_syncMs = 0
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
    
    Private m_name, m_items, m_currentItemIdx, m_priority

    Public Property Get Name()
        Name=m_name
    End Property
    
    Public Property Let Name(input)
        m_name = input
    End Property

    Public Property Get Priority()
        Priority=m_priority
    End Property

    Public Property Let Priority(input)
        m_priority = input
    End Property

    Public Property Get Items()
        Set Items = m_items
    End Property

    Public Property Get ItemByKey(key)
        Set ItemByKey = m_items(key)
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
        m_priority = 1000
    End Sub

    Public Sub AddItem(key, sequence, loops, speed, tokens, syncMs)
        If Not IsNull(sequence) Then

            Dim lSeq : Set lSeq = New LCSeq
            lSeq.Name = key
            lSeq.Tokens = tokens
            lSeq.Sequence = sequence
            lSeq.UpdateInterval = 200/speed
            lSeq.Frames = 200/speed
            lSeq.Loops = loops
            lSeq.SyncMs = syncMs

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
                items(m_currentItemIdx).Loops = items(m_currentItemIdx).Loops - 1
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

Class GlfDebugLogFile
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
		 & LZ(Hour(CurrTime),   2) & "_" _
		 & LZ(Minute(CurrTime), 2) & "_" _
		 & LZ(Second(CurrTime), 2) & "_" _
		 & LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message)
		If glf_debugEnabled = True Then
			Dim FormattedMsg, Timestamp
			
			Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile("logs\"&Filename, 8, True)
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
    If Not glf_pinEvents.Exists(e) Then
        glf_debugLog.WriteToLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    glf_debugLog.WriteToLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        glf_debugLog.WriteToLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            GetRef(handler(0))(Array(handler(2), kwargs, e))
        Else
            glf_debugLog.WriteToLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
    Glf_EventBlocks(e).RemoveAll
End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        'glf_debugLog.WriteToLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    'glf_debugLog.WriteToLog "DispatchReplayPinEvent", e
    For Each k In glf_pinEventsOrder(e)
        'glf_debugLog.WriteToLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        kwargs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
    Next
    DispatchRelayPinEvent = kwargs
    Glf_EventBlocks(e).RemoveAll
End Function

Sub AddPinEventListener(evt, key, callbackName, priority, args)
    Dim i, inserted, tempArray
    If Not glf_pinEvents.Exists(evt) Then
        glf_pinEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_pinEvents(evt).Exists(key) Then
        glf_pinEvents(evt).Add key, Array(callbackName, priority, args)
        Sortglf_pinEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePinEventListener(evt, key)
    If glf_pinEvents.Exists(evt) Then
        If glf_pinEvents(evt).Exists(key) Then
            glf_pinEvents(evt).Remove key
            Sortglf_pinEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_pinEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_pinEventsOrder to maintain order based on priority
    If Not glf_pinEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_pinEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_pinEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_pinEventsOrder(evt)
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
        
        ' Update the glf_pinEventsOrder with the newly ordered or modified list
        glf_pinEventsOrder(evt) = tempArray
    End If
End Sub
Function GetPlayerState(key)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If playerState(glf_currentPlayer).Exists(key)  Then
        GetPlayerState = playerState(glf_currentPlayer)(key)
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

Function Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            Getglf_currentPlayerNumber = 1
        Case "PLAYER 2":
            Getglf_currentPlayerNumber = 2
        Case "PLAYER 3":
            Getglf_currentPlayerNumber = 3
        Case "PLAYER 4":
            Getglf_currentPlayerNumber = 4
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(glf_currentPlayer) Then
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
    If playerState(glf_currentPlayer).Exists(key) Then
        prevValue = playerState(glf_currentPlayer)(key)
        playerState(glf_currentPlayer).Remove key
    End If
    playerState(glf_currentPlayer).Add key, value
    
    If glf_playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue
    End If
    
    SetPlayerState = Null
End Function

Sub FirePlayerEventHandlers(evt, value, prevValue)
    If Not glf_playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = glf_playerEvents(evt)
    For Each k In glf_playerEventsOrder(evt)
        GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), Array(evt,value,prevValue)))
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, callbackName, priority, args)
    If Not glf_playerEvents.Exists(evt) Then
        glf_playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_playerEvents(evt).Exists(key) Then
        glf_playerEvents(evt).Add key, Array(callbackName, priority, args)
        Sortglf_playerEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePlayerStateEventListener(evt, key)
    If glf_playerEvents.Exists(evt) Then
        If glf_playerEvents(evt).Exists(key) Then
            glf_playerEvents(evt).Remove key
            Sortglf_playerEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_playerEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_playerEventsOrder to maintain order based on priority
    If Not glf_playerEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_playerEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_playerEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_playerEventsOrder(evt)
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
        
        ' Update the glf_playerEventsOrder with the newly ordered or modified list
        glf_playerEventsOrder(evt) = tempArray
    End If
End Sub

Sub EmitAllglf_playerEvents()
    Dim key
    For Each key in playerState(glf_currentPlayer).Keys()
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
    If glf_gameStarted = True Then
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
